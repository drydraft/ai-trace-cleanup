#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ai-trace-cleanup · 机械层应用器（默认 dry-run）

    # 1) 先审计，拿到清单
    python docx_scan.py 稿件.docx
    # 2) 看一眼要改什么（不改文件）
    python docx_apply.py 稿件.docx --report 稿件.ai-trace/report.json
    # 3) 确认无误再落盘
    python docx_apply.py 稿件.docx --report 稿件.ai-trace/report.json --apply

只落 **I / S / P** 三层的字符级替换，而且：
  · 默认不改文件；`--apply` 才写，写前自动备份 `.bak`；
  · 落盘前先在内存里跑到**定点**（一条规则改出来的新串可能触发另一条规则）；
  · 写完立刻调 docx_guard **逐字节重放校验**，不过就整个丢弃临时文件，原稿分毫未动。

M/B/F/L/C 五层一律不落盘 —— 那几层要动版式或语义，必须人来做。
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_scan import scan_parts  # noqa: E402
from docx_text import read_docx, u, write_docx  # noqa: E402
from rules_core import NO_AUTO, apply_edits  # noqa: E402

MAX_ROUNDS = 8


def build(src: str, allowed=None, preset="word-cn", headers=True, quiet=False) -> tuple:
    """在内存里迭代到定点，返回 (doc, applied, converged)。不写任何文件。"""
    doc = read_docx(src)
    parts = [p for p in doc.text_parts]
    if not headers:
        parts = [p for p in parts if not p.startswith(("word/header", "word/footer"))]

    applied, seen_states, converged = [], set(), False
    for rnd in range(MAX_ROUNDS):
        state = hash(tuple(doc.data[p] for p in parts))
        if state in seen_states:
            print("⚠️ 规则出现往复（第 %d 轮回到已见过的状态），已中止以保安全" % rnd)
            break
        seen_states.add(state)
        issues, _ = scan_parts(doc, parts, "ISP", preset)
        todo = [i for i in issues if i.auto and (allowed is None or i.rule in allowed)]
        if not todo:
            converged = True
            break
        for part in parts:
            new_xml = apply_edits(doc.xml(part), part, todo)
            if new_xml != doc.xml(part):
                doc.set_xml(part, new_xml)
        for i in todo:
            applied.append(dict(round=rnd, rule=i.rule, part=i.part, pid=i.pid,
                                start=i.start, end=i.end, orig=i.orig, repl=i.repl))
    if not quiet:
        by_rule = {}
        for a in applied:
            by_rule[a["rule"]] = by_rule.get(a["rule"], 0) + 1
        print("待替换 %d 处，共 %d 轮：%s"
              % (len(applied), len({a["round"] for a in applied}) or 0,
                 "、".join("%s×%d" % kv for kv in sorted(by_rule.items())) or "无"))
    return doc, applied, converged


def main(argv=None) -> int:
    u()
    ap = argparse.ArgumentParser(description="机械层替换（默认 dry-run）")
    ap.add_argument("file")
    ap.add_argument("--report", default=None, help="docx_scan.py 产出的 report.json（用于确认清单）")
    ap.add_argument("--rules", default=None, help="只应用这些规则，如 S01,S03,P01")
    ap.add_argument("--exclude", default=None, help="排除这些规则")
    ap.add_argument("--preset", default="word-cn")
    ap.add_argument("--out", default=None, help="输出文件；默认写回原文件（原文件另存 .bak）")
    ap.add_argument("--backup", default=True, action="store_true", help="覆盖前备份 .bak（默认开）")
    ap.add_argument("--no-backup", dest="backup", action="store_false")
    ap.add_argument("--no-headers", action="store_true")
    ap.add_argument("--apply", action="store_true", help="真的写盘（不加就是 dry-run）")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)
    if not os.path.isfile(args.file):
        print("找不到文件：%s" % args.file)
        return 10

    allowed = None
    if args.rules:
        allowed = {r.strip().upper() for r in args.rules.split(",") if r.strip()}
        blocked = allowed & NO_AUTO
        if blocked:
            print("⛔ 这些规则被标为「不自动」，拒绝了：%s" % "、".join(sorted(blocked)))
            print("   它们误伤风险高（如 P02 直引号、P04 破折号），改完必须人过一眼。")
            return 10
        denied = {r.strip().upper() for r in (args.exclude or "").split(",") if r.strip()}
        allowed -= denied
    if args.exclude and allowed is None:
        allowed = None

    doc, applied, converged = build(args.file, allowed=allowed, preset=args.preset,
                                    headers=not args.no_headers, quiet=args.quiet)
    if not converged and not args.quiet:
        print("⚠️ 未在 %d 轮内收敛，请人工检查（可能有一条规则反复改同一处）" % MAX_ROUNDS)
    if not applied:
        print("无需改动。")
        return 0

    if not args.apply:
        print("\n（dry-run，未写文件）前 40 处预览：")
        for a in applied[:40]:
            print("  %-4s %s §%s  %r → %r" % (a["rule"], a["part"], a["pid"], a["orig"], a["repl"]))
        if len(applied) > 40:
            print("  … 还有 %d 处" % (len(applied) - 40))
        print("\n加 --apply 落盘。落盘后会自动跑校验，不过就丢弃。")
        return 0

    # ---- 落盘：先写临时文件 → 校验 → 通过才替换 ----
    out = args.out or args.file
    tmp = out + ".tmp-ai-trace"
    write_docx(doc, tmp)
    ap_path = out + ".applied.json"
    with open(ap_path, "w", encoding="utf-8") as fh:
        json.dump({"source": os.path.abspath(args.file), "preset": args.preset,
                   "note": "只含 I/S/P 三层的字符级替换；每轮坐标相对上一轮结果",
                   "applied": applied}, fh, ensure_ascii=False, indent=1)

    from docx_guard import verify
    ok, lines = verify(args.file, tmp, ap_path, quiet=True)
    if not ok:
        os.remove(tmp)
        print("⛔ 校验未通过，已丢弃改动，原文件未被动过。")
        for line in lines:
            print("   " + line)
        return 3

    if args.backup and os.path.abspath(out) == os.path.abspath(args.file):
        shutil.copy2(args.file, args.file + ".bak")
    shutil.move(tmp, out)
    print("✅ 已写入 %s（替换 %d 处）" % (out, len(applied)))
    if args.backup and os.path.abspath(out) == os.path.abspath(args.file):
        print("   原稿备份：%s" % (args.file + ".bak"))
    print("   替换清单：%s" % ap_path)
    for line in lines:
        print("   " + line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
