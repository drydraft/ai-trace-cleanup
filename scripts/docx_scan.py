#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ai-trace-cleanup · 审计器（只读，绝不修改文档）

    python docx_scan.py 稿件.docx [--out DIR] [--layers ISP] [--preset word-cn]
                                 [--no-headers] [--include-comments] [--quiet]

产出（默认写到 `<稿件同目录>/<稿件名>.ai-trace/`）：
    report.md    人读报告：摘要 + 分层明细 + 可自动修复清单
    report.json  机读结果：docx_apply / docx_guard 的输入

退出码（越大越严重，便于 CI 卡门）：
    0 无命中 · 1 仅建议 · 2 有警告 · 3 有硬阻 · 10 执行出错
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_text import iter_paragraphs, read_docx, snippet, tokens, u  # noqa: E402
from rules_core import (LAYER_NAME, LAYER_ORDER, NO_AUTO, RULES, SEV_SUG,  # noqa: E402
                        SEV_WARN, RULE_META, scan_paragraph)
from rules_report import scan_report_layers  # noqa: E402

SEV_RANK = {SEV_SUG: 1, SEV_WARN: 2, "hard": 3}
PRESETS = {
    "word-cn": "中文 Word 正式文档：全角标点、中英/中数之间不留空格",
    "word-cn-loose": "中文 Word 宽松版：只清标点与不可见字符，中英之间的空格保留",
    "word-tech": "技术文档：中英之间保留空格（盘古之白），标点仍按中文规范",
}
DEFAULT_LAYERS = "ISP"          # 默认只跑机械层；M/B/F/L/C 需显式开（--all 一键开）


# ---------------------------------------------------------------- 扫描

def scan_parts(doc, parts, layers, preset="word-cn"):
    """对已载入的文档（可能已被改过）跑规则，不碰磁盘。apply 的定点循环靠它。"""
    all_paras, issues = [], []
    for part in parts:
        paras, _ = iter_paragraphs(doc.xml(part), part)
        all_paras += paras
        if any(l in layers for l in "ISP"):
            for para in paras:
                issues += scan_paragraph(para, preset=preset)
        issues += scan_report_layers(paras, layers="".join(
            l for l in layers if l in "MBFLC"))
    return issues, all_paras


def scan(path: str, layers=DEFAULT_LAYERS, preset="word-cn", headers=True,
         include_comments=False) -> dict:
    """返回 {doc, paras, issues, stats}。doc/paras 供 apply 与 guard 复用。"""
    doc = read_docx(path)
    parts = [p for p in doc.text_parts]
    if not headers:
        parts = [p for p in parts if not p.startswith(("word/header", "word/footer"))]
    if include_comments and "word/comments.xml" in doc.data:
        parts.append("word/comments.xml")

    issues, all_paras = scan_parts(doc, parts, layers, preset)
    issues.sort(key=lambda i: (LAYER_ORDER.index(i.layer), -SEV_RANK.get(i.severity, 0),
                               i.part, i.pid, i.start))
    stats = {
        "file": os.path.abspath(path),
        "scanned_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "preset": preset,
        "layers": layers,
        "parts": parts,
        "paragraphs": len(all_paras),
        "chars": sum(len(p.body()) for p in all_paras),
        "words": len(tokens(" ".join(p.visible_text for p in all_paras))),
        "issues": len(issues),
        "auto": sum(1 for i in issues if i.auto),
        "hard": sum(1 for i in issues if i.severity == "hard"),
        "warn": sum(1 for i in issues if i.severity == SEV_WARN),
        "sug": sum(1 for i in issues if i.severity == SEV_SUG),
    }
    para_map = {(p.part, p.pid): p for p in all_paras}
    return {"doc": doc, "paras": all_paras, "para_map": para_map,
            "issues": issues, "stats": stats}


# ---------------------------------------------------------------- 报告

def render_md(res: dict, max_detail=200) -> str:
    st, issues = res["stats"], res["issues"]
    lines = []
    a = lines.append
    a("# AI 痕迹审计报告")
    a("")
    a("| 项 | 值 |")
    a("| --- | --- |")
    a("| 文件 | `%s` |" % st["file"])
    a("| 扫描时间 | %s |" % st["scanned_at"])
    a("| 预设 | `%s` — %s |" % (st["preset"], PRESETS.get(st["preset"], "")))
    a("| 扫描层 | `%s` |" % st["layers"])
    a("| 范围 | %d 段 / %d 字；部件：%s |"
      % (st["paragraphs"], st["chars"], ", ".join("`%s`" % p for p in st["parts"])))
    a("")

    a("## 一、摘要")
    a("")
    a("| 层 | 说明 | 硬阻 ⛔ | 警告 ⚠️ | 建议 💡 | 其中可自动 |")
    a("| --- | --- | --- | --- | --- | --- |")
    for layer in LAYER_ORDER:
        buf = [i for i in issues if i.layer == layer]
        if not buf:
            continue
        a("| %s | %s | %d | %d | %d | %d |" % (
            layer, LAYER_NAME[layer],
            sum(1 for i in buf if i.severity == "hard"),
            sum(1 for i in buf if i.severity == SEV_WARN),
            sum(1 for i in buf if i.severity == SEV_SUG),
            sum(1 for i in buf if i.auto)))
    a("| **合计** | | **%d** | **%d** | **%d** | **%d** |"
      % (st["hard"], st["warn"], st["sug"], st["auto"]))
    a("")

    autos = [i for i in issues if i.auto]
    a("## 二、可自动修复清单（%d 条）" % len(autos))
    a("")
    if not autos:
        a("无。")
    else:
        a("这些是纯字符级替换，不改任何句意。执行：")
        a("")
        a("```")
        a("python scripts/docx_apply.py \"%s\" --report \"report.json\" --apply"
          % os.path.basename(st["file"]))
        a("```")
        a("")
        by_rule = {}
        for i in autos:
            by_rule.setdefault(i.rule, []).append(i)
        a("| 规则 | 条数 | 说明 |")
        a("| --- | --- | --- |")
        for rid in sorted(by_rule, key=lambda r: (LAYER_ORDER.index(RULE_META[r]["layer"]), r)):
            m = RULE_META[rid]
            a("| `%s` | %d | %s → %s |" % (rid, len(by_rule[rid]), m["title"], m["desc"]))
        a("")

    a("## 三、明细")
    a("")
    shown = 0
    for layer in LAYER_ORDER:
        buf = [i for i in issues if i.layer == layer]
        if not buf:
            continue
        a("### [%s] %s（%d 条）" % (layer, LAYER_NAME[layer], len(buf)))
        a("")
        for i in buf:
            if shown >= max_detail:
                a("> 明细已截断（--max-detail 可放大）。完整清单见 `report.json`。")
                a("")
                break
            shown += 1
            loc = "`%s` §%s" % (_short_part(i.part), i.pid)
            if i.start >= 0 and (i.part, i.pid) in res["para_map"]:
                ctx = snippet(res["para_map"][(i.part, i.pid)], i.start, i.end)
            else:
                ctx = i.anchor
            a("- %s **%s** %s · %s" % (i.mark, i.rule, loc, ctx))
            a("  - %s%s" % (i.note, ("　→ 建议改成 `%s`" % i.repl) if i.repl else ""))
        a("")
    if not issues:
        a("**未发现问题。**")
        a("")
    a("---")
    a("")
    a("> 本报告只做建议，不改文档。机械层（I/S/P）可用 `docx_apply.py` 落盘；"
      "M/B/F/L/C 层必须人工处理。")
    return "\n".join(lines)


def _short_part(part: str) -> str:
    return re.sub(r"^word/", "", part)


def rule_catalog_md() -> str:
    """按代码里的规则表生成 markdown 目录 —— 文档不会再跟代码跑偏。"""
    from rules_report import REPORT_RULE_META, CAT_RULE, load_lexicon
    lex = load_lexicon()
    L = ["<!-- 本文件由 `python scripts/docx_scan.py --list-rules` 生成，不要手改。 -->", "",
         "# 规则目录", "",
         "## 机械层（可落盘）", "",
         "| 规则 | 层 | 强度 | 名称 | 判定 | 默认自动落盘 |",
         "| --- | --- | --- | --- | --- | --- |"]
    for rid, layer, sev, title, desc, _fn in RULES:
        L.append("| `%s` | %s | %s | %s | %s | %s |"
                 % (rid, layer, _sev_word(sev), title, desc,
                    "否（需显式指定）" if rid in NO_AUTO else "是"))
    L += ["", "## 报告层（只出建议，永不落盘）", "",
          "| 规则 | 层 | 强度 | 名称 | 判定 |", "| --- | --- | --- | --- | --- |"]
    from rules_report import M_PATTERNS, L_STRUCT, C_PATTERNS, REPORT_RULE_META as META
    rows, seen = [], set()
    for rid, sev, title, _pat, desc in M_PATTERNS:
        rows.append((rid, "M", sev, title, desc))
    for rid in META:
        layer, sev, title, desc = META[rid]
        rows.append((rid, layer, sev, title, desc))
    for rid, sev, title, _pat, desc in L_STRUCT:
        rows.append((rid, "L", sev, title, desc))
    for rid, sev, title, _pat, desc in C_PATTERNS:
        rows.append((rid, "C", sev, title, desc))
    for rid, layer, sev, title, desc in sorted(
            rows, key=lambda r: (LAYER_ORDER.index(r[1]), r[0])):
        if rid in seen:
            continue
        seen.add(rid)
        L.append("| `%s` | %s | %s | %s | %s |" % (rid, layer, _sev_word(sev), title, desc))
    for cat, (rid, title, sev) in sorted(CAT_RULE.items(), key=lambda kv: kv[1][0]):
        terms = "、".join(t for t, _ in lex.get(cat, []))
        L.append("| `%s` | L | %s | %s | 词表命中（%d 条）：%s |"
                 % (rid, _sev_word(sev), title, len(lex.get(cat, [])), terms))
    return "\n".join(L)


def _sev_word(sev: str) -> str:
    return {"hard": "⛔ 硬阻", "warn": "⚠️ 警告"}.get(sev, "💡 建议")


# ---------------------------------------------------------------- CLI

def main(argv=None) -> int:
    u()
    ap = argparse.ArgumentParser(description="AI 痕迹审计（只读）")
    ap.add_argument("file", nargs="?", help="要审计的 .docx；只有 --list-rules 时可省略")
    ap.add_argument("--out", default=None, help="报告输出目录")
    ap.add_argument("--layers", default=None, help="要跑的层，如 I,S,P 或 ISP；--all 开全部")
    ap.add_argument("--all", action="store_true", help="等同于 --layers ISPBFMLC")
    ap.add_argument("--preset", default="word-cn", choices=sorted(PRESETS))
    ap.add_argument("--no-headers", action="store_true", help="跳过页眉页脚")
    ap.add_argument("--include-comments", action="store_true", help="连批注一起扫")
    ap.add_argument("--max-detail", type=int, default=200)
    ap.add_argument("--json-only", action="store_true", help="只写 report.json")
    ap.add_argument("--list-rules", action="store_true",
                    help="只打印规则目录（markdown），用于生成 references/rule-index.md")
    ap.add_argument("--rules-out", default=None,
                    help="把规则目录直接以 UTF-8 写入该路径（不要用 shell 重定向，会按 GBK 弄坏中文）")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    if args.list_rules or args.rules_out:
        md = rule_catalog_md()
        if args.rules_out:
            with open(args.rules_out, "w", encoding="utf-8") as fh:
                fh.write(md + "\n")
            print("已写入 %s" % args.rules_out)
        else:
            print(md)
        return 0

    layers = args.layers or ("ISPBMFLC" if args.all else DEFAULT_LAYERS)
    layers = "".join(dict.fromkeys(l for l in layers.replace(",", "").upper() if l in LAYER_ORDER))
    if not args.file:
        print("缺参数：要审计的 .docx（或用 --list-rules 看规则目录）")
        return 10
    if not layers:
        print("--layers 里没有有效层（可选 %s）" % "/".join(LAYER_ORDER))
        return 10
    if not os.path.isfile(args.file):
        print("找不到文件：%s" % args.file)
        return 10

    res = scan(args.file, layers=layers, preset=args.preset,
               headers=not args.no_headers, include_comments=args.include_comments)
    outdir = args.out or os.path.join(os.path.dirname(os.path.abspath(args.file)),
                                      os.path.splitext(os.path.basename(args.file))[0] + ".ai-trace")
    os.makedirs(outdir, exist_ok=True)
    st = res["stats"]
    payload = dict(stats=st, issues=[i.as_dict() for i in res["issues"]])
    with open(os.path.join(outdir, "report.json"), "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)
    if not args.json_only:
        with open(os.path.join(outdir, "report.md"), "w", encoding="utf-8") as fh:
            fh.write(render_md(res, args.max_detail))

    if not args.quiet:
        print("文件：%s" % args.file)
        print("范围：%d 段 / %d 字 / %d 个语义 token" % (st["paragraphs"], st["chars"], st["words"]))
        print("命中：⛔%d  ⚠️%d  💡%d   （其中可自动修复 %d）"
              % (st["hard"], st["warn"], st["sug"], st["auto"]))
        for layer in LAYER_ORDER:
            n = sum(1 for i in res["issues"] if i.layer == layer)
            if n:
                print("  [%s] %-10s %d" % (layer, LAYER_NAME[layer], n))
        print("报告：%s" % outdir)

    return max([SEV_RANK.get(i.severity, 0) for i in res["issues"]] or [0])


if __name__ == "__main__":
    sys.exit(main())
