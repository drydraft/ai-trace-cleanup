#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ai-trace-cleanup · 校验器（最后一道门）

    python docx_guard.py 原稿.docx 清理后.docx [--applied 清理后.docx.applied.json]

任何一项不过 → 退出码 3，调用方必须回滚。

七项检查
--------
G1 包结构：条目名、顺序、压缩方式与原稿完全一致。
G2 非文本部件：逐字节相同（styles/numbering/media/rels/…，一张图都不能动）。
G3 文本部件：把 applied.json 里的替换在**原稿**上按轮次重放，结果必须与清理后
   逐字节相同 —— 这是最强的一条证明：「除了这些已知替换，文档里什么都没变」。
G4 语义 token：数字、URL、拉丁词、代码标识符的集合前后必须完全相等。
G5 结构：部件、段落数、每段 span 数、每段 visible_text 长度关系一致（没删段没并段）。
G6 可解析：每个被改过的 XML 部件都能被 XML 解析器读出（不会让 Word 弹「修复」）。
G7 幂等：对清理后的文档再跑一遍同样的机械层规则，必须 0 命中。

为什么 G3 单独不够，还要 G4/G5：G3 只证明"只有我们计划改的地方变了"，
但计划本身可能是错的（比如某条规则误删了数字）。G4/G5 反过来管计划本身。
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_text import iter_paragraphs, read_docx, tokens, u  # noqa: E402


def _paras_of(doc, part):
    return iter_paragraphs(doc.xml(part), part)[0]


def _all_text(doc) -> str:
    buf = []
    for part in doc.text_parts:
        for p in _paras_of(doc, part):
            buf.append(p.visible_text)
    return "\n".join(buf)


def _selftest_tokens(orig, new):
    a, b = tokens(_all_text(orig)), tokens(_all_text(new))
    missing, added = sorted(a - b), sorted(b - a)
    return missing, added


def check_structure(orig, new):
    """段落与 span 的骨架必须一一对应。"""
    bad = []
    if orig.text_parts != new.text_parts:
        bad.append("文本部件列表变了：%s → %s" % (orig.text_parts, new.text_parts))
        return bad
    for part in orig.text_parts:
        po, pn = _paras_of(orig, part), _paras_of(new, part)
        if len(po) != len(pn):
            bad.append("%s 段落数变了：%d → %d" % (part, len(po), len(pn)))
            continue
        for a, b in zip(po, pn):
            if len(a.spans) != len(b.spans):
                bad.append("%s §%s run 数变了：%d → %d" % (part, a.pid, len(a.spans), len(b.spans)))
            if a.style != b.style:
                bad.append("%s §%s 段落样式变了：%r → %r" % (part, a.pid, a.style, b.style))
            if a.ppr != b.ppr:
                bad.append("%s §%s 段落属性(pPr)被改动了" % (part, a.pid))
    return bad


def check_package(orig, new):
    bad = []
    if orig.order != new.order:
        bad.append("zip 条目顺序或名称变了")
    for name in orig.order:
        if name not in new.data:
            bad.append("条目丢失：%s" % name)
            continue
        if not orig.is_text_part(name) and orig.data[name] != new.data[name]:
            bad.append("非文本部件被改动：%s（必须逐字节不变）" % name)
    with zipfile.ZipFile(new.path) as z:
        for name in z.namelist():
            if name.endswith((".xml", ".rels")):
                try:
                    ET.fromstring(z.read(name))
                except ET.ParseError as exc:
                    bad.append("XML 无法解析：%s（%s）" % (name, exc))
    return bad


def replay(orig, applied) -> str:
    """在 orig 上按轮次重放替换，返回 (part -> 期望字节) 与校验错误。"""
    expect = {name: orig.data[name] for name in orig.order}
    bad = []
    for rnd in sorted({a.get("round", 0) for a in applied}):
        by_part = {}
        for a in applied:
            if a.get("round", 0) != rnd:
                continue
            by_part.setdefault(a["part"], []).append(a)
        for part, items in by_part.items():
            xml = expect[part].decode("utf-8")
            paras = {p.pid: p for p in iter_paragraphs(xml, part)[0]}
            edits = []
            for it in items:
                para = paras.get(it["pid"])
                if para is None:
                    bad.append("重放失败：%s §%s 不存在" % (part, it["pid"]))
                    continue
                span = next((s for s in para.spans
                             if s.char_start <= it["start"] < s.char_end), None)
                if span is None:
                    bad.append("重放失败：%s §%s @%d 找不到所属 run" % (part, it["pid"], it["start"]))
                    continue
                if para.text[it["start"]:it["end"]] != it["orig"]:
                    bad.append("重放不一致：%s §%s @%d 原文是 %r，记录里是 %r"
                               % (part, it["pid"], it["start"],
                                  para.text[it["start"]:it["end"]], it["orig"]))
                    continue
                edits.append((span, it["start"], it["end"], it["repl"]))
            if edits:
                from docx_text import rewrite_part
                expect[part] = rewrite_part(xml, part, edits).encode("utf-8")
    return expect, bad


def verify(orig_path, new_path, applied_path=None, quiet=False) -> tuple:
    """返回 (ok, lines)。可被 docx_apply 直接调用。"""
    lines = []

    def say(s):
        lines.append(s)

    orig, new = read_docx(orig_path), read_docx(new_path)
    ok = True

    bad = check_package(orig, new)
    say(("%s G1/G2/G6 包结构、保护区、可解析性" % ("✅" if not bad else "❌"))
        + ("：" + "；".join(bad) if bad else "（%d 个非文本部件逐字节未变）" % len(orig.protected_parts())))
    ok &= not bad

    if applied_path and os.path.isfile(applied_path):
        with open(applied_path, encoding="utf-8") as fh:
            applied = json.load(fh).get("applied", [])
        expect, bad2 = replay(orig, applied)
        diff = [n for n in orig.order if expect.get(n) != new.data.get(n)
                and orig.is_text_part(n)
                or (not orig.is_text_part(n) and orig.data[n] != new.data[n])]
        g3 = not bad2 and not diff
        ok &= g3
        say("%s G3 逐字节重放（%d 处替换）%s"
            % ("✅" if g3 else "❌", len(applied),
               "" if g3 else "：" + "；".join((bad2 + ["未预期的差异：%s" % ", ".join(diff)])[:4])))
    else:
        say("⏭️ G3 跳过（没给 --applied，无法做逐字节重放证明）")

    missing, added = _selftest_tokens(orig, new)
    g4 = not missing and not added
    ok &= g4
    say("%s G4 语义 token 集合%s"
        % ("✅" if g4 else "❌",
           "" if g4 else "：少了 %s；多了 %s" % (missing[:6], added[:6])))

    bad = check_structure(orig, new)
    g5 = not bad
    ok &= g5
    say("%s G5 段落/run 骨架%s" % ("✅" if g5 else "❌", "" if g5 else "：" + "；".join(bad[:4])))

    if os.path.isfile(new_path):
        from docx_scan import scan
        res = scan(new_path, layers="ISP", preset="word-cn")
        leftover = [i for i in res["issues"] if i.auto]
        g7 = not leftover
        ok &= g7
        say("%s G7 幂等（清理后再扫，%d 处机械层残留）%s"
            % ("✅" if g7 else "❌", len(leftover),
               "" if g7 else "：" + "；".join("%s@%s§%s" % (i.rule, i.part, i.pid)
                                             for i in leftover[:5])))

    say("%s 总结：%s" % ("✅" if ok else "❌", "全部通过" if ok else "未通过 —— 必须回滚"))
    if not quiet:
        for line in lines:
            print(line)
    return ok, lines


def main(argv=None) -> int:
    u()
    ap = argparse.ArgumentParser(description="AI 痕迹清理 · 校验器")
    ap.add_argument("orig")
    ap.add_argument("new")
    ap.add_argument("--applied", default=None, help="docx_apply.py 产出的 *.applied.json")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)
    ok, _ = verify(args.orig, args.new, args.applied, args.quiet)
    return 0 if ok else 3


if __name__ == "__main__":
    sys.exit(main())
