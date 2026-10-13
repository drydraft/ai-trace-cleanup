#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ai-trace-cleanup · 自检（回归闸门）

    python tests/selftest.py

它做的事，就是「这套规则还能不能用」的全部证明：
  T1 脏样本能生成，且 50 条规则**条条命中**（加了规则却没加样本 → 直接失败）
  T2 scan → apply → guard 全链路退出码为 0
  T3 关键替换逐条比对（跨 run 删空格这种核心能力，必须逐字对）
  T4 清理后文档能被 python-docx 独立打开（第三方解析器交叉验证）
  T5 清理后机械层零残留（幂等）
  T6 保护区逐字节未变、非文本部件 CRC 未变
  T7 二次 apply 无事可做（真幂等，不只是"扫不到"）
"""

from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "scripts"))

from docx_text import read_docx, iter_paragraphs, u  # noqa: E402
from rules_core import RULES  # noqa: E402
from rules_report import (CAT_RULE, C_PATTERNS, L_STRUCT, M_PATTERNS,  # noqa: E402
                          load_lexicon)

PY = sys.executable
SAMPLE = os.path.join(HERE, "sample-dirty.docx")
CLEAN = os.path.join(HERE, "clean.docx")


def run(*args) -> tuple:
    p = subprocess.run([PY] + list(args), capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=ROOT)
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def rule_spec() -> list:
    from rules_report import REPORT_RULE_META
    lcats = [CAT_RULE[c][0] for c in load_lexicon() if c in CAT_RULE]
    return sorted(set(
        [r[0] for r in RULES] + [m[0] for m in M_PATTERNS]
        + [s[0] for s in L_STRUCT] + list(REPORT_RULE_META)
        + lcats + [c[0] for c in C_PATTERNS]))


def paras(path):
    doc = read_docx(path)
    out = {}
    for part in doc.text_parts:
        for p in iter_paragraphs(doc.xml(part), part)[0]:
            out[(part, p.text.strip())] = p
    return doc, out


#: 关键替换的期望结果（改规则时这里必须跟着改，改不了说明规则坏了）
EXPECT = {
    "本次巡检共 发现 3 个问题，已闭环 2 项。": "本次巡检共发现3个问题，已闭环2项。",
    "2026 年第三季度巡检工作总结": "2026年第三季度巡检工作总结",
    "检查完成, 未发现异常.": "检查完成，未发现异常。",
    "省略号... 还有。。。两种写法": "省略号……还有……两种写法",
    "发现了严重问题！！为什么呢？？": "发现了严重问题！为什么呢？",
    "占比 50 ％，全角数字１２３": "占比50％，全角数字123",
    "括号内空格（ 说明文字 ）与尾随空格 ": "括号内空格（说明文字）与尾随空格",
    "共 3 个问题": "共3个问题",
    "已处理 2 个,剩余 1 个": "已处理2个，剩余1个",
    "文本框里也有 3 个空格问题": "文本框里也有3个空格问题",
    "巡检工作总结 共 5 页": "巡检工作总结共5页",
}


def main() -> int:
    u()
    fails, notes = [], []

    def check(ok, label, detail=""):
        print("%s %s%s" % ("✅" if ok else "❌", label, ("　" + detail) if detail else ""))
        if not ok:
            fails.append(label + (("　" + detail) if detail else ""))
        return ok

    print("── T1 生成脏样本并检查规则覆盖率 ──")
    rc, out = run(os.path.join(HERE, "make_sample.py"), SAMPLE)
    check(rc == 0, "脏样本生成", out.strip().splitlines()[-1] if out.strip() else "")
    rc, out = run(os.path.join(ROOT, "scripts", "docx_scan.py"), SAMPLE, "--all",
                  "--out", os.path.join(HERE, "scan-all"), "--quiet")
    check(rc == 3, "审计能报出硬阻（退出码 3）", "实际 %d" % rc)
    res = json.load(open(os.path.join(HERE, "scan-all", "report.json"), encoding="utf-8"))
    hit = {i["rule"] for i in res["issues"]}
    spec = rule_spec()
    uncovered = [r for r in spec if r not in hit]
    check(not uncovered, "每条规则都有样本命中（%d 条）" % len(spec),
          "未命中：%s" % uncovered)

    print("\n── T2 全链路：scan → apply → guard ──")
    for stale in (CLEAN, CLEAN + ".applied.json"):
        if os.path.isfile(stale):
            os.remove(stale)
    rc, out = run(os.path.join(ROOT, "scripts", "docx_apply.py"), SAMPLE,
                  "--out", CLEAN, "--apply")
    check(rc == 0, "apply 落盘成功（校验通过才写）", out.strip().splitlines()[0] if out.strip() else "")
    rc, out = run(os.path.join(ROOT, "scripts", "docx_guard.py"), SAMPLE, CLEAN,
                  "--applied", CLEAN + ".applied.json")
    check(rc == 0, "guard 七项校验全过")
    for line in out.splitlines():
        if line.strip().startswith(("✅", "❌", "⏭")):
            notes.append("     " + line.strip())

    print("\n── T3 关键替换逐条比对 ──")
    _, after = paras(CLEAN)
    for before, want in EXPECT.items():
        got = next((k[1] for k in after if k[1] == want), None)
        check(got is not None, "「%s」→「%s」" % (before[:18], want[:18]))

    print("\n── T4 第三方解析器交叉验证 ──")
    try:
        import docx as _docx
        d = _docx.Document(CLEAN)
        check(len(d.paragraphs) > 50 and len(d.tables) == 1,
              "python-docx 能独立打开", "段落 %d / 表格 %d" % (len(d.paragraphs), len(d.tables)))
    except ImportError:
        print("⏭️ 未装 python-docx，跳过（不影响结论）")

    print("\n── T5 幂等 ──")
    rc, out = run(os.path.join(ROOT, "scripts", "docx_scan.py"), CLEAN, "--layers", "ISP",
                  "--out", os.path.join(HERE, "scan-clean"), "--quiet")
    res2 = json.load(open(os.path.join(HERE, "scan-clean", "report.json"), encoding="utf-8"))
    left = [i for i in res2["issues"] if i["auto"]]
    check(not left, "清理后机械层零残留", "残留 %d 处" % len(left))

    print("\n── T6 保护区 ──")
    a, b = read_docx(SAMPLE), read_docx(CLEAN)
    changed = [n for n in a.order if not a.is_text_part(n) and a.data[n] != b.data[n]]
    check(not changed, "非文本部件逐字节未变", ",".join(changed) or "全部 %d 个部件" %
          len(a.protected_parts()))
    xml_changed = [n for n in a.order if a.is_text_part(n) and a.data[n] != b.data[n]]
    check(xml_changed, "文本部件确实变了", "、".join(xml_changed))
    # 只允许 <w:t> 的 body 变：把两边的 w:t 内容全挖掉后，XML 应完全一致
    def skeleton(buf: bytes) -> bytes:
        import re
        s = buf.decode("utf-8")
        return re.sub(r"(?<=>)[^<]*(?=</w:t>)", "", s).encode("utf-8")
    same = all(skeleton(a.data[n]) == skeleton(b.data[n]) for n in xml_changed)
    check(same, "XML 骨架未变（只有 <w:t> 里的字变了）")

    print("\n── T7 二次 apply 无事可做 ──")
    rc, out = run(os.path.join(ROOT, "scripts", "docx_apply.py"), CLEAN,
                  "--out", os.path.join(HERE, "clean2.docx"))
    check("无需改动" in out, "第二次 apply 报「无需改动」")

    if notes:
        print("\n── guard 明细 ──")
        print("\n".join(notes))

    print("\n" + ("=" * 56))
    if fails:
        print("❌ 自检失败 %d 项：" % len(fails))
        for f in fails:
            print("   · " + f)
        return 1
    print("✅ 自检全部通过（%d 条规则、%d 处命中、%d 处自动替换）"
          % (len(spec), len(res["issues"]), res["stats"]["auto"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
