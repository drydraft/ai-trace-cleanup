#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ai-trace-cleanup · 真 Word 往返验证（需要本机装了 Word）

    python tests/word_roundtrip.py

跟 `selftest.py` 的区别：selftest 用的是**手工拼的最小 docx**，
这里用的是 **Word 自己写出来的真 docx**（十几二十个部件、真样式表、真关系表）。

验的事：
  W1 Word 能生成真文档（用 InsertFile 灌入脏文本）
  W2 审计 → 清理 → 校验 全链路 0 退出码；受保护部件数远多于最小样本
  W3 Word 能把清理后的文档完整读回来，且读回来的文字正是我们想改成的样子
  W4 Word 读回的文本里，那些脏写法确实没了

任一步失败即退出码 1。没装 Word 会自动跳过（退出码 0），不影响别的测试。
"""

from __future__ import annotations

import io
import json
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "scripts"))

PY = sys.executable
PS = shutil.which("powershell") or shutil.which("pwsh")

#: 灌进 Word 的脏文本：每行一个段落，覆盖各层最典型的痕迹
DIRTY_LINES = [
    "2026 年第三季度巡检工作总结",
    "本次巡检共 发现 3 个问题，已闭环 2 项。",
    "**重点**：以下为问题清单说明",
    "## 二、主要问题",
    "- 设备老化，风扇异响",
    "1. 机房温湿度超标",
    "检查完成, 未发现异常.",
    "所谓\"闭环\"在这里被滥用",
    "省略号... 还有。。。两种写法",
    "发现了严重问题！！为什么呢？？",
    "占比 50 ％，全角数字１２３",
    "零宽字符\u200b夹在中间，不换行空格\u00a0也来了",
    "括号内空格（ 说明文字 ）与尾随空格 ",
    "共 发现 六个汉字中间的空格",
    "首先，我们需要明确目标。其次，要拆解任务。最后，形成闭环。",
    "综上所述，我们可以得出结论。",
    "作为一个 AI 语言模型，我无法直接访问机房。",
    "以上内容仅供参考，具体请以现场实际情况为准。",
]
#: 清理后必须出现的样子（不能出现「脏」半句）
EXPECT_CLEAN = [
    "2026年第三季度巡检工作总结",
    "本次巡检共发现3个问题，已闭环2项。",
    "检查完成，未发现异常。",
    "省略号……还有……两种写法",
    "发现了严重问题！为什么呢？",
    "占比50％，全角数字123",
    "零宽字符夹在中间，不换行空格也来了",
    "括号内空格（说明文字）与尾随空格",
    "共发现六个汉字中间的空格",
]
EXPECT_GONE = ["2026 年", "共 发现", "检查完成,", "...", "！！", "１２３", "\u200b", "\u00a0"]


def ps(script, *args):
    return subprocess.run([PS, "-NoProfile", "-ExecutionPolicy", "Bypass",
                           "-File", os.path.join(HERE, script)] + list(args),
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace")


def run(*args):
    p = subprocess.run([PY] + list(args), capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=ROOT)
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def main() -> int:
    from docx_text import u, read_docx
    u()
    if not PS:
        print("⏭️ 找不到 powershell，跳过真 Word 验证")
        return 0

    src_txt = os.path.join(HERE, "word-input.txt")
    real = os.path.join(HERE, "word-real.docx")
    clean = os.path.join(HERE, "word-clean.docx")
    out_txt = os.path.join(HERE, "word-text-out.txt")
    for f in (real, clean, out_txt, clean + ".applied.json"):
        if os.path.isfile(f):
            os.remove(f)
    with open(src_txt, "w", encoding="utf-8") as fh:
        fh.write("\n".join(DIRTY_LINES) + "\n")

    print("── W1 用真 Word 生成文档 ──")
    p = ps("word_make.ps1", src_txt, real)
    if p.returncode != 0 or not os.path.isfile(real):
        print("⏭️ Word 生成失败（可能未安装或被占用），跳过：%s" % (p.stderr or p.stdout).strip()[:200])
        return 0
    doc = read_docx(real)
    prot = doc.protected_parts()
    print("✅ Word 生成了 %s" % os.path.basename(real))
    print("   部件 %d 个，其中受保护（非文本）%d 个：%s"
          % (len(doc.order), len(prot), ", ".join(p for p in prot if p.endswith(".xml"))[:180]))

    print("\n── W2 审计 → 清理 → 校验 ──")
    rc, out = run(os.path.join(ROOT, "scripts", "docx_scan.py"), real, "--all",
                  "--out", os.path.join(HERE, "scan-word"), "--quiet")
    print("   scan 退出码 %d（3 = 有硬阻，符合预期）" % rc)
    rc, out = run(os.path.join(ROOT, "scripts", "docx_apply.py"), real,
                  "--out", clean, "--apply")
    print(out.rstrip())
    if rc != 0:
        print("❌ apply 失败")
        return 1
    print("✅ apply 与 guard 全过")

    print("\n── W3 用 Word 把清理后的文档读回来 ──")
    p = ps("word_read.ps1", clean, out_txt)
    if p.returncode != 0:
        print("❌ Word 无法读取清理后的文档（很可能触发了「需要修复」）：%s" % p.stderr.strip()[:300])
        return 1
    got = open(out_txt, encoding="utf-8").read()
    print("✅ Word 读回成功（%s）" % p.stdout.strip())

    print("\n── W4 比对文字 ──")
    bad = [want for want in EXPECT_CLEAN if want not in got]
    left = [d for d in EXPECT_GONE if d in got]
    print("   应出现的 %d 条，缺少 %d 条 %s" % (len(EXPECT_CLEAN), len(bad), bad or ""))
    print("   应消失的 %d 条，仍在 %d 条 %s" % (len(EXPECT_GONE), len(left), left or ""))
    if bad or left:
        print("❌ Word 往返后的文字与预期不符")
        with open(os.path.join(HERE, "word-text-out.txt"), encoding="utf-8") as fh:
            pass
        print("---- Word 读回的原文 ----")
        print(got)
        return 1
    print("\n" + "=" * 56)
    print("✅ 真 Word 往返验证全部通过：Word 打得开、读得对、脏痕迹确实没了")
    return 0


if __name__ == "__main__":
    sys.exit(main())
