#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ai-trace-cleanup · 文本级保护区（掩码）

文档里有一类 token，它们的**点号和标点是语义的一部分**，绝不能按中文排版规则去改：

    .github/workflows/          ← 把 `.` 改成 `。` 就废了
    C:\\Users\\Administrator\\.pi  ← 同上，而且路径里的空格也不能并
    os.path.join                ← 同上
    python docx_scan.py --all   ← `--` 会被当成破折号改成 `——`
    `code span`                 ← 里面是代码，一个字符都不许动
    run-all.ps1                 ← 同上

所以规则引擎在动手之前先扫一遍**掩码**：落在掩码区间里的字符，I/S/P 三层既不修改也不报告。
这是"保护"，不是"漏报"—— 报告里不会出现"这里有个空格要不要删"，因为它压根不该删。

三层设计里的分工
----------------
- **字节级保护区**（哪些 `<w:t>` 不许动）→ `references/docx-protection.md`
- **文本级保护区**（哪些字符序列不许动）→ 就是这个文件
- 两层都由代码强制，不靠"我记得跳过"。

加一条掩码前必须自问两句
------------------------
1. **它会不会挡住我们本来想做的修复？** 例如不能给 `%` 加掩码，否则 `50 %` → `50%` 永远修不了；
   不能给全角 `：` 加掩码，否则 `9：30` → `9:30` 也修不了。
   所以本文件里的时间、单位、数字类掩码**只匹配半角字符**。
2. **它会不会把正常中文标点也吞掉？** 例如「检查完成, 未完成.」里的 `,` `.` 必须保持未掩码，
   才能被 P01 转成全角。

这两条都有反向测试守着（`tests/selftest.py` 的 T8）。
"""

from __future__ import annotations

import re

# ---------------------------------------------------------------- 片段定义

#: token 的停止字符：空白 + 中文标点 + 引号类 + 管道/尖括号 + 半角逗号分号。
#: 刻意**不含** ASCII 连字符 `-` 和 `.`（文件名里极常见）。
_STOP = r"""\s，。、；：？！“”‘’（）【】《》〈〉「」『』…—,;<>|"'`"""

#: 路径里的一小段（可含中文，但不含空白与停止符）—— 中文目录名靠它
_ANY = r"[^%s\s]" % _STOP
#: 纯 ASCII 段（文件名主体）
_SEG_ASCII = r"[A-Za-z0-9_.\-~]"
_SEP = r"[\\/]"

#: 路径尾部。为什么长这样：
#:   · 三种单位交替出现：`\段\`（被分隔符夹住的内容，中文目录名靠它）、
#:     可选的 `\` + 纯 ASCII 段（`Users`、`Administrator`、`.pi`、`2026-W42`）、单独的 `\`；
#:   · 于是 `C:\Users\Administrator\Desktop\周报\计划\2026-W42` 整条落在保护里；
#:   · 而 `放在C:\temp下` 的 `下`（末尾中文、后面没有分隔符）**不进保护** ——
#:     否则会把紧跟路径的中文正文一起吞进去，让那句正文里的问题漏报。
_PATH_TAIL = r"(?:%s%s+%s|%s?%s+|%s+)*" % (_SEP, _ANY, _SEP, _SEP, _SEG_ASCII, _SEP)


# ---------------------------------------------------------------- 家族表

#: (id, 名称, 覆盖什么, 正则)
FAMILIES = [
    # ---- 网络地址 ----
    ("P-URL", "URL", "带协议的网址",
     r"https?://[^%s]+" % _STOP),
    ("P-WWW", "www 地址", "没有协议头的网址",
     r"www\.[A-Za-z0-9.\-]+(?:/[^%s]*)?" % _STOP),
    ("P-MAIL", "电子邮箱", "邮箱地址里的点和 @ 都是语义",
     r"[\w.+-]+@[\w-]+\.[\w.-]+"),
    ("P-DOMAIN", "域名 / 主机", "域名、主机名、带端口或路径的地址（也顺带保护 os.path 这类点号名）",
     r"[A-Za-z0-9][A-Za-z0-9\-]*(?:\.[A-Za-z0-9\-]+)+(?::\d+)?(?:/[^%s]*)?" % _STOP),

    # ---- 文件系统路径 ----
    ("P-WIN", "Windows 路径", r"盘符路径：C:\Users\Administrator\.pi\agent\skills\（中文目录名也一起保护）",
     r"[A-Za-z]:[\\/]" + _PATH_TAIL),
    ("P-UNC", "UNC 网络路径", r"\\server\share 形式的路径",
     r"\\\\%s+%s%s" % (_ANY, _SEP, _PATH_TAIL)),
    ("P-POSIX", "POSIX 绝对路径", "/usr/local/bin、/etc/hosts",
     r"/(?:%s+/)+%s*" % (_ANY, _SEG_ASCII)),
    ("P-REL", "相对 / 家目录路径", "./x、../y、~/.pi/agent 形式的路径",
     r"(?:\.\.?/|~/)" + _PATH_TAIL),
    ("P-SLASH", "含斜杠的 token", "tests/、a/b/c、and/or、n/a 这种一段带一段的写法",
     r"%s+/%s+(?:/%s+)*" % (_SEG_ASCII, _SEG_ASCII, _SEG_ASCII)),

    # ---- 点号类（`.github`、`os.path`、`README.md` 全靠它）----
    ("P-DOTNAME", "点号开头的名字", ".github、.gitignore、.pi、os.path、git.exe 这类「点+字母」整体保护",
     r"\.[A-Za-z][A-Za-z0-9_.\-]*"),
    ("P-FILEEXT", "带扩展名的文件名", "run-all.ps1、ci.yml、data.7z（含数字开头的扩展名）",
     r"[\w\-]+\.(?:"
     r"docx?|docm|dotx?|xlsx?|xlsm|xlsb|csv|tsv|pptx?|ppsx?|"
     r"pdf|md|markdown|rst|txt|log|rtf|epub|mobi|"
     r"py|pyi|ipynb|js|mjs|cjs|jsx|ts|tsx|json5?|jsonc|ya?ml|toml|ini|cfg|conf|env|properties|"
     r"sh|bash|zsh|fish|ps1|psm1|psd1|bat|cmd|vbs|"
     r"c|h|cc|cpp|hpp|cxx|cs|java|kt|kts|go|rs|rb|php|pl|pm|lua|swift|m|mm|scala|dart|ex|exs|erl|hs|jl|r|"
     r"html?|xhtml|css|scss|sass|less|styl|vue|svelte|"
     r"xml|xsd|xsl|svg|plist|service|desktop|reg|"
     r"sql|db|sqlite|mdb|accdb|"
     r"png|jpe?g|jfif|gif|webp|bmp|tiff?|ico|heic|avif|psd|ai|sketch|fig|"
     r"mp3|wav|flac|aac|m4a|ogg|opus|mp4|mov|avi|mkv|webm|wmv|flv|"
     r"zip|rar|7z|tar|gz|bz2|xz|tgz|jar|war|apk|ipa|dmg|iso|exe|msi|dll|so|dylib|bin|dat|lock|"
     r"ttf|otf|woff2?|eot|map|min|patch|diff"
     r")\b"),

    # ---- 代码与命令行 ----
    ("P-CODE", "反引号代码段", "`…` 之间的内容（含反引号）一个字都不许动",
     r"`[^`\n]+`"),
    ("P-FLAG", "命令行参数", "--flag、-o；前面必须是分隔符，且不能是连字符，避免误吞 A--B",
     r"(?<![A-Za-z0-9_-])--?[A-Za-z][A-Za-z0-9\-]*"),
    ("P-ENV", "环境变量", "$HOME、${VAR}、%USERPROFILE%",
     r"\$\{?[A-Za-z_][A-Za-z0-9_]*\}?|%[A-Za-z_][A-Za-z0-9_]*%"),
    ("P-IDENT", "下划线标识符", "snake_case 形式的代码标识符",
     r"[A-Za-z_][A-Za-z0-9_]*_[A-Za-z0-9_]+"),
    ("P-CAMEL", "驼峰标识符", "getElementById、parseXML 这类驼峰命名",
     r"\b[a-z][A-Za-z0-9]*[A-Z][A-Za-z0-9]*\b"),
    ("P-TASKID", "编号类 token", "JIRA-1234、UTF-8、ISO-9001",
     r"[A-Za-z]{2,}-\d+"),
    ("P-KEYCOLON", "配置行的键", "段落以 key: 开头时保护那个半角冒号（避免被 P01 转成全角）",
     r"(?m:^[ \t]*[A-Za-z_\"'][A-Za-z0-9_.\-\"'][A-Za-z0-9_.\-\"']*[ \t]*:)"),

    # ---- 数字类（只匹配半角，否则会挡住 P07/P08 的修复）----
    ("P-NUM", "小数 / 版本号", "1.5、v0.1.0（也保护全角点，因为没有规则会去修它）",
     r"\d+[.．]\d+(?:[.．]\d+)*"),
    ("P-THOUSAND", "千分位数字", "1,234,567",
     r"\d{1,3}(?:[,，]\d{3})+"),
    ("P-TIME", "时间", "只保护半角冒号；全角 `9：30` 要留给 P08 去修",
     r"\d{1,2}:\d{2}(?::\d{2})?"),
    ("P-DATE", "日期", "2026-10-13、2026/10/13",
     r"\d{4}[-/]\d{1,2}[-/]\d{1,2}"),
    ("P-UNIT", "带单位的数值", "10 kg、100 MB（不含 % 与 ℃ —— 那两个要留给 S06 修）",
     r"\d+(?:\.\d+)*\s*(?:kg|g|mg|t|km|cm|mm|um|nm|ml|L|"
     r"KiB|MiB|GiB|TiB|KB|MB|GB|TB|B|kbps|Mbps|Gbps|"
     r"Hz|kHz|MHz|GHz|ms|min|V|mV|kV|W|kW|mA|kWh|Pa|hPa|dB|px|pt|em|rem|dpi|fps|ppm|rpm|qps|tps)\b"),

    # ---- 其他 ----
    ("P-ENTITY", "HTML 实体", "&nbsp;、&#160;、&#xA0;",
     r"&[A-Za-z]{2,10};|&#\d{2,6};|&#x[0-9A-Fa-f]{2,6};"),
]

MASK_PATTERNS = [f[3] for f in FAMILIES]

_MASK_RE = re.compile("|".join("(?:%s)" % p for p in MASK_PATTERNS), re.M)


def masks_of(text: str) -> list:
    """返回 [(start, end)]：落在这些区间内的字符，I/S/P 三层既不修改也不报告。"""
    if not text:
        return []
    return [(m.start(), m.end()) for m in _MASK_RE.finditer(text)]


def protected_tokens(text: str) -> list:
    """返回 (start, end, 文本)，调试与测试用。"""
    if not text:
        return []
    return [(m.start(), m.end(), m.group(0)) for m in _MASK_RE.finditer(text)]


def families_md() -> str:
    """给文档用的家族清单（`docx_scan.py --list-protections` 写进 references/）。"""
    lines = [
        "<!-- 本文件由 `python scripts/docx_scan.py --list-protections` 生成，不要手改。 -->",
        "",
        "# 文本级保护区（掩码）家族清单",
        "",
        "这些 token 里的点号与标点是语义的一部分，机械层（I/S/P）**既不修改也不报告**。",
        "规则在 `scripts/protect.py`；为什么这么做见 `references/docx-protection.md` 第六节。",
        "",
        "| ID | 名称 | 覆盖什么 |",
        "| --- | --- | --- |",
    ]
    for fid, name, desc, _pat in FAMILIES:
        lines.append("| `%s` | %s | %s |" % (fid, name, desc.replace("|", "\\|")))
    lines += [
        "",
        "## 刻意**不**保护的东西（否则会挡住该做的修复）",
        "",
        "| 不保护 | 因为 |",
        "| --- | --- |",
        "| 半角 `%`、`℃`、`°` 前的空格 | `50 %` → `50%` 是 `S06` 要修的目标 |",
        "| 全角 `：`（如 `9：30`） | `9：30` → `9:30` 是 `P08` 要修的目标，所以时间掩码只匹配半角冒号 |",
        "| 全角数字 `１２３` | → `123` 是 `P07` 要修的目标 |",
        "| 中文语境里的半角 `,` `.` | 转全角是 `P01` 要修的目标 |",
        "| 字母间的双连字符 `A--B` | `--` → `——` 是 `P04` 要修的目标，所以命令行掩码要求 `--` 前是分隔符 |",
        "| 全角波浪号 `～`（如 `3～5`） | → `-` 是 `P09` 要修的目标 |",
        "| 紧跟路径之后、且其后没有分隔符的中文串 | 保护它会把那句正文里的问题一起吞掉；实测 `放在C:\\temp下` 只保护 `C:\\temp` |",
        "",
        "> 这些都有反向测试守着（`tests/selftest.py` 的 T8）：谁要是把掩码写宽了，测试立刻变红。",
    ]
    return "\n".join(lines)


__all__ = ["FAMILIES", "MASK_PATTERNS", "masks_of", "protected_tokens", "families_md"]
