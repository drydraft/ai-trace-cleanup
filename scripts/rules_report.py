#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ai-trace-cleanup · 报告层规则（M / B / F / L / C）

这一层的规则**永不落盘**，只产出建议。原因：
  · M 层要把字面 Markdown 换成 Word 原生格式 —— 那是重排版，不是清洗；
  · B/F 层要动段落与样式 —— 一动就改版式，超出「去痕」范围；
  · L/C 层是语义判断 —— 机器只负责**指出**，删改必须由人拍板。

判据一律写死成可判定的形式（正则/计数/比例），不写"读起来像 AI"这种没法验的话。
"""

from __future__ import annotations

import os
import re

from docx_text import char_kind
from rules_core import SEV_HARD, SEV_SUG, SEV_WARN, Issue, _anchor

LEXICON_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "assets", "lexicon-zh.txt")


# ---------------------------------------------------------------- 词表

def load_lexicon(path: str = None) -> dict:
    """读 `词|类别|建议` 三列的词表；读不到就退回内置最小集。"""
    path = path or LEXICON_PATH
    cats = {}
    try:
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = [x.strip() for x in line.split("|")]
                if len(parts) < 2:
                    continue
                term, cat = parts[0], parts[1]
                sug = parts[2] if len(parts) > 2 else "（删或改写）"
                cats.setdefault(cat, []).append((term, sug))
    except OSError:
        cats = {"filler": [("首先", "（删，或改“先看”）"), ("综上所述", "（删整段结论）")]}
    return cats


CAT_RULE = {
    "filler": ("L01", "套话连接词", SEV_WARN),
    "summary": ("L02", "模板化总结词", SEV_WARN),
    "vague": ("L03", "空泛名词", SEV_SUG),
    "hollow_adv": ("L04", "空泛副词/形容词", SEV_SUG),
    "meta": ("L05", "元话语/旁白", SEV_WARN),
    "era": ("L08", "时代背景开场", SEV_WARN),
}

#: 报告层的规则目录。**这是唯一真理源**：`docx_scan.py --list-rules` 按它生成
#: `references/rule-index.md`，报告里也用它取标题。加规则必须同时加在这里。
REPORT_RULE_META = {
    "B01": ("B", SEV_WARN, "段内手动换行符", "`<w:br/>`（Shift+Enter）被当成排版手段"),
    "B02": ("B", SEV_WARN, "多余空段落", "用空段落凑间距，而非样式里的段前/段后间距"),
    "B03": ("B", SEV_WARN, "连续性短段（一句话一段）", "连续 ≥5 个 ≤25 字的段落"),
    "B04": ("B", SEV_WARN, "空格伪缩进", "段首用空格或全角空格做首行缩进"),
    "B05": ("B", SEV_SUG, "句子缺终结标点", "以汉字结尾、没有。？！"),
    "B06": ("B", SEV_SUG, "单段过长", ">300 字不分段"),
    "F01": ("F", SEV_WARN, "字号漂移", "正文出现 >2 种字号"),
    "F02": ("F", SEV_WARN, "同段混用字体", "同一段内 >1 种字体"),
    "F03": ("F", SEV_WARN, "网页/英文默认字体", "正文里出现 Calibri / Times New Roman 等"),
    "F04": ("F", SEV_SUG, "非黑文字", "正文出现彩色字"),
    "F05": ("F", SEV_WARN, "加粗率过高", ">15% 的字符是加粗"),
    "F06": ("F", SEV_SUG, "段落对齐不统一", "justify 与 left 混用"),
    "F07": ("F", SEV_SUG, "直接格式绕过样式", "≥3 个无样式段落单设了间距"),
    "F08": ("F", SEV_WARN, "用加粗大字当标题", "有标题外观但没用 Heading 样式"),
    "F09": ("F", SEV_SUG, "未设中文字体", "只给了 w:ascii、没给 w:eastAsia"),
    "C01": ("C", SEV_HARD, "自指/AI 身份", "作为 AI／作为一个语言模型"),
    "C02": ("C", SEV_HARD, "服务性寒暄", "希望对你有帮助／请随时告诉我"),
    "C03": ("C", SEV_WARN, "免责声明", "以上内容仅供参考／具体以现场为准"),
    "C04": ("C", SEV_WARN, "空壳总结段", "40–160 字、含总结词、零数字"),
    "C05": ("C", SEV_WARN, "重复句", ">=12 字的同句在文中出现两次"),
    "C06": ("C", SEV_HARD, "占位符残留", "[待补充]／TODO／此处省略"),
    "C07": ("C", SEV_SUG, "承诺式开头", "下面我将从三个方面／本文将从"),
    "L06": ("L", SEV_WARN, "对举句式", "不是…而是…／并非…而是…"),
    "L07": ("L", SEV_SUG, "递进套式", "不仅…而且…／既…又…"),
    "L09": ("L", SEV_SUG, "空壳谓语", "促进了…的发展／有效地提升了…"),
    "L10": ("L", SEV_SUG, "三项排比密度", ">3 个段落出现「A、B、C」结构"),
    "L11": ("L", SEV_SUG, "破折号密度", ">6 处/千字"),
}


# ================================================================ M 层

M_PATTERNS = [
    ("M01", SEV_HARD, "字面加粗标记", r"\*\*[^*\n]{1,80}\*\*|__[^_\n]{1,80}__",
     "改为 Word 加粗格式（选中文字 → Ctrl+B），不要留字面 **"),
    ("M02", SEV_WARN, "字面斜体标记", r"(?<![\w*])\*[^*\n]{1,60}\*(?![\w*])",
     "改为 Word 斜体格式"),
    ("M03", SEV_HARD, "字面标题标记", r"(?m)^\s{0,3}#{1,6}\s+\S",
     "改为 Word 标题样式（样式栏选「标题 1/2/3」）"),
    ("M04", SEV_HARD, "手打列表标记", r"(?m)^\s{0,3}(?:[-*+]\s+|\d+[.、)]\s+)",
     "改为 Word 自动编号/项目符号（开始 → 项目符号/编号）"),
    ("M05", SEV_WARN, "字面行内代码", r"`[^`\n]{1,80}`",
     "用等宽字体（Consolas/等线）表示，去掉反引号"),
    ("M06", SEV_HARD, "字面代码围栏", r"(?m)^\s{0,3}(?:```|~~~)",
     "代码块应改为段落底纹 + 等宽字体，或改用表格框"),
    ("M07", SEV_HARD, "字面 Markdown 链接", r"\[[^\]\n]{1,80}\]\([^)\n]{1,200}\)",
     "拆成正文文字 + Word 超链接（插入 → 链接）"),
    ("M08", SEV_HARD, "字面表格分隔行", r"(?m)^\s*\|?[\s:|\-]{6,}\|?\s*$",
     "Markdown 表格需重画成 Word 表格"),
    ("M09", SEV_WARN, "字面水平线", r"(?m)^\s{0,3}(?:-{3,}|\*{3,}|_{3,})\s*$",
     "删掉，改用段落边框或分页"),
    ("M10", SEV_WARN, "字面引用块", r"(?m)^\s{0,3}>\s?\S",
     "改用 Word「引用」样式或缩进段落"),
    ("M11", SEV_SUG, "转义符残留", r"\\[*_`\[\]()#+\-.!]",
     "删除多余反斜杠"),
    ("M12", SEV_WARN, "HTML 实体/标签残留", r"&(?:nbsp|amp|lt|gt|quot|#\d+);|<br\s*/?>|</?(?:p|div|span|b|i|code)\b[^>]*>",
     "替换为真实字符或删除标签"),
]


def scan_markdown(para) -> list:
    text = para.text
    out = []
    for rid, sev, title, pat, note in M_PATTERNS:
        for m in re.finditer(pat, text):
            out.append(Issue(rule=rid, layer="M", severity=sev, part=para.part,
                             pid=para.pid, anchor=_anchor(para),
                             start=m.start(), end=m.end(), orig=m.group(0),
                             repl="", note=note))
    return out


# ================================================================ B 层

SHORT_PARA_CHARS = 25
SHORT_RUN_MIN = 5           # 连续多少个短段才算「一句话一段」
LONG_PARA_CHARS = 300
_M_QUICK = re.compile("|".join(p.replace("(?m)", "") for _, _, _, p, _ in M_PATTERNS), re.M)


def scan_structure(paras) -> list:
    out = []
    empty_run = 0
    short_run = 0
    first_short = None
    for para in paras:
        raw, body = para.text, para.body()
        is_empty = not body
        # ---- B01 手动换行符
        if para.manual_breaks:
            idx = raw.find("\n")
            out.append(Issue(rule="B01", layer="B", severity=SEV_WARN, part=para.part,
                             pid=para.pid, anchor=_anchor(para), start=idx, end=idx + 1,
                             orig="\\n", repl="",
                             note="段内手动换行符（Shift+Enter），%d 处。应改为分段，"
                                  "或用「段前/段后间距」控距" % para.manual_breaks))
        # ---- B02 多余空段落
        if is_empty:
            empty_run += 1
            continue
        if empty_run:
            out.append(Issue(rule="B02", layer="B", severity=SEV_WARN, part=para.part,
                             pid=para.pid, anchor="(空段落 ×%d)" % empty_run,
                             note="连续 %d 个空段落。AI 常用空行凑间距，应改用样式里的段前/段后间距"
                                  % empty_run))
            empty_run = 0

        # ---- B04 段首空格伪缩进
        lead = re.match(r"[ \u3000]+", raw)
        if lead and body:
            out.append(Issue(rule="B04", layer="B", severity=SEV_WARN, part=para.part,
                             pid=para.pid, anchor=_anchor(para), start=0, end=lead.end(),
                             orig=raw[:lead.end()], repl="",
                             note="用空格做首行缩进。应改为段落 → 缩进 → 首行缩进 2 字符"))

        is_heading = para.style.startswith("Heading") or para.style.startswith("标题")
        # ---- B03 连续短句成段
        if body and not is_heading and len(body) <= SHORT_PARA_CHARS and para.manual_breaks == 0:
            if short_run == 0:
                first_short = para
            short_run += 1
        else:
            if short_run >= SHORT_RUN_MIN:
                out.append(Issue(rule="B03", layer="B", severity=SEV_WARN,
                                 part=first_short.part, pid=first_short.pid,
                                 anchor="(以下连续 %d 个短段)" % short_run,
                                 note="从本段起连续 %d 个 ≤%d 字的短段落，是「一句话一段」的典型 AI 排版；"
                                      "人写的中文正文通常会把同类意思合并成段"
                                      % (short_run, SHORT_PARA_CHARS)))
            short_run = 0
            first_short = None

        # ---- B06 单段过长
        if body and len(body) > LONG_PARA_CHARS:
            out.append(Issue(rule="B06", layer="B", severity=SEV_SUG, part=para.part,
                             pid=para.pid, anchor=_anchor(para),
                             note="单段 %d 字，超过 %d 字建议按语义切分" % (len(body), LONG_PARA_CHARS)))

        # ---- B05 正文缺句末标点（只打「最后一个字是汉字且无标点」这种真信号）
        if (body and not is_heading and len(body) >= 20 and not para.is_list
                and char_kind(body[-1]) == "cjk" and not _M_QUICK.search(para.text)):
            out.append(Issue(rule="B05", layer="B", severity=SEV_SUG, part=para.part,
                             pid=para.pid, anchor=_anchor(para),
                             note="句子以汉字结尾、没有终结标点（AI 断行没接上的典型表现）"))

    if empty_run:
        out.append(Issue(rule="B02", layer="B", severity=SEV_WARN, part=paras[-1].part,
                         pid=paras[-1].pid, anchor="(文末空段落 ×%d)" % empty_run,
                         note="文末连续 %d 个空段落" % empty_run))
    if short_run >= SHORT_RUN_MIN:
        out.append(Issue(rule="B03", layer="B", severity=SEV_WARN,
                         part=first_short.part, pid=first_short.pid,
                         anchor="(以下连续 %d 个短段)" % short_run,
                         note="文末连续 %d 个 ≤%d 字的短段落" % (short_run, SHORT_PARA_CHARS)))
    return out


# ================================================================ F 层

WEB_FONTS = {"Calibri", "Times New Roman", "Arial", "Segoe UI", "Helvetica",
             "Cambria", "Verdana", "Georgia", "Roboto", "Inter", "Consolas"}
CN_FONTS = {"宋体", "SimSun", "微软雅黑", "Microsoft YaHei", "等线", "DengXian",
            "仿宋", "FangSong", "黑体", "SimHei", "楷体", "KaiTi", "思源宋体", "思源黑体"}
#: 由 pPr 直接指定、绕过样式的段落属性
DIRECT_PPR = ("<w:spacing", "<w:ind", "<w:jc", "<w:numPr")
BOLD_RATIO_LIMIT = 0.15


def scan_formatting(paras) -> list:
    out = []
    sizes, fonts, web_hits, colors = {}, {}, [], []
    bold_chars = total_chars = 0
    jc_seen, direct_spacing, no_east_asia = set(), [], []
    mixed_font_paras = []

    for para in paras:
        body = para.body()
        if not body:
            continue
        total_chars += len(body)
        fams = set()
        for sp in para.spans:
            if not sp.text.strip():
                continue
            if sp.bold:
                bold_chars += len(sp.text)
            if sp.size:
                sizes[sp.size] = sizes.get(sp.size, 0) + 1
            for f in sp.fonts:
                if f:
                    fonts[f] = fonts.get(f, 0) + 1
                    fams.add(f)
                    if f in WEB_FONTS:
                        web_hits.append((para, f, sp))
            if sp.color and sp.color.lower() not in ("auto", "000000"):
                colors.append((para, sp.color, sp))
        if len(fams) > 1:
            mixed_font_paras.append((para, sorted(fams)))
        if para.jc:
            jc_seen.add(para.jc)
        if para.spacing and not para.style:
            direct_spacing.append(para)
        # 只设了西文字体、没设 w:eastAsia —— 中文会跟着西文字体跑
        if (any(char_kind(c) == "cjk" for c in body)
                and any(sp.fonts for sp in para.spans)
                and not any(sp.ea_font for sp in para.spans)):
            no_east_asia.append(para)

    # ---- F01 字号漂移（含只出现一次的字号，那正是「跳了一下」的证据）
    body_sizes = sorted(sizes, key=lambda x: int(x or 0))
    if len(body_sizes) > 2:
        out.append(Issue(rule="F01", layer="F", severity=SEV_WARN, part="(全文)", pid="-",
                         anchor="字号集合=%s" % ",".join(body_sizes),
                         note="正文出现 %d 种字号（半磅值 %s）。人写的一份文档通常只有 2–3 级字号，"
                              "字号乱跳是「多段拼凑」的直接证据" % (len(body_sizes), body_sizes)))

    # ---- F02 同段混用字体
    for para, fams in mixed_font_paras[:8]:
        out.append(Issue(rule="F02", layer="F", severity=SEV_WARN, part=para.part,
                         pid=para.pid, anchor=_anchor(para),
                         note="同一段内混用 %d 种字体：%s。中文文档正文应统一" % (len(fams), "、".join(fams))))

    # ---- F03 网页/英文默认字体
    seen = set()
    for para, f, sp in web_hits:
        if (f, para.pid) in seen:
            continue
        seen.add((f, para.pid))
        out.append(Issue(rule="F03", layer="F", severity=SEV_WARN, part=para.part,
                         pid=para.pid, anchor=_anchor(para), orig=f,
                         note="正文里出现 %s —— 这是从网页/其他文档粘贴进来的字体，"
                              "在中文文档里会明显「格格不入」" % f))
        if len(seen) >= 10:
            break

    # ---- F04 非黑文字
    for para, color, sp in colors[:8]:
        out.append(Issue(rule="F04", layer="F", severity=SEV_SUG, part=para.part,
                         pid=para.pid, anchor=_anchor(para), orig="#" + color,
                         note="彩色文字 #%s。正式文档正文宜用自动色/黑色" % color))

    # ---- F05 加粗比例
    if total_chars and bold_chars / total_chars > BOLD_RATIO_LIMIT:
        out.append(Issue(rule="F05", layer="F", severity=SEV_WARN, part="(全文)", pid="-",
                         anchor="加粗 %.1f%%" % (100 * bold_chars / total_chars),
                         note="全文 %.1f%% 的字符是加粗（建议 ≤%.0f%%）。AI 爱用加粗当强调，"
                              "一片黑里全是重点等于没有重点"
                              % (100 * bold_chars / total_chars, 100 * BOLD_RATIO_LIMIT)))

    # ---- F06 对齐方式混用
    if len(jc_seen) > 1:
        out.append(Issue(rule="F06", layer="F", severity=SEV_SUG, part="(全文)", pid="-",
                         anchor="对齐=%s" % ",".join(sorted(jc_seen)),
                         note="段落对齐方式不统一（%s）。同一份文档正文宜统一" % ",".join(sorted(jc_seen))))

    # ---- F07 直接格式绕过样式
    if len(direct_spacing) >= 3:
        out.append(Issue(rule="F07", layer="F", severity=SEV_SUG, part="(全文)", pid="-",
                         anchor="%d 段" % len(direct_spacing),
                         note="有 %d 个无样式段落单独设了段前/段后间距。逐段手调间距是 AI 拼装的典型手法，"
                              "应统一到样式里" % len(direct_spacing)))

    # ---- F08 用加粗大字当标题（没有用 Heading 样式）
    fake_heads = []
    for para in paras:
        body = para.body()
        if not body or para.style or len(body) > 40 or para.is_list:
            continue
        big = [sp for sp in para.spans if sp.text.strip()
               and sp.size and int(sp.size or 0) >= 28 and sp.bold]
        if big and sum(len(sp.text) for sp in big) >= len(body) * 0.6:
            fake_heads.append(para)
    for para in fake_heads[:8]:
        out.append(Issue(rule="F08", layer="F", severity=SEV_WARN, part=para.part,
                         pid=para.pid, anchor=_anchor(para),
                         note="这一行是「加粗 + 大字号」的手工标题，没用 Word 的标题样式。"
                              "后果：不能生成目录、导航窗格看不见、统一改字号时漏掉它"))

    # ---- F09 缺 eastAsia 字体设置
    if len(no_east_asia) >= 3:
        out.append(Issue(rule="F09", layer="F", severity=SEV_SUG, part="(全文)", pid="-",
                         anchor="%d 段" % len(no_east_asia),
                         note="%d 个含中文的段落只设了西文字体、没设中文字体（w:eastAsia），"
                              "换台机器或换 Word 版本打开，中文会变字体" % len(no_east_asia)))
    return out


# ================================================================ L 层

L_STRUCT = [
    ("L06", SEV_WARN, "对举句式「不是…而是…」",
     r"不是[^，。；\n]{1,24}，?而是|并非[^，。；\n]{1,24}，?而是",
     "改成直陈句，例如「问题在流程，不在设备」"),
    ("L07", SEV_SUG, "递进套式「不仅…而且…」",
     r"不仅[^，。；\n]{1,30}[，,][^。\n]{0,20}(?:而且|还|也)|既[^，。；\n]{1,20}又[^，。；\n]{1,20}",
     "直接写结论，删掉递进壳子"),
    ("L09", SEV_SUG, "空壳谓语",
     r"(?:促进|推动|助力|赋能|赋能于|提升|增强)[^，。；\n]{0,12}的(?:发展|提升|增长|进步|优化|转型)",
     "换成具体动作 + 具体结果"),
]

#: 三项排比：单次出现很正常，密度高才是 AI 节奏
TRIAD_RE = re.compile(r"[^\s，。；、]{2,8}、[^\s，。；、]{2,8}、[^\s，。；、]{1,8}")
TRIAD_PARA_LIMIT = 3
DASH_PER_KILO = 6
DASH_CHARS = "—–"


def scan_language(paras, lexicon=None) -> list:
    lexicon = lexicon if lexicon is not None else load_lexicon()
    out = []
    total = sum(len(p.body()) for p in paras)
    dash_n = 0
    triad_paras = []
    for para in paras:
        text = para.visible_text
        if not text.strip():
            continue
        # 词表驱动。同一处会被长词和短词同时命中（“综上所述”里含“综上”），
        # 所以按词长降序、跳过已被长词占用的区间。
        for cat, terms in lexicon.items():
            rid, title, sev = CAT_RULE.get(cat, ("L99", cat, SEV_SUG))
            taken = []
            for term, sug in sorted(terms, key=lambda x: -len(x[0])):
                start = 0
                while True:
                    i = text.find(term, start)
                    if i < 0:
                        break
                    start = i + 1
                    if any(i < b and a < i + len(term) for a, b in taken):
                        continue
                    taken.append((i, i + len(term)))
                    out.append(Issue(rule=rid, layer="L", severity=sev, part=para.part,
                                     pid=para.pid, anchor=_anchor(para), start=i,
                                     end=i + len(term), orig=term, repl="",
                                     note="%s：命中「%s」—— %s" % (title, term, sug)))
        # 结构型
        for rid, sev, title, pat, note in L_STRUCT:
            for m in re.finditer(pat, text):
                out.append(Issue(rule=rid, layer="L", severity=sev, part=para.part,
                                 pid=para.pid, anchor=_anchor(para), start=m.start(),
                                 end=m.end(), orig=m.group(0), repl="",
                                 note="%s：%s" % (title, note)))
        if TRIAD_RE.search(text):
            triad_paras.append(para)
        dash_n += sum(1 for c in text if c in DASH_CHARS)

    if len(triad_paras) > TRIAD_PARA_LIMIT:
        out.append(Issue(rule="L10", layer="L", severity=SEV_SUG, part=triad_paras[0].part,
                         pid=triad_paras[0].pid,
                         anchor="(全文 %d 段有三项排比)" % len(triad_paras),
                         note="%d 个段落都出现「A、B、C」三项排比（阈值 %d）。这是 AI 的标志性节奏；"
                              "人写中文时三项会自然变成两项或四项" % (len(triad_paras), TRIAD_PARA_LIMIT)))
    if total and dash_n / max(1, total / 1000.0) > DASH_PER_KILO:
        out.append(Issue(rule="L11", layer="L", severity=SEV_SUG, part="(全文)", pid="-",
                         anchor="破折号 %d 处 / %d 字" % (dash_n, total),
                         note="破折号密度 %.1f 处/千字（阈值 %d）。破折号是 AI 最爱用的连接方式"
                              % (dash_n / max(1, total / 1000.0), DASH_PER_KILO)))
    return out


# ================================================================ C 层

C_PATTERNS = [
    ("C01", SEV_HARD, "自指/AI 身份", r"作为\s*(?:一个)?\s*(?:AI|人工智能|语言模型|大模型|智能助手)|我是\s*(?:一个)?\s*(?:AI|人工智能|语言模型)",
     "整句删除。这是最硬的 AI 痕迹——交付文档里绝不能出现"),
    ("C02", SEV_HARD, "服务性寒暄", r"希望对(?:你|您|大家)有(?:所)?帮助|如需(?:进一步|更多)[^。\n]{0,20}(?:请|可)|"
                            r"(?:请|可以)随时(?:告诉|联系)我|我来(?:帮你|为你)|以上是我(?:为|给)你",
     "整句删除。这些是对话里的客套，不是文档内容"),
    ("C03", SEV_WARN, "免责声明", r"以上内容仅供参考|仅供参考[，,]?不构成|具体请以[^。\n]{0,16}为准|"
                          r"如有(?:错误|疏漏)[^。\n]{0,12}(?:指正|告知)|不构成任何[^。\n]{0,12}建议",
     "按文档性质决定删或改写成正式的口径说明"),
    ("C06", SEV_HARD, "占位符残留", r"\[待补充\]|【待补充】|\bTODO\b|\bTBD\b|此处(?:省略|略)|待填(?:写|充)?|XX+",
     "填上真实内容或删除。占位符出现在交付文档里是事故"),
    ("C07", SEV_SUG, "承诺式开头", r"下面我将[^。\n]{0,30}|接下来我们(?:将|来)?看|本文将(?:从|围绕)[^。\n]{0,30}|"
                          r"本文(?:将)?(?:从)[^。\n]{0,10}(?:个方面|个维度)",
     "改成直接进入内容，例如「巡检发现三类问题」"),
]


def scan_content(paras) -> list:
    out = []
    for para in paras:
        text = para.visible_text
        if not text.strip():
            continue
        for rid, sev, title, pat, note in C_PATTERNS:
            for m in re.finditer(pat, text):
                out.append(Issue(rule=rid, layer="C", severity=sev, part=para.part,
                                 pid=para.pid, anchor=_anchor(para), start=m.start(),
                                 end=m.end(), orig=m.group(0), repl="", note="%s：%s" % (title, note)))

    # ---- C04 空壳总结段（不限末尾：只要「有总结词 + 无数字 + 中等长度」就报）
    for para in [p for p in paras if p.body()][1:]:
        t = para.body()
        if 40 <= len(t) <= 160 and re.search(r"综上所述|总而言之|总的来说|综上|由此可见", t) \
                and not re.search(r"\d", t):
            out.append(Issue(rule="C04", layer="C", severity=SEV_WARN, part=para.part,
                             pid=para.pid, anchor=_anchor(para),
                             note="%d 字的模板化总结段，且不含任何数字/新信息。"
                                  "这类段落几乎只在 AI 输出里出现，建议整段删除" % len(t)))
            break

    # ---- C05 重复句
    seen = {}
    for para in paras:
        for sent in re.split(r"[。！？\n]", para.visible_text):
            s = sent.strip()
            if len(s) < 12:
                continue
            if s in seen:
                out.append(Issue(rule="C05", layer="C", severity=SEV_WARN, part=para.part,
                                 pid=para.pid, anchor=_anchor(para), orig=s,
                                 note="与 %s 处重复（%d 字）。AI 复读的典型表现" % (seen[s], len(s))))
            else:
                seen[s] = para.pid + ":" + para.part
    return out


# ================================================================ 总入口

def scan_report_layers(paras, lexicon=None, layers="MBFLC") -> list:
    out = []
    for layer in layers:
        if layer == "M":
            for para in paras:
                out += scan_markdown(para)
        elif layer == "B":
            out += scan_structure(paras)
        elif layer == "F":
            out += scan_formatting(paras)
        elif layer == "L":
            out += scan_language(paras, lexicon)
        elif layer == "C":
            out += scan_content(paras)
    return out
