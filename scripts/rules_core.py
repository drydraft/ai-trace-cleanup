#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ai-trace-cleanup · 规则核（Issue 数据结构 + 机械层规则 I/S/P）

分层与自动化边界（这是整个 skill 的骨架，不要随意挪动）
---------------------------------------------------------
    I  不可见字符    全自动
    S  空白          全自动（preset 控制取向）
    P  标点          全自动（有邻字符条件，不做全局替换）
    M  Markdown 残留 只报告（要转成 Word 原生格式，必须人做）
    B  段落与换行    只报告
    F  格式与样式    只报告
    L  语言层 AI 味  只报告
    C  内容层        只报告（删减需人确认）

铁律：本文件里的规则只允许做「字符级替换」，**永远不允许增删句子**。
任何会改变 token 集合的规则都不该出现在这里，guard 会当场抓住它。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx_text import char_kind  # noqa: F401  (tokens 由调用方从 docx_text 取)

LAYER_ORDER = ["I", "S", "P", "M", "B", "F", "L", "C"]
LAYER_NAME = {
    "I": "不可见字符", "S": "空白与间距", "P": "标点符号", "M": "Markdown 残留",
    "B": "段落与换行", "F": "格式与样式", "L": "语言层 AI 味", "C": "内容层",
}
#: 只有这几层允许自动应用
AUTO_LAYERS = ("I", "S", "P")
SEV_HARD, SEV_WARN, SEV_SUG = "hard", "warn", "sug"
SEV_MARK = {SEV_HARD: "⛔", SEV_WARN: "⚠️", SEV_SUG: "💡"}

CJK = r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\u3005\u3007"
CJK_PUNCT = "，。、；：？！“”‘’（）《》〈〉【】「」『』—…·～％"
ASCII_PUNCT = ",.;:?!()[]{}<>\"'`~-/%$&*+=@#|\\^_"


# ---------------------------------------------------------------- 数据结构

@dataclass
class Issue:
    rule: str
    layer: str
    severity: str
    part: str
    pid: str
    anchor: str
    start: int = -1
    end: int = -1
    orig: str = ""
    repl: str = ""
    note: str = ""
    auto: bool = False          # 是否可被 --apply 落盘
    span: object = None         # 替换落点（第一条命中的 Span）
    extra: dict = field(default_factory=dict)

    def as_dict(self) -> dict:
        d = dict(rule=self.rule, layer=self.layer, severity=self.severity,
                 part=self.part, pid=self.pid, anchor=self.anchor,
                 start=self.start, end=self.end, orig=self.orig,
                 repl=self.repl, note=self.note, auto=self.auto)
        if self.extra:
            d.update(self.extra)
        return d

    @property
    def mark(self) -> str:
        return SEV_MARK.get(self.severity, "·")


# ---------------------------------------------------------------- 掩码（文本级保护区）
#
# 具体有哪些家族、为什么这么写 → scripts/protect.py（单一真理源）。
# 这里只负责使用它：落在掩码里的字符，I/S/P 三层既不修改，也不报告。

from protect import masks_of  # noqa: E402  (re-export：规则核的门面仍叫 masks_of)


def _hits_mask(masks, s, e) -> bool:
    for a, b in masks:
        if s < b and a < e:
            return True
    return False


# ---------------------------------------------------------------- 正则工具

SP = "[ \u3000]"


def _re_all(pattern, text, repl_fn, flags=0):
    out = []
    for m in re.finditer(pattern, text, flags):
        if m.start() == m.end():
            continue
        out.append((m.start(), m.end(), repl_fn(m), None))
    return out


# ================================================================ I 层

def i01_invisible(text, ctx):
    """零宽字符与控制符：看不见但会污染检索、复制、比对。"""
    return _re_all(r"[\u200b\u200c\u200d\u2060\ufeff\u00ad\u202a-\u202e\u2066-\u2069]",
                   text, lambda m: "")


def i02_nbsp(text, ctx):
    """不换行空格 / 细空格 / 数字空格：按预设换成普通空格或直接删。"""
    def rep(m):
        i = m.start()
        left = text[i - 1] if i > 0 else ""
        right = text[m.end()] if m.end() < len(text) else ""
        if char_kind(left) in ("cjk", "cjk_punct") or char_kind(right) in ("cjk", "cjk_punct"):
            return ""
        if left in " " or right in " ":
            return ""
        return " "
    return _re_all(r"[\u00a0\u2009\u202f\u2007\u2008]", text, rep)


def i03_ideographic_space(text, ctx):
    """全角空格混入正文（行首的全角空格属于缩进问题，交给 B04）。"""
    out = []
    for m in re.finditer(r"\u3000+", text):
        if m.start() == 0:
            continue
        out.append((m.start(), m.end(), " ", None))
    return out


# ================================================================ S 层

def s01_cjk_num_space(text, ctx):
    """CJK 与数字之间的空格：`共 3 个` → `共3个`（preset=word-cn）。"""
    if not ctx.get("preset", "word-cn").startswith("word"):
        return []
    out = []
    out += _re_all(r"(?<=[%s])%s+(?=\d)" % (CJK, SP), text, lambda m: "")
    out += _re_all(r"(?<=\d)%s+(?=[%s])" % (SP, CJK), text, lambda m: "")
    return out


def s02_cjk_latin_space(text, ctx):
    """CJK 与拉丁字母/单词之间的空格：`使用 Python 编写` → `使用Python编写`。"""
    if not ctx.get("preset", "word-cn").startswith("word"):
        return []
    out = []
    out += _re_all(r"(?<=[%s])%s+(?=[A-Za-z])" % (CJK, SP), text, lambda m: "")
    out += _re_all(r"(?<=[A-Za-z])%s+(?=[%s])" % (SP, CJK), text, lambda m: "")
    return out


def s03_punct_space(text, ctx):
    """中文标点前后的空格（空格出现在标点外侧，是折行残留）。"""
    pat_l = r"(?<=[%s])%s+" % (re.escape(CJK_PUNCT), SP)
    pat_r = r"%s+(?=[%s])" % (SP, re.escape(CJK_PUNCT))
    out = []
    out += _re_all(pat_l, text, lambda m: "")
    out += _re_all(pat_r, text, lambda m: "")
    return out


def s04_multi_space(text, ctx):
    """连续多个半角空格（≥2）：合并为一个。段首的另算 B09。"""
    out = []
    for m in re.finditer(r" {2,}", text):
        if m.start() == 0 or text[:m.start()].strip(" \t\n") == "":
            continue
        out.append((m.start(), m.end(), " ", None))
    return out


def s05_bracket_pad(text, ctx):
    """括号内侧的填充空格：`（ 说明 ）` → `（说明）`。"""
    out = []
    out += _re_all(r"(?<=[（(])%s+" % SP, text, lambda m: "")
    out += _re_all(r"%s+(?=[）)])" % SP, text, lambda m: "")
    return out


def s06_num_unit_space(text, ctx):
    """数字与 %／‰／℃ 之间的空格：`50 %` → `50%`。"""
    return _re_all(r"(?<=\d)%s+(?=[%%‰℃°§])" % SP, text, lambda m: "")


def s07_para_edge_space(text, ctx):
    out = []
    for m in re.finditer(r"[ \t\u3000]+$", text):
        out.append((m.start(), m.end(), "", None))
    return out


def s08_cjk_cjk_space(text, ctx):
    """两个汉字之间的半角空格（段首/段尾交给 B04 与 S07）。

    `共 发现` 这种空格没有合法用途，是 Markdown 软折行粘进 Word 后留下的。
    两个汉字之间凭空空格就是 AI 痕迹——人写中文不会在汉字中间敲空格。
    """
    return _re_all(r"(?<=[%s]) +(?=[%s])" % (CJK, CJK), text, lambda m: "")


# ================================================================ P 层

_END_STOP = "。！？…”’）】」》"


def _zh_context(text, i, j, allow_digit_left=True):
    """判断 [i,j) 处的半角标点是否处在中文语境（用来决定要不要转全角）。"""
    left = text[i - 1] if i > 0 else ""
    right = text[j] if j < len(text) else ""
    lk, rk = char_kind(left), char_kind(right)
    if lk in ("cjk", "cjk_punct") or rk in ("cjk",):
        if not allow_digit_left and lk == "digit":
            return False
        return True
    if rk in ("cjk_punct",):
        return True
    return False


def p01_halfwidth_punct(text, ctx):
    """半角标点出现在中文语境 → 转全角。逐字符判断左右邻，不做全局替换。"""
    table = {",": "，", ".": "。", ";": "；", ":": "：", "?": "？", "!": "！"}
    out = []
    for i, ch in enumerate(text):
        if ch not in table:
            continue
        j = i + 1
        left = text[i - 1] if i > 0 else ""
        right = text[j] if j < len(text) else ""
        lk = char_kind(left)
        # 小数点、千分位、时间、缩写一律不动
        if ch in ".," and lk == "digit" and char_kind(right) == "digit":
            continue
        if ch == "." and (left.isalpha() and right.isalpha()):
            continue
        if ch == ":" and lk == "digit" and char_kind(right) == "digit":
            continue
        # `...` 是省略号，交给 P03；单个点参与连点也一律不碰
        if ch == "." and (left == "." or right == "."):
            continue
        if not _zh_context(text, i, j, allow_digit_left=(ch in ".?!")):
            continue
        if lk == "digit" and ch in ".,":
            continue
        out.append((i, j, table[ch], None))
    return out


def p02_straight_quote(text, ctx):
    """直引号 → 中文弯引号（仅当整段是中文语境，且引号成对出现）。"""
    out = []
    dq = [i for i, c in enumerate(text) if c == '"']
    if len(dq) % 2 == 0 and dq and "cjk" in [char_kind(text[i - 1]) for i in dq if i] + \
            [char_kind(text[i + 1]) for i in dq if i + 1 < len(text)]:
        for n, i in enumerate(dq):
            out.append((i, i + 1, "“" if n % 2 == 0 else "”", None))
    sq = [i for i, c in enumerate(text) if c == "'"]
    if len(sq) % 2 == 0 and len(sq) >= 2:
        for n, i in enumerate(sq):
            out.append((i, i + 1, "‘" if n % 2 == 0 else "’", None))
    return out


def p03_ellipsis(text, ctx):
    """省略号：`...`／`。。。`／单个 `…` → `……`（中文规范是两个）。"""
    out = []
    out += _re_all(r"\.{3,}", text, lambda m: "……")
    out += _re_all(r"。{3,}", text, lambda m: "……")
    for m in re.finditer(r"…+", text):
        if len(m.group(0)) == 1:
            out.append((m.start(), m.end(), "……", None))
    return out


def p04_dash(text, ctx):
    """双连字符破折号 → `——`。只认**恰好两个**连字符：`---` 是 Markdown 分隔线（M09），
    单个 `-` 是连字符/范围号，都不归它管。"""
    return _re_all(r"(?<!-)-{2}(?!-)", text, lambda m: "——")


def p05_repeat_punct(text, ctx):
    """叠标点：`！！`／`？？`／`。。` → 单个（省略号另由 P03 处理）。"""
    return _re_all(r"([，、；：？！])\1+", text, lambda m: m.group(1))


def p06_halfwidth_paren(text, ctx):
    """括号内侧是中文 → 换全角括号。被反斜杠转义的括号（M11）不归它管。"""
    out = []
    for m in re.finditer(r"\((.*?)\)", text):
        if m.start() > 0 and text[m.start() - 1] == "\\":
            continue
        inner = m.group(1)
        if not inner:
            continue
        cjk_n = sum(1 for c in inner if char_kind(c) in ("cjk", "cjk_punct"))
        if cjk_n >= max(1, len(inner) // 2):
            out.append((m.start(), m.start() + 1, "（", None))
            out.append((m.end() - 1, m.end(), "）", None))
    return out


def p07_fullwidth_alnum(text, ctx):
    """全角数字与字母 → 半角（全角字符会让检索、复制、比对全部失配）。"""
    table = {chr(0xFF10 + i): chr(ord("0") + i) for i in range(10)}
    table.update({chr(0xFF21 + i): chr(ord("A") + i) for i in range(26)})
    table.update({chr(0xFF41 + i): chr(ord("a") + i) for i in range(26)})
    out = []
    for i, ch in enumerate(text):
        if ch in table:
            out.append((i, i + 1, table[ch], None))
    return out


def p08_fullwidth_symbol(text, ctx):
    """全角符号 → 半角：％ → %，时间里的全角冒号 → :。"""
    out = []
    out += _re_all(r"％", text, lambda m: "%")
    for m in re.finditer(r"(?<=\d)：(?=\d)", text):
        out.append((m.start(), m.end(), ":", None))
    return out


def p09_tilde(text, ctx):
    """全角波浪号 `～` 出现在纯数值范围里 → 半角 `-`。"""
    return _re_all(r"(?<=\d)～(?=\d)", text, lambda m: "-")


# ================================================================ 规则表

RULES = [
    ("I01", "I", SEV_HARD, "零宽字符与控制符", "整段删除", i01_invisible),
    ("I02", "I", SEV_WARN, "不换行空格/细空格", "按左右字符决定删或换普通空格", i02_nbsp),
    ("I03", "I", SEV_WARN, "全角空格混入正文", "换为半角空格；段首的交给 B04", i03_ideographic_space),
    ("S01", "S", SEV_HARD, "CJK 与数字之间的空格", "删除（preset=word-cn）", s01_cjk_num_space),
    ("S02", "S", SEV_WARN, "CJK 与拉丁词之间的空格", "删除（preset=word-cn）", s02_cjk_latin_space),
    ("S03", "S", SEV_HARD, "中文标点前后的空格", "删除", s03_punct_space),
    ("S04", "S", SEV_WARN, "连续多个空格", "合并为一个", s04_multi_space),
    ("S05", "S", SEV_WARN, "括号内侧填充空格", "删除", s05_bracket_pad),
    ("S06", "S", SEV_WARN, "数字与 %/℃ 之间的空格", "删除", s06_num_unit_space),
    ("S07", "S", SEV_WARN, "段落收尾空白", "删除", s07_para_edge_space),
    ("S08", "S", SEV_WARN, "汉字之间的多余空格", "删除（Markdown 软折行搬进来的）", s08_cjk_cjk_space),
    # 顺序有意义：省略号(P03)必须排在半角标点(P01)之前，否则 `...` 会被 P01 先吃掉一个点
    ("P03", "P", SEV_HARD, "省略号写法不规范", "统一为 ……", p03_ellipsis),
    ("P01", "P", SEV_HARD, "中文语境使用半角标点", "转全角（逐字符判左右邻）", p01_halfwidth_punct),
    ("P02", "P", SEV_WARN, "直引号未转中文引号", "转 “ ” ‘ ’（成对）", p02_straight_quote),
    ("P04", "P", SEV_SUG, "破折号使用连字符", "`--` → `——`", p04_dash),
    ("P05", "P", SEV_WARN, "叠用标点", "合并为单个", p05_repeat_punct),
    ("P06", "P", SEV_SUG, "中文内容使用半角括号", "外侧换全角", p06_halfwidth_paren),
    ("P07", "P", SEV_HARD, "全角数字/字母", "转半角", p07_fullwidth_alnum),
    ("P08", "P", SEV_SUG, "全角符号", "％ → %；时间冒号 → :", p08_fullwidth_symbol),
    ("P09", "P", SEV_SUG, "数值范围使用全角波浪号", "→ 半角 `-`", p09_tilde),
]
RULE_META = {r[0]: dict(layer=r[1], severity=r[2], title=r[3], desc=r[4]) for r in RULES}

#: 默认不自动应用的规则（误伤风险高于收益，或需要人看上下文）
NO_AUTO = {"P02", "P04", "P06", "P08", "P09", "S02", "S04"}


# ---------------------------------------------------------------- 引擎

def scan_paragraph(para, preset="word-cn", rules=None):
    """对单个段落跑机械层规则，产出 Issue 列表。"""
    text = para.text
    if not text or not text.strip(" \t\n"):
        return []
    masks = masks_of(text)
    ctx = {"preset": preset, "para": para}
    taken = []           # 已占用的字符区间，避免两条规则改同一处
    issues = []
    protected = all(s.in_field_result or s.in_insertion for s in para.spans) if para.spans else False

    for rid, layer, sev, title, desc, fn in (rules or RULES):
        for s, e, repl, note in fn(text, ctx):
            if s >= e and not repl:
                continue
            if repl == text[s:e]:
                continue
            if _hits_mask(masks, s, e):
                continue
            if "\t" in text[s:e] or "\n" in text[s:e]:
                continue                      # 绝不跨哨兵
            if any(s < b and a < e for a, b in taken):
                continue                      # 与更高优先级的规则冲突
            span = next((sp for sp in para.spans if sp.char_start <= s < sp.char_end), None)
            if span is None:
                continue
            auto = (layer in AUTO_LAYERS) and rid not in NO_AUTO \
                and not span.in_field_result and not span.in_insertion
            taken.append((s, e))
            issues.append(Issue(
                rule=rid, layer=layer, severity=sev, part=para.part, pid=para.pid,
                anchor=_anchor(para), start=s, end=e, orig=text[s:e], repl=repl,
                note=note or desc, auto=auto, span=span))
    return issues


def _anchor(para) -> str:
    body = para.visible_text.strip()
    return body[:20] + ("…" if len(body) > 20 else "")


def apply_edits(xml: str, part: str, issues) -> str:
    """把 Issue 列表落回 XML 字节流（只改 `<w:t>` 内容）。"""
    from docx_text import rewrite_part
    edits = [(i.span, i.start, i.end, i.repl) for i in issues
             if i.auto and i.span is not None and i.span.part == part]
    if not edits:
        return xml
    return rewrite_part(xml, part, edits)
