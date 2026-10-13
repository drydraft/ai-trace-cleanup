#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ai-trace-cleanup · 共享层（零依赖，仅标准库）

把 .docx 拆成「可定位、可字节级回写」的文本单元。

设计要点（改动本文件前先读懂这四条）
------------------------------------
1. **不用 ElementTree 重新序列化整个包。** 只动 `<w:t>` 元素的**内容**，
   其余字节一律不动。这样「文档除了我们要改的字符之外完全没变」是能被
   diff 证明的，而不是靠信心。
2. **段落的文字被切碎在多个 run 里。** Word 会按格式变化、按编辑历史
   (rsid) 把一句话切成任意多段 `<w:r><w:t>`。所以必须建立「段落字符坐标
   → run 内部坐标」的映射，否则 `共 3 个` 这种跨 run 的空格永远删不掉。
3. **不可见元素要变成哨兵字符。** `<w:tab/>`/`<w:br/>`/`<w:cr/>` 不是文字但
   占据排版位置；在段落文本里映射成 `\\t`/`\\n`，并禁止任何规则跨越哨兵匹配
   （否则 `行尾。\\n下一行` 会被当成一句话误伤）。
4. **保护区是显式标记的，不是"我记得跳过"。** 域代码、修订痕迹、公式、图片
   各自带 flag，由调用方按 flag 决定动不动它。
"""

from __future__ import annotations

import re
import sys
import unicodedata
import zipfile
from dataclasses import dataclass, field

# ---------------------------------------------------------------- 基础工具

#: 正文文字所在的部件。header/footer 用前缀匹配。
MAIN_PART = "word/document.xml"
PART_PREFIXES = ("word/header", "word/footer")
PART_EXTRA = ("word/footnotes.xml", "word/endnotes.xml")
#: 默认不碰的部件（含批注：那是审阅者的意见，不是作者的正文）
PART_OPT_IN = ("word/comments.xml", "docProps/core.xml")


def u(text: str = "") -> str:
    """Windows 控制台默认 GBK，打印中文会炸。所有 CLI 都先调它。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    return text


def xml_unescape(s: str) -> str:
    """比标准库 unescape 多处理数字字符引用（Word 会写 &#xNNNN;）。"""
    if "&" not in s:
        return s
    s = re.sub(r"&#x([0-9A-Fa-f]+);", lambda m: chr(int(m.group(1), 16)), s)
    s = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), s)
    return (
        s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def xml_escape(s: str) -> str:
    # 非 ASCII 直接留原文，OOXML 是 UTF-8，Word 接受
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ---------------------------------------------------------------- 扫描器

#: 匹配一个 XML 标签。属性段用「引号内的整体」或「非 > 非引号字符」消费，
#: 因此属性值里出现 > 也不会把标签切错。
TAG_RE = re.compile(r"""<(/?)([A-Za-z_][\w.:\-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(/?)>""", re.S)
NON_TAG_RE = re.compile(r"<!--.*?-->|<!\[CDATA\[.*?\]\]>|<\?.*?\?>", re.S)

#: 段落文本里代表不可见元素的哨兵
SENTINELS = "".join(["\t", "\n"])
_SENTINEL_OWNER = {"w:tab": "\t", "w:br": "\n", "w:cr": "\n"}

#: 这些元素内部不产生「可见正文」，整段跳过
OPAQUE = {"w:instrText", "w:delText", "w:fldChar", "w:footnoteRef", "w:endnoteRef",
          "w:tblPr", "w:tcPr", "w:trPr", "w:sectPr",
          "w:tblGrid", "w:sdtPr", "w:sdtEndPr", "w:numPr", "w:framePr", "w:tcPrChange"}
# 注：`w:pPr` / `w:rPr` 不在此列——它们被单独处理（只读不遍历）。

_ROOT_OPEN = re.compile(r"<w:document\b")
_BODY_OPEN = re.compile(r"<w:body\b")


@dataclass
class Span:
    """一个 `<w:t>`：既记录它在 XML 字节流里的位置，也记录它在段落文本里的坐标。"""
    part: str
    tag_start: int          # '<w:t ...>' 起始
    tag_end: int            # '>' 之后一位
    body_start: int         # 元素内容起始（= tag_end 或自闭合时同）
    body_end: int           # 元素内容结束
    text: str               # 未转义文字（用作段落文本拼接源）
    char_start: int         # 在所属段落文本中的起点
    char_end: int
    in_field_result: bool = False
    in_insertion: bool = False
    bold: bool = False
    italic: bool = False
    fonts: tuple = ()
    ea_font: str = ""          # w:rFonts/@w:eastAsia —— 没设它，中文换机器就会变字体
    size: str = ""
    color: str = ""
    highlight: str = ""

    @property
    def empty(self) -> bool:
        return self.body_start == self.body_end


@dataclass
class Para:
    """一个 `<w:p>`（含文本框内嵌套段落，各自独立成 Para）。"""
    part: str
    pid: str
    start: int
    end: int
    spans: list = field(default_factory=list)     # List[Span]，文档顺序
    text: str = ""                                # 含哨兵的段落全文
    ppr: str = ""                                 # <w:pPr> 原文
    depth: int = 0

    # ---- pPr 派生属性 ----
    @property
    def style(self) -> str:
        m = re.search(r'<w:pStyle\s+w:val="([^"]*)"', self.ppr)
        return m.group(1) if m else ""

    @property
    def is_list(self) -> bool:
        return "<w:numPr" in self.ppr

    @property
    def jc(self) -> str:
        m = re.search(r'<w:jc\s+w:val="([^"]*)"', self.ppr)
        return m.group(1) if m else ""

    @property
    def spacing(self) -> str:
        m = re.search(r"<w:spacing\b([^>]*)>", self.ppr)
        return m.group(1).strip() if m else ""

    @property
    def manual_breaks(self) -> int:
        return self.text.count("\n")

    @property
    def visible_text(self) -> str:
        return self.text.replace("\t", " ").replace("\n", " ")

    def body(self) -> str:
        """去掉空白与哨兵后的纯文字，用于长度判断。"""
        return self.visible_text.strip()

    def first_span(self, default: int = 0) -> int:
        """用于只报段落级问题的伪字符坐标。"""
        return 0 if self.spans else default


@dataclass
class Doc:
    path: str
    order: list = field(default_factory=list)          # zip 条目顺序
    infos: dict = field(default_factory=dict)          # name -> ZipInfo
    data: dict = field(default_factory=dict)           # name -> bytes
    text_parts: list = field(default_factory=list)     # 我们要扫描的部件

    def xml(self, part: str) -> str:
        return self.data[part].decode("utf-8")

    def set_xml(self, part: str, text: str) -> None:
        self.data[part] = text.encode("utf-8")

    def is_text_part(self, name: str) -> bool:
        if name == MAIN_PART or name in PART_EXTRA:
            return True
        return name.endswith(".xml") and name.startswith(PART_PREFIXES)

    def protected_parts(self) -> list:
        """除文本部件外的一切条目：这些必须逐字节不变。"""
        return [n for n in self.order if not self.is_text_part(n)]


def read_docx(path: str) -> Doc:
    with zipfile.ZipFile(path) as z:
        order = [i.filename for i in z.infolist()]
        infos = {i.filename: i for i in z.infolist()}
        data = {n: z.read(n) for n in order}
    doc = Doc(path=path, order=order, infos=infos, data=data)
    if MAIN_PART not in data:
        raise SystemExit("这不是一个 .docx（缺少 word/document.xml）：%s" % path)
    doc.text_parts = [n for n in order if doc.is_text_part(n)]
    return doc


def write_docx(doc: Doc, out_path: str) -> None:
    """原顺序、原压缩方式、原时间戳写回；未登记的部件不动。"""
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in doc.order:
            info = doc.infos[name]
            zi = zipfile.ZipInfo(name, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zi.internal_attr = info.internal_attr
            zi.create_system = info.create_system
            z.writestr(zi, doc.data[name])


# ---------------------------------------------------------------- 段落切分

def iter_paragraphs(xml: str, part: str, start_pid: int = 0):
    """扫一遍 XML，产出 (Para, 下一个可用 pid)。

    用标签栈维护嵌套：`<w:p>` 在 `<w:txbxContent>` 里可以再套 `<w:p>`，
    所以不能拿 `<w:p>.*?</w:p>` 非贪婪匹配。
    """
    skip_ranges = [(m.start(), m.end()) for m in NON_TAG_RE.finditer(xml)]
    skip_at = [r[0] for r in skip_ranges]

    stack = []          # 元素名栈
    para_stack = []     # 当前打开的 Para（可嵌套）
    fields = []         # 域状态栈：'instr' | 'result'
    ins_depth = 0       # 修订插入深度
    run_stack = []      # 当前打开的 <w:r> 的 (start, end, props)
    pid_n = start_pid
    paras = []

    def in_skipped(pos):
        for s, e in skip_ranges:
            if s <= pos < e:
                return True
            if pos < s:
                break
        return False

    pos = 0
    for m in TAG_RE.finditer(xml):
        if in_skipped(m.start()):
            continue
        closing, name, _attrs, selfclose = m.group(1), m.group(2), m.group(3), m.group(4) == "/"

        if closing:
            # ---- 闭合 ----
            if name == "w:p" and para_stack:
                p = para_stack.pop()
                p.end = m.end()
                paras.append(p)
            elif name == "w:r" and run_stack:
                run_stack.pop()
            elif name == "w:ins":
                ins_depth = max(0, ins_depth - 1)
            elif name == "w:fldSimple" and fields:
                fields.pop()
            if stack and stack[-1] == name:
                stack.pop()
            pos = m.end()
            continue

        if selfclose:
            if name in _SENTINEL_OWNER and para_stack:
                para_stack[-1].text += _SENTINEL_OWNER[name]
            elif name == "w:b" and run_stack:
                run_stack[-1][2]["bold"] = _on(_attrs, True)
            elif name == "w:i" and run_stack:
                run_stack[-1][2]["italic"] = _on(_attrs, True)
            elif name == "w:u" and run_stack:
                run_stack[-1][2]["underline"] = _on(_attrs, False)
            elif name == "w:p":
                # 自闭合的空段落（<w:p/>）也要登记，否则「多余空段」永远看不见。
                # 注意不能加 `and para_stack`——顶层空段的 para_stack 本来就是空的。
                pid_n += 1
                para_stack.append(Para(part=part, pid="p%04d" % pid_n,
                                       start=m.start(), end=m.end(), depth=len(stack)))
                q = para_stack.pop()
                paras.append(q)
            elif name == "w:t" and para_stack:
                _emit_span(para_stack[-1], part, m.start(), m.end(), m.end(), "",
                           props=(run_stack[-1][2] if run_stack else None))
            elif name == "w:fldChar":
                _fld_state(_attrs, fields)
            pos = m.end()
            continue

        # ---- 开标签 ----
        if name in OPAQUE:
            # 需要成对跳过的：找到它的闭合，整段忽略（不影响外层正文）
            if name in ("w:instrText", "w:delText"):
                # 域代码与修订删除的文字：不是可见正文，整段不参与清洗
                close = xml.find("</%s>" % name, m.end())
                if close != -1:
                    pos = close + len(name) + 3
                    continue
            stack.append(name)
            pos = m.end()
            continue

        if name == "w:t":
            if para_stack:
                close = xml.find("</w:t>", m.end())
                if close == -1:
                    close = m.end()
                body = xml[m.end():close]
                _emit_span(para_stack[-1], part, m.start(), m.end(), m.end(), xml_unescape(body),
                           body_end=close,
                           in_field_result=bool(fields and fields[-1] == "result"),
                           in_insertion=ins_depth > 0,
                           props=(run_stack[-1][2] if run_stack else None))
                pos = close + len("</w:t>")
                continue
            pos = m.end()
            continue

        if name == "w:r":
            run_stack.append((m.start(), m.end(), {}))
        elif name == "w:p":
            pid_n += 1
            para_stack.append(Para(part=part, pid="p%04d" % pid_n, start=m.start(),
                                   end=m.end(), depth=len(stack)))
        elif name == "w:ins":
            ins_depth += 1
        elif name == "w:fldChar":
            _fld_state(_attrs, fields)
        elif name == "w:fldSimple":
            fields.append("result")
        elif name == "w:rPr":
            if run_stack:
                close = xml.find("</w:rPr>", m.end())
                if close != -1:
                    _read_rpr(run_stack[-1][2], xml[m.end():close])
                    pos = close + len("</w:rPr>")
                    continue
        elif name == "w:pPr":
            if para_stack:
                close = xml.find("</w:pPr>", m.end())
                if close != -1:
                    para_stack[-1].ppr = xml[m.end():close]
                    pos = close + len("</w:pPr>")
                    continue

        stack.append(name)
        pos = m.end()

    # 未闭合的段落（异常文档）也收回来，别静默丢
    while para_stack:
        p = para_stack.pop()
        p.end = len(xml) if not p.end or p.end < p.start else p.end
        paras.append(p)

    paras.sort(key=lambda p: p.start)
    for p in paras:
        p.spans.sort(key=lambda s: s.char_start)
    return paras, pid_n


def _on(attrs: str, default: bool) -> bool:
    m = re.search(r'w:val="([^"]*)"', attrs or "")
    if not m:
        return default
    return m.group(1).lower() not in ("0", "false", "off", "none")


def _fld_state(attrs: str, fields: list) -> None:
    m = re.search(r'w:fldCharType="([^"]*)"', attrs or "")
    kind = m.group(1) if m else ""
    if kind == "begin":
        fields.append("instr")
    elif kind == "separate":
        if fields:
            fields[-1] = "result"
    elif kind == "end":
        if fields:
            fields.pop()


def _read_rpr(props: dict, rpr: str) -> None:
    props["bold"] = props.get("bold") or bool(re.search(r"<w:b\s*/>|<w:b\s+w:val=\"(?:1|true|on)\"", rpr))
    props["italic"] = props.get("italic") or bool(re.search(r"<w:i\s*/>|<w:i\s+w:val=\"(?:1|true|on)\"", rpr))
    fonts = []
    for key in ("w:ascii", "w:eastAsia", "w:hAnsi", "w:cs"):
        f = re.search(r'<w:rFonts\b[^>]*%s="([^"]*)"' % key, rpr)
        if f:
            fonts.append(f.group(1))
    if fonts:
        props["fonts"] = tuple(fonts)
    ea = re.search(r'<w:rFonts\b[^>]*w:eastAsia="([^"]*)"', rpr)
    if ea:
        props["ea_font"] = ea.group(1)
    sz = re.search(r'<w:sz\s+w:val="([^"]*)"', rpr)
    if sz:
        props["size"] = sz.group(1)
    color = re.search(r'<w:color\s+w:val="([^"]*)"', rpr)
    if color:
        props["color"] = color.group(1)
    hi = re.search(r'<w:highlight\s+w:val="([^"]*)"', rpr)
    if hi:
        props["highlight"] = hi.group(1)


def _emit_span(para: Para, part: str, tag_start: int, tag_end: int, b0: int, text: str,
               body_end=None, in_field_result: bool = False, in_insertion: bool = False,
               props: dict = None) -> None:
    """把一个 `<w:t>` 登记进段落；tag_start/tag_end 用于回写时改 xml:space。"""
    props = props or {}
    cs = len(para.text)
    para.text += text
    para.spans.append(Span(
        part=part, tag_start=tag_start, tag_end=tag_end,
        body_start=b0, body_end=(b0 if body_end is None else body_end),
        text=text, char_start=cs, char_end=cs + len(text),
        in_field_result=in_field_result, in_insertion=in_insertion,
        bold=bool(props.get("bold")), italic=bool(props.get("italic")),
        fonts=tuple(props.get("fonts") or ()), size=props.get("size") or "",
        ea_font=props.get("ea_font") or "",
        color=props.get("color") or "", highlight=props.get("highlight") or "",
    ))


# ---------------------------------------------------------------- 回写

def rewrite_part(xml: str, part: str, edits) -> str:
    """把「段落字符坐标」的替换落回 XML，只改 `<w:t>` 元素内容。

    edits: [(span, s, e, repl)]  —— s/e 是段落字符坐标，且必须落在同一个 Span 内。
    返回新的 XML 字符串。
    """
    by_span = {}
    for span, s, e, repl in edits:
        by_span.setdefault(id(span), (span, []))[1].append((s, e, repl))

    xml_edits = []
    for _sid, (span, items) in by_span.items():
        items.sort()
        new_body, _ = _rewrite_text(span.text, span.char_start, items)
        if new_body == span.text:
            continue
        tag = xml[span.tag_start:span.tag_end]
        needs = new_body != new_body.strip() or new_body.startswith(" ") or new_body.endswith(" ")
        if needs and 'xml:space="preserve"' not in tag:
            newtag = tag[:-1].rstrip() + ' xml:space="preserve">'
            if tag.endswith("/>"):
                newtag = tag[:-2].rstrip() + ' xml:space="preserve"/>'
            xml_edits.append((span.tag_start, span.tag_end, newtag))
        body = xml_escape(new_body)
        xml_edits.append((span.body_start, span.body_end, body))

    for s, e, r in sorted(xml_edits, key=lambda x: -x[0]):
        xml = xml[:s] + r + xml[e:]
    return xml


def _rewrite_text(text: str, cs: int, items) -> tuple:
    """按 (s,e,repl) 重写一段文字；跨到本段之外的 edit 由本段负责拼接。"""
    ce = cs + len(text)
    out, pos = [], 0
    changed = False
    for s, e, repl in items:
        if e <= cs or s >= ce:
            continue
        ls = max(s, cs) - cs
        le = min(e, ce) - cs
        out.append(text[pos:ls])
        if cs <= s < ce:
            out.append(repl)
        pos = le
        changed = True
    out.append(text[pos:])
    return "".join(out), changed


# ---------------------------------------------------------------- 定位辅助

#: 看不见的字符在报告里用这个名字显示，否则「命中了」但人眼看不到任何东西。
INVISIBLE_NAMES = {
    "\u200b": "ZWSP", "\u200c": "ZWNJ", "\u200d": "ZWJ", "\u2060": "WJ",
    "\ufeff": "BOM", "\u00ad": "SHY", "\u00a0": "NBSP", "\u2009": "THIN",
    "\u202f": "NNBSP", "\u2007": "FIGSP", "\u2008": "PUNCTSP", "\u3000": "IDSP",
    "\u202a": "LRE", "\u202b": "RLE", "\u202c": "PDF", "\u202d": "LRO",
    "\u202e": "RLO", "\u2066": "LRI", "\u2067": "RLI", "\u2068": "FSI", "\u2069": "PDI",
}


def _esc_inline(x: str) -> str:
    """把控制字符写成可读的转义形式。为什么不直接展开：报告里出现真换行会截断整行。"""
    return x.replace("\t", "\\t").replace("\n", "\\n")


def _vis_hit(x: str) -> str:
    """把「命中的那几个字符」渲染成人能看见的东西。

    删一个零宽字符、报一行「零宽字符« »夹在中间」，读者是看不到任何东西的。
    所以隐形字符一律换成名字，实在没名字的就写码位。
    """
    out = []
    for ch in x:
        if ch in INVISIBLE_NAMES:
            out.append(INVISIBLE_NAMES[ch])
        elif ch == "\t":
            out.append("\\t")
        elif ch == "\n":
            out.append("\\n")
        elif unicodedata.category(ch) in ("Cf", "Cc", "Zs") and ch != " ":
            out.append("U+%04X" % ord(ch))
        else:
            out.append(ch)
    return "".join(out)


def snippet(para: Para, s: int, e: int, pad: int = 12) -> str:
    """把命中的那几个字符用 «» 框起来。

    为什么不用 ⟪⟫（U+27EA/U+27EB）：它在等宽字体里普遍缺字形，Consolas 会回退到
    中日韩字体把它画成全宽括号，读起来像书名号《》。«» 是 Latin-1，任何字体都有。
    """
    t = para.text
    a = max(0, s - pad)
    b = min(len(t), e + pad)
    head = "…" if a > 0 else ""
    tail = "…" if b < len(t) else ""
    return "%s%s«%s»%s%s" % (head, _esc_inline(t[a:s]), _vis_hit(t[s:e]),
                              _esc_inline(t[e:b]), tail)


_CONTROL_RE = re.compile(r"[\u200b-\u200f\u2060\ufeff\u00ad\u202a-\u202e\u2066-\u2069]")


def char_kind(ch: str) -> str:
    import unicodedata
    if ch in "\t\n":
        return "sentinel"
    o = ord(ch)
    if 0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF or 0xF900 <= o <= 0xFAFF:
        return "cjk"
    if ch in "，。、；：？！“”‘’（）《》〈〉【】「」『』—…·～":
        return "cjk_punct"
    if unicodedata.east_asian_width(ch) in ("W", "F") and not ch.isascii():
        return "cjk"
    if ch.isdigit():
        return "digit"
    if ch.isalpha():
        return "latin" if ch.isascii() else "other"
    if ch in ",.;:?!()[]{}<>\"'`~-/%$&*+=@#|\\^_":
        return "ascii_punct"
    if ch.isspace():
        return "space"
    return "other"


def tokens(text: str) -> set:
    """语义护栏用的 token 集合：数字、拉丁词、URL、代码标识符。

    只要清洗前后这个集合不变，就可以断言「没有改到事实性内容」。

    先做 NFKC 归一化：本 skill 的规则**故意**要把全角数字/字母/％ 转成半角，
    那是我们要的结果，不能在护栏里被误判成「改了事实」，所以护栏只关心
    「这个数字/这个词还在不在、值变没变」。
    """
    import unicodedata
    text = unicodedata.normalize("NFKC", text)
    out = set()
    for m in re.finditer(r"https?://[^\s，。、；）】」》，)>\]]+", text):
        out.add("url:" + m.group(0).rstrip(".,;"))
    body = re.sub(r"https?://\S+", " ", text)
    for m in re.finditer(r"\d+(?:[.,]\d+)*\s*(?:%|‰)?", body):
        tok = m.group(0).replace(" ", "").replace("\u3000", "")
        out.add("num:" + tok.replace(",", ""))
    for m in re.finditer(r"[A-Za-z][A-Za-z0-9_\-]*", body):
        w = m.group(0)
        if len(w) >= 2 or w.isupper():
            out.add("lat:" + w.lower())
    for m in re.finditer(r"[A-Za-z_][A-Za-z0-9_]*\.[A-Za-z0-9_.]+|[a-z]+_[a-z_]+", body):
        out.add("id:" + m.group(0))
    return out


def para_anchor(para: Para, header: str = "") -> str:
    label = header or para.style or "正文"
    return "%s §%s [%s]" % (label, para.pid, (para.body()[:18] + "…") if len(para.body()) > 18 else para.body())
