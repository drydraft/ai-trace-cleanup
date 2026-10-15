#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 tests/sample-dirty.docx —— 一份「AI 输出痕迹大全」的脏样本。

它同时承担两个角色：
1. 开发期的回归样本（golden sample）：改规则后跑 scan，命中集必须只增不减；
2. 校验器的靶子：清理后 guard 必须报「保护区逐字节未变」。

样本故意做得「像 Word 真的写出来的」：一句话被拆到多个 run、
带 xml:space 的首尾空格、域代码、修订痕迹、文本框、表格。
"""
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.docx_text import u, xml_escape  # noqa: E402

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
NS_DOC = '%s %s xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' % (W, R)


def t(text, **rpr):
    """一个 run。text 的首尾空格会自动配 xml:space，和 Word 的行为一致。"""
    props = ""
    if rpr:
        bits = ""
        if rpr.get("bold"):
            bits += "<w:b/>"
        if rpr.get("italic"):
            bits += "<w:i/>"
        if rpr.get("color"):
            bits += '<w:color w:val="%s"/>' % rpr["color"]
        if rpr.get("size"):
            bits += '<w:sz w:val="%s"/>' % rpr["size"]
        if rpr.get("font"):
            # only_ascii=True 模拟「只设了西文字体，没设 w:eastAsia」的粘贴残留（F09）
            if rpr.get("only_ascii"):
                bits += '<w:rFonts w:ascii="%s" w:hAnsi="%s"/>' % (rpr["font"], rpr["font"])
            else:
                bits += ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
                         % (rpr["font"], rpr["font"], rpr.get("ea", "宋体")))
        elif rpr.get("only_ascii"):
            bits += '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        props = "<w:rPr>%s</w:rPr>" % bits
    sp = ' xml:space="preserve"' if text != text.strip() else ""
    return "<w:r>%s<w:t%s>%s</w:t></w:r>" % (props, sp, xml_escape(text))


def p(*runs, **kw):
    ppr = ""
    if kw.get("style"):
        ppr += '<w:pStyle w:val="%s"/>' % kw["style"]
    if kw.get("jc"):
        ppr += '<w:jc w:val="%s"/>' % kw["jc"]
    if kw.get("num"):
        ppr += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
    if kw.get("spacing"):
        ppr += '<w:spacing %s/>' % kw["spacing"]
    if kw.get("rpr"):
        ppr += "<w:rPr>%s</w:rPr>" % kw["rpr"]
    body = "".join(runs) if runs else ""
    head = "<w:pPr>%s</w:pPr>" % ppr if ppr else ""
    return "<w:p>%s%s</w:p>" % (head, body)


def br():
    return "<w:r><w:br/></w:r>"


def table(rows):
    grid = "<w:tblGrid>%s</w:tblGrid>" % ("<w:gridCol w:w=\"2400\"/>" * len(rows[0]))
    out = ["<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr>", grid]
    for row in rows:
        out.append("<w:tr>")
        for cell in row:
            out.append('<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>%s</w:tc>' % cell)
        out.append("</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def field_result(instr, result_runs):
    """一个带结果的域（如 TOC / PAGE），结果文字属于保护区。"""
    return ("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            "<w:r><w:instrText xml:space=\"preserve\">%s</w:instrText></w:r>"
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            "%s"
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>") % (instr, "".join(result_runs))


def build_body():
    b = []
    a = b.append

    # ── 标题（含 数字+空格）
    a(p(t("2026 年第三季度巡检工作总结", size="32"), style="Heading1"))
    a(p(t("编制：运维组", size="21"), style="Heading2"))

    # ── ★ 关键用例：一句话被切成 3 个 run，" 3 " 的空格在独立的 run 里
    a(p(t("本次巡检共"), t(" 发现 "), t("3 个问题"), t("，已闭环 2 项。")))

    # ── Markdown 残留（AI 直接粘贴进 Word 的头号痕迹）
    a(p(t("**重点**：以下为问题清单说明")))
    a(p(t("## 二、主要问题")))
    a(p(t("- 设备老化，风扇异响")))
    a(p(t("- 网络抖动，丢包率 3 %")))
    a(p(t("1. 机房温湿度超标")))
    a(p(t("2. UPS 电池组告警")))
    a(p(t("详见 `JIRA-1234` 与 [巡检报告](https://example.com/report?id=9)")))
    a(p(t("| 项 | 值 |")))
    a(p(t("|---|---|")))
    a(p(t("---")))
    a(p(t("> 这是一段被当作引用块的文字")))
    a(p(t(r"转义残留：\* 和 \_ 和 \(括号\)")))
    a(p(t("HTML 残留：&nbsp; 与 <br> 标签")))

    # ── 标点层
    a(p(t("检查完成, 未发现异常.")))
    a(p(t("所谓\"闭环\"在这里被滥用")))
    a(p(t("省略号... 还有。。。两种写法")))
    a(p(t("数量 10-20 台，另一处写作 A--B")))
    a(p(t("发现了严重问题！！为什么呢？？")))
    a(p(t("括号(半角) 和 全角（混用）")))
    a(p(t("占比 50 ％，全角数字１２３")))
    a(p(t("时间 9：30 使用了全角冒号")))

    # ── 空白 / 不可见字符
    a(p(t("零宽字符\u200b夹在中间，不断行空格\u00a0也来了")))
    a(p(t("全角空格\u3000混入，软连字符\u00ad")))
    a(p(t("括号内空格（ 说明文字 ）与尾随空格 ")))
    a(p(t("数值 5 千米，单位 10 kg，百分比 50 %")))

    # ── 换行 / 段落结构
    a(p(t("这一行是人工换行"), br(), t("这一行接在后面")))
    a("<w:p/>")
    a("<w:p/>")
    a(p(t("第一句很短。")))
    a(p(t("第二句也很短。")))
    a(p(t("第三句还是短。")))
    a(p(t("第四句依然短。")))
    a(p(t("第五句照样短。")))
    a(p(t("第六句同样短。")))

    # ── 格式 / 样式层：字体漂移、字号漂移、彩色、加粗滥用
    a(p(t("正文应当是宋体", font="宋体", size="21"),
        t("这里被粘贴成 Calibri", font="Calibri", size="24"),
        t("字号又跳回小四", font="宋体", size="21")))
    a(p(t("红色强调", color="FF0000", size="21"), t("普通文字", size="21")))
    a(p(t("整段加粗", bold=True, size="21"), t("继续加粗", bold=True, size="21"),
        t("还在加粗", bold=True, size="21"), t("依然加粗", bold=True, size="21")))
    a(p(t("左对齐段落", size="21"), jc="left"))
    a(p(t("两端对齐段落", size="21"), jc="both"))

    # ── 语言层 AI 味
    a(p(t("首先，我们需要明确目标。其次，要拆解任务。最后，形成闭环。")))
    a(p(t("值得注意的是，这一现象并非偶然。")))
    a(p(t("综上所述，我们可以得出结论。")))
    a(p(t("这不仅提升了效率，而且降低了成本。")))
    a(p(t("在当今快速发展的时代，数字化转型势在必行。")))
    a(p(t("让我们一起来看下面这组数据。")))
    a(p(t("该方案极大地促进了业务的发展，能够有效地提升整体效能。")))
    a(p(t("通过赋能一线团队，打造数据驱动的生态体系，提升管理颗粒度。")))
    a(p(t("不是设备的问题，而是流程的问题。")))
    a(p(t("需要注意的是，以上仅为初步判断。")))

    # ── 内容层
    a(p(t("作为一个 AI 语言模型，我无法直接访问机房。")))
    a(p(t("以上是我为你整理的内容，希望对你有帮助。如果需要进一步分析请告诉我。")))
    a(p(t("以上内容仅供参考，具体请以现场实际情况为准。")))
    a(p(t("本节内容待补充：[待补充]")))
    a(p(t("综上所述，本文围绕巡检工作展开了系统梳理，从设备、网络、环境三个维度进行了分析，"
          "并结合实际情况提出了相应的改进建议，具有一定的参考价值。")))

    # ── 文本级保护区：路径 / 命令 / 代码段里的点号与标点绝不能按中文规则改。
    #    每句都故意在同段里放一个**该被修的**问题，用来证明「路径没动 + 相邻的修了」。
    a(p(t("保护路径 .github/workflows/ 共 3 个文件。"), size="21"))
    a(p(t(r"路径 C:\Users\Administrator\.pi\agent\skills\ 下共 2 个目录。"), size="21"))
    a(p(t("命令 python scripts/docx_scan.py --all 与版本 v0.1.0 都不许改，但共 3 处空格要删。"),
          size="21"))
    a(p(t("时间 9:30、单位 10 kg、邮箱 a@b.com、代码 `x = 1` 都不能动。"), size="21"))

    # ── 补齐样本覆盖面：让每条规则都至少有一次真命中（回归时可检出规则失效）
    a(p(t("　　使用全角空格做首行缩进，这是伪缩进")))                       # B04
    a(p(t("　另外一段也是用空格缩进的正文")))                              # B04
    a(p(t("这一段故意写得很长，用来触发「单段过长」规则。" + "巡检过程中，我们逐台核对了设备的运行状态、"
          "风扇转速、温度传感器读数与网络接口的丢包统计，并把异常项逐一登记到台账里；" * 6,
          size="21")))                                                        # B06 (>300 字)
    a(p(t("这里有两个　空格在中间。", size="21")))
    a(p(t("这里有两个  半角空格在中间。", size="21")))                      # S04
    a(p(t("括号里( 半角填充空格 )也要清。", size="21")))                    # S05
    a(p(t("温度应在 3～5 度之间。", size="21")))                            # P09
    a(p(t("*这里是字面斜体标记*，应当改为 Word 斜体。", size="21")))       # M02
    a(p(t("```", size="21")))                                               # M06
    a(p(t("下面我将从三个方面展开说明。", size="21")))                      # C07
    a(p(t("本月进展——设备台账已更新——网络拓扑已重绘——环境监测已接入——"
          "告警规则已收敛——巡检路线已优化——值班表已排定——备件库存已盘点——"
          "机房平面图已归档——线缆标签已补齐——UPS 测试已完成——空调滤网已更换。",
          size="21")))                                                       # L11 破折号密度
    a(p(t("下面这三段用手工间距设段距，没有用样式。", size="21"),
        spacing='w:before="120" w:after="120"'))                              # F07
    a(p(t("第二段也是手工段距。", size="21"), spacing='w:before="120" w:after="120"'))
    a(p(t("第三段还是手工段距。", size="21"), spacing='w:before="120" w:after="120"'))
    for _ in range(4):
        a(p(t("需要重点关注设备、网络、环境三个方面。", size="21")))       # L10 三项排比密度
    a(p(t("这里只设了西文字体没有设中文字体，换台电脑字体就会变。", font="Calibri",
          only_ascii=True, size="21")))                                       # F09
    a(p(t("这一段只设了西文字体。", only_ascii=True, size="21")))          # F09
    a(p(t("又一段只设了西文字体。", only_ascii=True, size="21")))          # F09
    a(p(t("这段是整段加粗的长文本，用来把全文加粗率推上去。"
          "运维文档里到处加粗等于没有重点，读者会自动忽略所有黑体。"
          "所以整段加粗本身就是一种 AI 输出痕迹，应当拆成个别关键词加粗。"
          "这里再补一些文字让加粗字符数超过全文的百分之十五。"
          "同时把同样的道理再多说几遍，因为重复也是 AI 输出的典型特征。"
          "重点太多等于没有重点，强调太多等于没有强调，信息密度反而下降。"
          "文档应当只对真正需要记住的数字和结论加粗。"
          "继续堆加粗文字，把比例推到阈值以上。"
          "这一句也是加粗的，用来验证加粗率规则能不能真的触发。"
          "再加一句，确保万无一失。",
          bold=True, size="21")))                                             # F05
    a(p(t("第二段整段加粗的正文，继续把全文加粗率推高。"
          "真实情况是：AI 生成后用 Word 打开时，几乎每一个关键名词都被加粗了，"
          "因为它把「重要」理解成了「要加粗」。人写文档时加粗量通常不到百分之五。"
          "所以加粗率是一个非常好用的可量化指标。",
          bold=True, size="21")))                                             # F05（第二例）
    a(p(t("人工标题没用样式", bold=True, size="32")))                        # F08
    a(p(t("巡检已经完成，现场无异常。", size="21")))
    a(p(t("补充说明：巡检已经完成，现场无异常。", size="21")))             # C05 重复句
    a(p(t("本段落结束时的标点被吃掉了，最后几个汉字", size="21")))        # B05

    # ── 表格（单元格里的段落也要被扫到）
    a(table([
        [p(t("项目")), p(t("共 3 个问题"))],
        [p(t("状态")), p(t("已处理 2 个,剩余 1 个"))],
    ]))

    # ── 域：结果文字属于保护区，应报告但不自动改
    a(p(t("目录项示例："), field_result('TOC \\o "1-3" \\h \\z \\u',
                                     [t("第一章 概述\t3")])))

    # ── 修订痕迹：ins 记入正文但标记，del 完全跳过
    a("<w:p><w:ins w:id=\"1\" w:author=\"A\" w:date=\"2026-10-01T00:00:00Z\">"
      + t("这是修订插入的文字 共 3 处") + "</w:ins>"
      + "<w:del w:id=\"2\" w:author=\"A\" w:date=\"2026-10-01T00:00:00Z\">"
      + "<w:r><w:delText>这是被删除的文字</w:delText></w:r></w:del></w:p>")

    # ── 文本框（嵌套段落）
    a('<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml">'
      '<v:textbox><w:txbxContent>'
      + p(t("文本框里也有 3 个空格问题"))
      + "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>")

    return "".join(b)


SECT = ('<w:sectPr><w:headerReference w:type="default" r:id="rId2"/>'
        '<w:footerReference w:type="default" r:id="rId3"/>'
        '<w:pgSz w:w="11906" w:h="16838"/></w:sectPr>')


def build():
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<w:document %s><w:body>%s%s</w:body></w:document>'
                % (NS_DOC, build_body(), SECT))
    header = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<w:hdr %s>%s</w:hdr>' % (NS_DOC, p(t("巡检工作总结 共 5 页"))))
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<w:ftr %s>%s</w:ftr>' % (NS_DOC, p(field_result("PAGE",
                                                               [t("1")]))))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<w:styles %s>'
              '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
              '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
              '<w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>'
              '</w:styles>' % W)
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                     '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                     '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
                     '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                     '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                     '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '</Relationships>')
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
                '</Relationships>')
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/">'
            '<dc:title>2026 年第三季度巡检工作总结</dc:title><dc:creator>运维组</dc:creator>'
            '</cp:coreProperties>')
    return [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", rels),
        ("word/document.xml", document),
        ("word/_rels/document.xml.rels", doc_rels),
        ("word/styles.xml", styles),
        ("word/header1.xml", header),
        ("word/footer1.xml", footer),
        ("docProps/core.xml", core),
    ]


def main(out=None):
    out = out or os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample-dirty.docx")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in build():
            # 固定时间戳：样本是被提交进仓库的回归基准，每次跑测试都把它改成
            # 「已修改」是不能接受的（zip 头里的 mtime 会变）。
            info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            z.writestr(info, text.encode("utf-8"))
    u()
    print("已生成脏样本：%s" % out)
    return out


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)
