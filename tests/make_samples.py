#!/usr/bin/env python3
"""Generate sample ``.docx`` files used by the test suite and manual checks.

Run directly::

    python3 tests/make_samples.py [output_dir]

It writes ``sample_cn.docx`` (Chinese, mixed formatting, table, image, lists,
header/footer) and ``sample_en.docx`` (English, justified text, table).
"""

from __future__ import annotations

import io
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #


def add_numbering(doc: Document, num_id: int, abstract_id: int, fmt: str,
                  texts, start: int = 1, left: int = 720, hanging: int = 360,
                  symbol_font: bool = False) -> None:
    """Add an abstractNum + num pair to ``numbering.xml``."""
    numbering = doc.part.numbering_part.element
    rpr = ('<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>'
           '</w:rPr>') if symbol_font else ""
    levels = []
    for ilvl, text in enumerate(texts):
        levels.append(
            f'<w:lvl w:ilvl="{ilvl}">'
            f'<w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{left + ilvl * 360}" w:hanging="{hanging}"/></w:pPr>'
            f'{rpr}'
            f'</w:lvl>')
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">'
        + "".join(levels) + "</w:abstractNum>")
    numbering.insert(0, abstract)
    numbering.append(parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/></w:num>'))


def apply_numbering(paragraph, num_id: int, ilvl: int = 0) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    numPr = parse_xml(
        f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{ilvl}"/>'
        f'<w:numId w:val="{num_id}"/></w:numPr>')
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is not None:
        pStyle.addnext(numPr)
    else:
        pPr.insert(0, numPr)


def shade_cell(cell, fill: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:fill="{fill}"/>'))


def add_field(paragraph, instr: str) -> None:
    paragraph._p.append(parse_xml(
        f'<w:fldSimple {nsdecls("w")} w:instr=" {instr} ">'
        f'<w:r><w:t>1</w:t></w:r></w:fldSimple>'))


def make_png(width: int = 640, height: int = 320) -> bytes:
    """A small bar-chart-like PNG (no external text rendering needed)."""
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    draw.rectangle([0, 0, width - 1, height - 1], outline=(80, 80, 80), width=2)
    bars = [(40, 120, (66, 133, 244)), (160, 210, (219, 68, 55)),
            (280, 160, (244, 180, 0)), (400, 260, (15, 157, 88)),
            (520, 90, (171, 71, 188))]
    for x, bar_height, color in bars:
        draw.rectangle([x, height - 40 - bar_height, x + 80, height - 40], fill=color)
    draw.line([20, height - 40, width - 20, height - 40], fill=(60, 60, 60), width=2)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


# --------------------------------------------------------------------------- #
# samples
# --------------------------------------------------------------------------- #


def build_chinese(path: str) -> None:
    doc = Document()

    # page setup: A4 with 2.54cm margins is the default; make it explicit
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)

    # numbering definitions: decimal, bullet, Chinese counting
    add_numbering(doc, num_id=50, abstract_id=90, fmt="decimal", texts=["%1.", "%2."])
    add_numbering(doc, num_id=51, abstract_id=91, fmt="bullet",
                  texts=["\uf0b7", "o"], left=720, hanging=360, symbol_font=True)
    add_numbering(doc, num_id=52, abstract_id=92, fmt="chineseCountingThousand",
                  texts=["%1、"], left=720, hanging=360)

    doc.add_heading("季度业务报告 Quarterly Report", level=0)

    intro = doc.add_paragraph()
    intro.add_run("本报告总结了 ").bold = False
    run = intro.add_run("2024 年第三季度")
    run.bold = True
    intro.add_run(" 的运营情况，包含 ")
    run = intro.add_run("Revenue")
    run.italic = True
    intro.add_run("、成本与客户满意度等指标。Mixed English and 中文 text line. ")
    run = intro.add_run("重点结论")
    run.bold = True
    run.underline = True
    run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
    intro.add_run("：整体表现稳中向好。")

    doc.add_heading("一、核心指标 Key Metrics", level=1)
    doc.add_paragraph(
        "营业收入为 1,234 万元，同比增长 18.6%；毛利率提升 2.1 个百分点。"
        "The quick brown fox jumps over the lazy dog, 0123456789.")

    doc.add_heading("1.1 要点摘要", level=2)
    for text in ("华东区域增长最快，贡献 42% 收入。",
                 "客户续约率达到 91%，创历史新高。",
                 "供应链成本下降 3.4%。"):
        paragraph = doc.add_paragraph(style="List Bullet")
        paragraph.add_run(text)
        apply_numbering(paragraph, 51)

    doc.add_heading("1.2 待办事项", level=2)
    for text in ("完成第四季度预算评审。", "上线新的客户管理系统。", "组织两次行业交流活动。"):
        paragraph = doc.add_paragraph()
        paragraph.add_run(text)
        apply_numbering(paragraph, 50)

    doc.add_heading("1.3 分阶段目标", level=2)
    for text in ("第一阶段：渠道梳理。", "第二阶段：产品迭代。", "第三阶段：规模化推广。"):
        paragraph = doc.add_paragraph()
        paragraph.add_run(text)
        apply_numbering(paragraph, 52)

    doc.add_heading("二、数据明细 Details", level=1)
    table = doc.add_table(rows=4, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["区域 Region", "收入（万元）", "同比", "备注"]
    for index, header in enumerate(headers):
        cell = table.cell(0, index)
        cell.text = header
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade_cell(cell, "D9E2F3")
    rows = [["华东", "520", "+24%", "增长最快"],
            ["华南", "330", "+11%", "稳定"],
            ["华北", "384", "+17%", "新客户较多"]]
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            table.cell(r, c).text = value
    # merged cell spanning two columns on the last row
    merged = table.cell(3, 0).merge(table.cell(3, 1))
    merged.text = "合计 Total"
    for paragraph in merged.paragraphs:
        for run in paragraph.runs:
            run.bold = True
    table.cell(3, 2).text = "1234"
    table.cell(3, 3).text = "+18.6%"

    doc.add_paragraph()
    doc.add_picture(io.BytesIO(make_png()), width=Cm(12))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption = doc.add_paragraph("图 1  各区域收入对比 / Revenue by region")
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in caption.runs:
        run.font.size = Pt(9)
        run.italic = True

    # page break + second page
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    doc.add_heading("三、附录 Appendix", level=1)
    doc.add_paragraph(
        "附录包含术语表与统计数据来源说明。Symbols: → ★ ① ± ½ © ® ™ 中文标点：，。！？"
        "（括号）【方括号】“引号”。")
    quote = doc.add_paragraph("这是引用段落，用于检查缩进与对齐效果。" * 4)
    quote.paragraph_format.left_indent = Cm(1)
    quote.paragraph_format.right_indent = Cm(1)
    quote.paragraph_format.first_line_indent = Cm(0.74)
    quote.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # header / footer
    header = section.header
    header.paragraphs[0].text = "季度业务报告 · 内部资料"
    header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in header.paragraphs[0].runs:
        run.font.size = Pt(9)
    footer = section.footer
    footer_paragraph = footer.paragraphs[0]
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_paragraph.add_run("第 ")
    add_field(footer_paragraph, "PAGE")
    footer_paragraph.add_run(" 页，共 ")
    add_field(footer_paragraph, "NUMPAGES")
    footer_paragraph.add_run(" 页")
    for run in footer_paragraph.runs:
        run.font.size = Pt(9)

    doc.save(path)


def build_english(path: str) -> None:
    doc = Document()
    doc.add_heading("Product Requirements Document", level=0)
    doc.add_paragraph("Author: QA Team    Version: 1.0    Status: Draft")
    doc.add_heading("1. Background", level=1)
    body = doc.add_paragraph(
        "The legacy pipeline converts documents by shelling out to an office "
        "suite. This document specifies a lightweight alternative that renders "
        "WordprocessingML directly to PDF. " * 3)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_heading("2. Requirements", level=1)
    for text in ("Support headings, lists and tables.",
                 "Preserve inline formatting such as bold, italic and underline.",
                 "Handle CJK text without any external font configuration."):
        doc.add_paragraph(text, style="List Bullet")

    doc.add_heading("3. Data", level=1)
    table = doc.add_table(rows=3, cols=3)
    table.style = "Table Grid"
    for c, header in enumerate(["Metric", "Target", "Actual"]):
        cell = table.cell(0, c)
        cell.text = header
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
        shade_cell(cell, "EEEEEE")
    for r, row in enumerate([["Latency", "< 1s", "0.4s"],
                             ["Coverage", "100%", "98%"]], start=1):
        for c, value in enumerate(row):
            table.cell(r, c).text = value

    doc.add_paragraph()
    doc.add_picture(io.BytesIO(make_png(480, 240)), width=Cm(10))
    doc.save(path)


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    out_dir = argv[0] if argv else os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples")
    os.makedirs(out_dir, exist_ok=True)
    cn = os.path.join(out_dir, "sample_cn.docx")
    en = os.path.join(out_dir, "sample_en.docx")
    build_chinese(cn)
    build_english(en)
    print(f"生成样本 / wrote:\n  {cn}\n  {en}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
