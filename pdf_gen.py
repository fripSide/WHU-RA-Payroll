"""Render the supported Word payroll tables to PDF without an Office installation.

Reads the generated DOCX, including its text, table grid, merged cells, type size,
page size and margins. This is a form renderer, not a general Word converter.
"""
import os
import re
import threading
from xml.sax.saxutils import escape

from docx import Document
from docx.oxml.ns import qn
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, LongTable, TableStyle

import docx_gen

_FONT_LOCK = threading.Lock()


def _font():
    with _FONT_LOCK:
        if "PayrollSong" in pdfmetrics.getRegisteredFontNames():
            return "PayrollSong"
        candidates = [os.environ.get("BAOXIAO_PDF_FONT", ""),
                      os.path.join(os.environ.get("WINDIR", "C:/Windows"), "Fonts", "simsun.ttc"),
                      "/System/Library/Fonts/Supplemental/Songti.ttc",
                      "/usr/share/fonts/truetype/arphic/uming.ttc"]
        for path in candidates:
            if path and os.path.isfile(path):
                try:
                    pdfmetrics.registerFont(TTFont("PayrollSong", path, subfontIndex=0))
                    return "PayrollSong"
                except Exception:
                    continue
        if "STSong-Light" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        return "STSong-Light"


def _paragraph(element, font, default_size=10.5, default_align=0):
    text = "".join((node.text or "") if node.tag == qn("w:t") else "\n"
                   for node in element.iter() if node.tag in (qn("w:t"), qn("w:br")))
    size = next((float(n.get(qn("w:val"))) / 2 for n in element.iter(qn("w:sz"))), default_size)
    # Text before the table can contain the document title and signature line.
    align = element.find("w:pPr/w:jc", namespaces=element.nsmap)
    align = {"center": 1, "right": 2}.get(align.get(qn("w:val")), 0) if align is not None else default_align
    style = ParagraphStyle("form", fontName=font, fontSize=size, leading=size * 1.35,
                           alignment=align, wordWrap="CJK", spaceAfter=0, spaceBefore=0)
    markup = escape(text).replace("\n", "<br/>")
    markup = re.sub(r" {2,}", lambda m: "&#160;" * len(m.group()), markup)
    return Paragraph(markup or "&#160;", style)


def generate_pdf(docx_path, out_path):
    doc = Document(docx_path)
    layout = docx_gen.inspect_template(doc)
    if doc.inline_shapes:
        raise ValueError("PDF 导出支持文字和表格底板；含图片的自定义底板请先移除图片或用 Word 另存 PDF")
    section = doc.sections[0]
    font = _font()
    width, height = section.page_width.pt, section.page_height.pt
    left, right = section.left_margin.pt, section.right_margin.pt
    usable = width - left - right - 12  # Platypus frame includes 6pt padding per side.
    story = []
    for element in doc.element.body:
        if element.tag == qn("w:p"):
            raw_text = "".join(n.text or "" for n in element.iter(qn("w:t"))).strip()
            if "单位盖章" in raw_text and "负责人签字" in raw_text:
                signature = LongTable([[Paragraph("单位盖章：", ParagraphStyle("left",fontName=font,fontSize=10.5)),
                                        Paragraph("负责人签字：", ParagraphStyle("right",fontName=font,fontSize=10.5))]],
                                      colWidths=[usable * .65, usable * .35])
                story.extend([signature, Spacer(1, 5)])
                continue
            if raw_text.count("（线下报销）") == 2:
                labels = LongTable([[Paragraph("（线下报销）", ParagraphStyle("left",fontName=font,fontSize=10.5)),
                                     Paragraph("（线下报销）", ParagraphStyle("right",fontName=font,fontSize=10.5))]],
                                   colWidths=[usable * .65, usable * .35])
                story.extend([labels, Spacer(1, 5)])
                continue
            if "".join(element.itertext()).strip():
                story.append(_paragraph(element, font))
                story.append(Spacer(1, 5))
            else:
                story.append(Spacer(1, 5))
        elif element.tag == qn("w:tbl"):
            grid = element.find(qn("w:tblGrid"))
            widths = [int(c.get(qn("w:w"))) / 20 for c in grid] if grid is not None else []
            if not widths:
                raise ValueError("Word 模板缺少表格列宽")
            scale = min(1, usable / sum(widths))
            widths = [w * scale for w in widths]
            rows, spans, backgrounds, minimums = [], [], [], []
            for ri, row in enumerate(element.findall(qn("w:tr"))):
                values, ci = [""] * len(widths), 0
                for cell in row.findall(qn("w:tc")):
                    span = cell.find("w:tcPr/w:gridSpan", namespaces=cell.nsmap)
                    count = int(span.get(qn("w:val"))) if span is not None else 1
                    values[ci] = [_paragraph(p, font, default_align=1)
                                  for p in cell.findall(qn("w:p"))]
                    if count > 1:
                        spans.append(("SPAN", (ci, ri), (ci + count - 1, ri)))
                    shade = cell.find("w:tcPr/w:shd", namespaces=cell.nsmap)
                    if shade is not None and re.fullmatch(r"[0-9a-fA-F]{6}", shade.get(qn("w:fill"), "")):
                        backgrounds.append(("BACKGROUND", (ci, ri), (ci + count - 1, ri),
                                            colors.HexColor("#" + shade.get(qn("w:fill")))))
                    ci += count
                rows.append(values)
                row_height = row.find("w:trPr/w:trHeight", namespaces=row.nsmap)
                minimums.append(float(row_height.get(qn("w:val"))) / 20 if row_height is not None else 20)
            table = LongTable(rows, colWidths=widths, minRowHeights=minimums,
                              repeatRows=layout["header"] + 1, hAlign="CENTER")
            table.setStyle(TableStyle([
                ("GRID", (0, 0), (-1, -1), .5, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ] + spans + backgrounds))
            story.append(table)
            story.append(Spacer(1, 10))
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    result = SimpleDocTemplate(out_path, pagesize=(width, height), leftMargin=left, rightMargin=right,
                               topMargin=section.top_margin.pt, bottomMargin=section.bottom_margin.pt,
                               title="武汉大学学生劳务费发放明细表", author="")
    result.build(story)
    return {"path": out_path}
