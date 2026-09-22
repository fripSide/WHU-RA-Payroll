# -*- coding: utf-8 -*-
"""DOCX 明细表生成器。

直接从「助研费用发放说明」模板复制一份，在 OOXML 层面填充数据，
因此生成的文档与模板在字体、字号、对齐、边框、行高上完全一致。
"""
import os
import re
import copy

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.oxml.text.paragraph import CT_P

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DOCX = os.path.join(HERE, "templates", "武汉大学学生劳务费发放明细表（科研经费）.docx")

# 表格行号（按模板固定；改了模板结构要同步这里）
ROW_HEAD_1 = 0    # 发放单位 / 财务项目编号 / 兼职时段
ROW_HEAD_2 = 1    # 项目名称
ROW_NOTE = 2      # 发放事项说明
ROW_FIRST_DETAIL = 5   # 第一个明细行（模板自带两个学生，剩下的空行按需克隆）
ROW_TOTAL = 14    # 合计行

# 注意：这里刻意不定义任何字体常量。
# 生成时的字体一律从模板对应的单元格/段落里取（见 _split_by_script），
# 这样模板改成什么字体，导出的就是什么字体。


# --------------------------------------------------------------------------
# 数值格式化
# --------------------------------------------------------------------------
def fmt_number(value, blank_zero=False):
    """把输入格式化成适合填表的字符串；非数字原样返回。"""
    if value is None:
        return ""
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text == "":
            return ""
        try:
            number = float(text)
        except ValueError:
            return value.strip()
    else:
        try:
            number = float(value)
        except (TypeError, ValueError):
            return str(value)

    if blank_zero and number == 0:
        return ""
    if number == int(number) and abs(number) < 1e15:
        return str(int(number))
    return ("%.2f" % number).rstrip("0").rstrip(".")


def to_money(value):
    try:
        return float(str(value).replace(",", "").strip() or 0)
    except (TypeError, ValueError):
        return 0.0


def _id_sort_key(student):
    """按学号升序的排序键。

    学号按文本比较（常常超过 15 位，转数字会丢精度）；
    没填学号的排到最后。
    """
    sid = str(student.get("studentId") or "").strip()
    return (1, "") if sid == "" else (0, sid)


def effective_amount(student):
    """统一的金额口径：手填优先，否则 标准 × 工时。Word 和 Excel 共用。"""
    amount = student.get("amount")
    if student.get("manual") and amount not in (None, ""):
        return amount
    rate, hours = to_money(student.get("rate")), to_money(student.get("hours"))
    if rate == 0 and hours == 0:
        return amount if amount not in (None, "") else ""
    return fmt_number(rate * hours)


def batch_fill(students, field, mapping):
    """把一列数值按顺序套到学生上，返回每人配对后的结果。

    这是「批量填入一列」的唯一实现——预览界面和真正写入都走这里，
    所以界面上看到的金额一定等于导出文件里的金额。

    field:   "hours" / "rate" / "amount"，决定值写到哪个字段。
    mapping: {id: 值}，使用稳定人员 ID，允许学号暂时为空。
    """
    result = []
    for student in students:
        sid = student.get("studentId")
        key = student.get("id") or sid
        if key not in mapping:
            result.append({
                "id": student.get("id"),
                "studentId": sid,
                "name": student.get("name", ""),
                "value": None,
                "manual": bool(student.get("manual")),
                "amount": effective_amount(student),
            })
            continue

        value = mapping[key]
        probe = dict(student)
        probe[field] = value
        if field == "amount":
            probe["manual"] = True     # 直接填金额 = 手填，优先于标准×工时

        result.append({
            "id": student.get("id"),
            "studentId": sid,
            "name": student.get("name", ""),
            "value": value,
            "manual": bool(probe.get("manual")),
            "amount": effective_amount(probe),
        })
    return result


# --------------------------------------------------------------------------
# XML 小工具
#
# 原则：字体完全跟随模板，一个字都不自己定。
# 做法是先把模板里的 run（含它自己的 w:rPr）拷一份出来当作"字体样板"，
# 再只替换其中的文字。这样模板里那些显式设置（例如学号列的 sz=16、
# 表头字段的 Times New Roman + 宋体）都会原样保留。
# --------------------------------------------------------------------------
def _first_run_style(parent):
    """取父节点里第一个 run 的 rPr，作为字体样板；没有就返回 None。"""
    for run in parent.iter(qn("w:r")):
        rPr = run.find(qn("w:rPr"))
        if rPr is not None:
            return copy.deepcopy(rPr)
    return None


def _set_text_on_paragraph(para, text, style_rPr=None, fallback_para=None):
    """清空段落里的 run，写入新文字，字体用 style_rPr（模板的 rPr）。

    Word 的字体是分开设置的：w:ascii/w:hAnsi 管西文，w:eastAsia 管中文。
    模板里同一个单元格可能出现"数字用 Times New Roman、中文用宋体"的情况
    （例如项目名称 2026年才科研启动经费）。这里按字符类型拆分 run，
    把这两种字体都还原出来，保证和模板一模一样。
    """
    # 先找字体样板：模板该段落自己的 -> 调用方给的 -> 备用段落的
    if style_rPr is None:
        style_rPr = _first_run_style(para)
    if style_rPr is None and fallback_para is not None:
        style_rPr = _first_run_style(fallback_para)

    runs_to_copy = _run_styles(para)
    if not runs_to_copy and fallback_para is not None:
        runs_to_copy = _run_styles(fallback_para)

    for run in para.findall(qn("w:r")):
        para.remove(run)
    for tag in ("w:hyperlink", "w:bookmarkStart", "w:bookmarkEnd"):
        for node in para.findall(qn(tag)):
            para.remove(node)
    for node in list(para):
        if node.tag == qn("w:t"):
            para.remove(node)

    body = "" if text is None else str(text)
    chunks = _split_by_script(body, runs_to_copy, style_rPr)

    for chunk, rPr in chunks:
        run = OxmlElement("w:r")
        if rPr is not None:
            run.append(copy.deepcopy(rPr))
        para.append(run)
        text_node = OxmlElement("w:t")
        text_node.set(qn("xml:space"), "preserve")
        text_node.text = chunk
        run.append(text_node)

    if not chunks:  # 空字符串也要留一个 run，保持段落结构
        run = OxmlElement("w:r")
        if style_rPr is not None:
            run.append(copy.deepcopy(style_rPr))
        para.append(run)
        text_node = OxmlElement("w:t")
        text_node.set(qn("xml:space"), "preserve")
        text_node.text = ""
        run.append(text_node)

    return para


def _run_styles(parent):
    """收集段落里所有 run 的 rPr（去重、保序），用于还原混排字体。"""
    out = []
    for run in parent.findall(qn("w:r")):
        rPr = run.find(qn("w:rPr"))
        xml = rPr.xml if rPr is not None else ""
        if not any(other.xml == xml for other in out):
            out.append(rPr if rPr is not None else OxmlElement("w:rPr"))
    return out


ASCII_RE = re.compile(r"[\x00-\x7F]+")
CJK_RE = re.compile(r"[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef]")


def _split_by_script(text, run_styles, style_rPr):
    """把文字按"西文/非西文"切开，分别套用模板里对应的字体。

    例：模板 r1c1 里 "2026" 单独一个 run 且不带 rFonts（继承 Times New Roman），
    "年才科研启动经费" 则是 ascii=宋体 + hint=eastAsia。合并成一段会丢掉这个区别，
    所以这里按字符类型还原成同样的结构。
    """
    if not text:
        return []

    # 找出模板里的"中文字体 run"（带 eastAsia）和"西文字体 run"（没带）
    east_rPr = None
    latin_rPr = None
    for rPr in run_styles:
        fonts = rPr.find(qn("w:rFonts"))
        if fonts is None:
            if latin_rPr is None:
                latin_rPr = rPr
            continue
        if fonts.get(qn("w:eastAsia")):
            if east_rPr is None:
                east_rPr = rPr
        elif latin_rPr is None:
            latin_rPr = rPr

    if east_rPr is None:
        # 模板这一格只有一种字体，整体照用
        return [(text, latin_rPr if latin_rPr is not None else style_rPr)]

    east_ascii = _font_attr(east_rPr, "ascii")
    east_east = _font_attr(east_rPr, "eastAsia")
    latin_ascii = _font_attr(latin_rPr, "ascii") if latin_rPr is not None else None

    # 中文 run 的西文字体和中文字体一致，说明没有区分，直接一段
    if latin_rPr is None or latin_ascii == east_ascii or not CJK_RE.search(text):
        return [(text, east_rPr)]

    # 西文部分照抄模板里那段「没有 rFonts」的 run：字体交给段落/表格默认，
    # 但要保留它自己的字号、加粗等设置。
    latin_chunk_rPr = copy.deepcopy(latin_rPr)
    fonts = latin_chunk_rPr.find(qn("w:rFonts"))
    if fonts is not None:
        latin_chunk_rPr.remove(fonts)
    # rPr 空了就整个去掉，和模板保持一致
    if len(latin_chunk_rPr) == 0:
        latin_chunk_rPr = None

    chunks = []
    pos = 0
    for match in ASCII_RE.finditer(text):
        if match.start() > pos:
            segment = text[pos:match.start()]
            if segment:
                chunks.append((segment, east_rPr))
        if match.group():
            chunks.append((match.group(), latin_chunk_rPr))
        pos = match.end()
    if pos < len(text):
        chunks.append((text[pos:], east_rPr))
    return chunks


def _font_attr(rPr, attr):
    if rPr is None:
        return None
    fonts = rPr.find(qn("w:rFonts"))
    if fonts is None:
        return None
    return fonts.get(qn("w:" + attr))


def set_cell_text(tc, text, fallback_cell=None):
    """替换单元格文字，字体沿用模板该单元格自身的 run 格式。

    fallback_cell: 该单元格在模板里是空的（没有 run）时，
    从这个单元格借字体样板，保证字体和模板一致。
    """
    paragraphs = tc.findall(qn("w:p"))
    keep = paragraphs[0] if paragraphs else None
    for para in paragraphs[1:]:
        tc.remove(para)
    if keep is None:
        keep = CT_P()
        tc.append(keep)

    fallback_para = None
    if fallback_cell is not None:
        fallback_paras = fallback_cell.findall(qn("w:p"))
        fallback_para = fallback_paras[0] if fallback_paras else None

    return _set_text_on_paragraph(keep, text, fallback_para=fallback_para)


def replace_paragraph_text(para, text, fallback_para=None):
    """替换段落文字，段落级格式（缩进、行距）和字体都保持原样。"""
    return _set_text_on_paragraph(para, text, fallback_para=fallback_para)


def cell_text(tc):
    return "".join(node.text or "" for node in tc.iter(qn("w:t"))).strip()


# OOXML 里这些容器内的子元素顺序是固定的，新元素必须插到正确位置，
# 否则 Word 打开时会提示文档已损坏。
TCPR_ORDER = ("w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge",
              "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
              "w:tcFitText", "w:vAlign", "w:hideMark")


def insert_ordered(parent, element, order):
    """按 schema 顺序把 element 插入 parent，避免子元素顺序非法。"""
    tag = element.tag
    index = None
    for position, name in enumerate(order):
        if qn(name) == tag:
            index = position
            break
    if index is None:
        parent.append(element)
        return element

    for child in list(parent):
        child_index = None
        for position, name in enumerate(order):
            if qn(name) == child.tag:
                child_index = position
                break
        if child_index is not None and child_index > index:
            child.addprevious(element)
            return element
    parent.append(element)
    return element


def set_paragraph_align(para, value="center"):
    """设置段落对齐。

    模板里的单元格段落本来就是居中的，存在 w:jc 时只改值、不动顺序——
    w:pPr 的子元素顺序在 OOXML 里是固定的（w:jc 必须在 w:rPr 之前），
    删除后重新 append 会把 w:jc 排到 w:rPr 后面，属于非法结构。
    """
    pPr = para.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        para.insert(0, pPr)

    existing = pPr.find(qn("w:jc"))
    if existing is not None:
        existing.set(qn("w:val"), value)
        return existing

    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), value)

    # 按 schema 顺序插入：w:jc 位于 w:spacing/w:ind 之后、w:rPr 之前
    anchor = None
    for tag in ("w:ind", "w:spacing", "w:snapToGrid", "w:adjustRightInd"):
        found = pPr.find(qn(tag))
        if found is not None:
            anchor = found
            break
    if anchor is not None:
        anchor.addnext(jc)
    else:
        rPr = pPr.find(qn("w:rPr"))
        if rPr is not None:
            rPr.addprevious(jc)
        else:
            pPr.append(jc)
    return jc


def set_cell_borders_single(tc, sz="4"):
    """给单元格加全边框（用于“合计”行突出显示）。"""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement("w:" + edge)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), sz)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "000000")
        borders.append(node)
    insert_ordered(tcPr, borders, TCPR_ORDER)
    return borders


def set_cell_shading(tc, fill="F2F2F2"):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    insert_ordered(tcPr, shd, TCPR_ORDER)
    return shd


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------
def build_note(settings, students):
    """生成发放事项说明正文。

    - 说明里带 {人数}/{第一位学生} 等占位符：替换占位符
    - 说明是模板自带的样例（如“本人项目组柯昀志等9位…”）：按当前名单重写
    - 其它（用户自己写的）：原样保留
    """
    template = (settings.get("note") or "").strip()
    count = len(students)
    names = [s.get("name", "").strip() for s in students if s.get("name", "").strip()]
    lead_name = names[0] if names else ""

    if not template:
        return _auto_note(settings, count, lead_name)

    if any(token in template for token in ("{人数}", "{count}", "{第一位学生}",
                                           "{姓名}", "{时段}", "{项目名称}")):
        out = template
        out = out.replace("{人数}", str(count)).replace("{count}", str(count))
        out = out.replace("{第一位学生}", lead_name).replace("{姓名}", lead_name)
        out = out.replace("{时段}", str(settings.get("period", "")))
        out = out.replace("{项目名称}", str(settings.get("projectName", "")))
        return out

    # 模板自带的样例说明：人数/姓名会过期，按当前名单重新生成
    if re.search(r"本人项目组.*?参与科研项目", template):
        return _auto_note(settings, count, lead_name)

    return template


def _auto_note(settings, count, lead_name):
    period = str(settings.get("period", "")).strip()
    project = str(settings.get("projectName", "")).strip()
    who = ("本人项目组%s等%d位本科生、研究生" % (lead_name, count)) if lead_name \
        else ("本人项目组%d位本科生、研究生" % count)
    parts = [who, "在%s" % period if period else "",
             "参与科研项目%s" % project if project else "参与科研项目",
             "的研究。负责相关研发工作。明细见下表："]
    return "".join(parts)


def inspect_template(document):
    """Locate detail headers/total by content, supporting both bundled layouts."""
    if not document.tables:
        raise ValueError("模板里没有表格")
    rows = document.tables[0]._tbl.findall(qn("w:tr"))
    header = None
    for index, row in enumerate(rows):
        cells = row.findall(qn("w:tc"))
        labels = [re.sub(r"\s+", "", cell_text(cell)) for cell in cells]
        if "姓名" in labels and "学号" in labels and any("实发金额" in label for label in labels):
            header = index
            columns = labels
            break
    if header is None or len(columns) not in (5, 7):
        raise ValueError("模板须包含姓名、学号、实发金额，以及科研表的事由列或非科研表的学院/标准/工时列")
    total = next((i for i in range(header + 1, len(rows))
                  if re.sub(r"\s+", "", cell_text(rows[i].findall(qn('w:tc'))[0])) == "合计"), None)
    if total is None or total <= header + 1:
        raise ValueError("模板需要至少一行人员明细和一行合计")
    if len(rows[0].findall(qn("w:tc"))) < 6 or len(rows[1].findall(qn("w:tc"))) < 2:
        raise ValueError("模板表头须包含发放单位、财务项目编号、时段和项目名称")
    return {"rows": rows, "header": header, "first": header + 1, "total": total,
            "research": len(columns) == 5}


def extract_template(document):
    layout = inspect_template(document)
    rows = layout["rows"]
    head = [cell_text(c) for c in rows[0].findall(qn("w:tc"))]
    settings = dict(unitName=head[1], projectCode=head[3], period=head[5],
                    projectName=cell_text(rows[1].findall(qn("w:tc"))[1]),
                    projectType="research" if layout["research"] else "non_research")
    students = []
    for row in rows[layout["first"]:layout["total"]]:
        cells = [cell_text(c) for c in row.findall(qn("w:tc"))]
        if not cells[1] and not cells[2]:
            continue
        data = dict(name=cells[1], studentId=cells[2], checked=True, manual=True)
        if layout["research"]:
            data.update(amount=cells[3], reason=cells[4])
        else:
            data.update(college=cells[3], rate=cells[4], hours=cells[5], amount=cells[6])
        students.append(data)
    return settings, students


def generate_docx(payload, out_path, template_path=None):
    """按 payload 生成明细表 docx。

    payload = {
      "settings": {...}, "students": [ {studentId,name,college,rate,hours,amount}, ... ]
    }
    """
    source = template_path or payload.get("docxTemplate") or DEFAULT_DOCX
    if not os.path.isfile(source):
        raise FileNotFoundError("找不到明细表模板：%s" % source)

    doc = Document(source)
    layout = inspect_template(doc)
    table = doc.tables[0]._tbl
    rows = layout["rows"]
    first, total_index = layout["first"], layout["total"]

    settings = payload.get("settings") or {}
    students = [s for s in (payload.get("students") or []) if s.get("checked", True)]
    # 明细表按学号从小到大排；没填学号的排在最后（排序稳定，保持原相对顺序）
    students = sorted(students, key=_id_sort_key)

    # ---- 1. 头部字段（字体沿用模板单元格自己的格式）--------------------
    head1 = rows[ROW_HEAD_1].findall(qn("w:tc"))
    set_cell_text(head1[1], settings.get("unitName", ""))
    set_cell_text(head1[3], settings.get("projectCode", ""))
    set_cell_text(head1[5], settings.get("period", ""))

    head2 = rows[ROW_HEAD_2].findall(qn("w:tc"))
    set_cell_text(head2[1], settings.get("projectName", ""))

    # ---- 2. 发放事项说明 ------------------------------------------------
    # 模板该段落自带字体（Times New Roman + 等线 + sz20），直接沿用
    for row in rows[2:layout["header"]]:
        note_cell = row.findall(qn("w:tc"))[0]
        if "发放事项说明" in re.sub(r"\s+", "", cell_text(note_cell)):
            note_paras = note_cell.findall(qn("w:p"))
            if len(note_paras) >= 2:
                replace_paragraph_text(note_paras[1], build_note(settings, students))

    # ---- 3. 明细行 ------------------------------------------------------
    template_row = rows[first]
    total_row = rows[total_index]

    # 模板里只有前两行填了学生，其余明细行是空的（没有 run）。
    # 记下这两行的字体样板，空行就照它们来，保证字体和模板一致。
    style_rows = [copy.deepcopy(rows[first])]

    def style_cell(source_row, col):
        """从样板行里取第 col 个单元格，作为字体来源。"""
        cells = source_row.findall(qn("w:tc"))
        return cells[col] if col < len(cells) else None

    # 先补齐到足够行数（模板只有 10 个明细行位置）
    available = total_index - first
    if len(students) > available:
        anchor = template_row
        for _ in range(len(students) - available):
            clone = copy.deepcopy(template_row)
            anchor.addnext(clone)
            anchor = clone
        rows = table.findall(qn("w:tr"))
        total_row = rows[total_index + (len(students) - available)]

    # 刷新明细行内容
    detail_rows = rows[first:rows.index(total_row)]
    total_amount = 0.0
    for index, row in enumerate(detail_rows):
        cells = row.findall(qn("w:tc"))
        # 这一行自己的 run 就是最好的字体样板；空行则借用前两行
        donor = row if _first_run_style(row) is not None else (
            style_rows[index] if index < len(style_rows) else style_rows[0])

        if index < len(students):
            student = students[index]
            amount = effective_amount(student)
            total_amount += to_money(amount)
            values = [
                str(index + 1),
                student.get("name", ""),
                student.get("studentId", ""),
                student.get("college", ""),
                fmt_number(student.get("rate"), blank_zero=True),
                fmt_number(student.get("hours"), blank_zero=True),
                fmt_number(amount, blank_zero=True),
            ]
            if layout["research"]:
                values = [str(index + 1), student.get("name", ""), student.get("studentId", ""),
                          fmt_number(amount, blank_zero=True), student.get("reason") or build_note(settings, students)]
            height = row.find("w:trPr/w:trHeight", namespaces=row.nsmap)
            if height is not None:
                height.set(qn("w:hRule"), "atLeast")
            for col, value in enumerate(values):
                if col < len(cells):
                    set_cell_text(cells[col], value, fallback_cell=style_cell(donor, col))
                    for para in cells[col].findall(qn("w:p")):
                        set_paragraph_align(para, "center")
        else:
            for col, cell in enumerate(cells):
                set_cell_text(cell, "", fallback_cell=style_cell(donor, col))
                for para in cell.findall(qn("w:p")):
                    set_paragraph_align(para, "center")

    # ---- 4. 合计行 ------------------------------------------------------
    # 模板的合计行第一个单元格是合并的“合 计”区（含字体），
    # 最后一个单元格是空的，字体要从第一个借。
    total_cells = total_row.findall(qn("w:tc"))
    set_cell_text(total_cells[0], "合           计")
    for para in total_cells[0].findall(qn("w:p")):
        set_paragraph_align(para, "center")
    set_cell_text(total_cells[-1], fmt_number(total_amount),
                  fallback_cell=total_cells[0])
    for para in total_cells[-1].findall(qn("w:p")):
        set_paragraph_align(para, "center")
    for cell in total_cells:
        set_cell_borders_single(cell, "4")
        set_cell_shading(cell, "F2F2F2")

    # Repeat the form's header on additional pages; keep the supplied approval text.
    for row in rows[:layout["header"] + 1]:
        props = row.get_or_add_trPr()
        if props.find(qn("w:tblHeader")) is None:
            props.append(OxmlElement("w:tblHeader"))

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)

    return {
        "path": out_path,
        "students": len(students),
        "total": total_amount,
    }
