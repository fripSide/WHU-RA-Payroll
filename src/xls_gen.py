# -*- coding: utf-8 -*-
"""名单 .xls 读写。

导出用 xlwt 生成真正的 BIFF8/OLE 格式（后缀 .xls），与模板一致；
导入同时兼容 .xls / .xlsx / .csv。
"""
import os
import re
import csv
import io
import sys

# 本文件在 src/ 下，vendor/、templates/ 都在项目根（上一级）
SRC_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SRC_DIR)
VENDOR = os.path.join(ROOT, "vendor", "pylib")
if VENDOR not in sys.path:
    sys.path.insert(0, VENDOR)

ID_KEYS = ("学号", "学生学号", "工号", "编号", "id")
NAME_KEYS = ("姓名", "学生姓名", "名字", "name")
AMOUNT_KEYS = ("助研津贴", "津贴", "补贴", "金额", "实发", "费")
COLLEGE_KEYS = ("学院", "所在学院", "院系", "单位")
RATE_KEYS = ("标准", "单价")
HOURS_KEYS = ("工时", "时长", "小时")
CHECKED_KEYS = ("勾选", "选中", "是否发放", "本次发放")

# 读「勾选」列时，这些写法算"要发"
CHECKED_TRUE = ("是", "y", "yes", "true", "1", "√", "✓", "✔", "v", "对", "勾选", "发放")
# 这些写法算"不发"
CHECKED_FALSE = ("否", "n", "no", "false", "0", "×", "x", "✗", "不", "不发", "-", "—")


def sort_key_by_id(student):
    """按学号升序的排序键。

    学号一律按文本比较（学号常常超过 15 位，转成数字会丢精度）；
    没填学号的排到最后，保持它们原有的相对顺序（Python 排序是稳定的）。
    """
    sid = str(student.get("studentId") or "").strip()
    return (1, "") if sid == "" else (0, sid)


def fmt_number(value, blank_zero=False):
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


def _match_column(header, keys):
    for index, cell in enumerate(header):
        text = str(cell).strip()
        for key in keys:
            if key in text:
                return index
    return None


def _normalize_header(header):
    return {
        "id": _match_column(header, ID_KEYS),
        "name": _match_column(header, NAME_KEYS),
        "amount": _match_column(header, AMOUNT_KEYS),
        "college": _match_column(header, COLLEGE_KEYS),
        "rate": _match_column(header, RATE_KEYS),
        "hours": _match_column(header, HOURS_KEYS),
        "checked": _match_column(header, CHECKED_KEYS),
    }


def _looks_like_data(row):
    """判断这一行是数据还是表头。

    表头里是人写的列名（学号、姓名…），数据行里是学号/姓名/数字。
    区分不清的按表头处理（保守），但只要出现明显的"数据特征"就按数据算：
      - 有一格是 7 位以上的连续数字（典型学号）
      - 有 2 格以上是纯数字
    """
    cells = [str(c).strip() for c in row if str(c).strip() != ""]
    if not cells:
        return False
    for cell in cells:
        if re.search(r"\d{7,}", re.sub(r"\s", "", cell)):
            return True
    numeric = sum(1 for c in cells if re.fullmatch(r"-?\d+(\.\d+)?", c))
    return numeric >= 2


COLLEGE_HINT = re.compile(r"(学院|大学|学部|院系|系$|研究所|中心|实验室)")


def _guess_columns(width, sample_rows):
    """没有表头时猜列位置。

    列数能说明一部分问题，但第 3 列到底是"学院"还是"金额"光看列数分不出来，
    所以再看一眼实际值：像"计算机学院"这种就按学院，纯数字就按金额。
    """
    def col_values(index):
        out = []
        for row in sample_rows:
            if index < len(row):
                v = str(row[index]).strip()
                if v:
                    out.append(v)
        return out

    base = {"id": 0, "name": 1, "amount": None,
            "college": None, "rate": None, "hours": None, "checked": None}
    if width <= 2:
        return base

    third = col_values(2)
    if third:
        texty = sum(1 for v in third if COLLEGE_HINT.search(v))
        numeric = sum(1 for v in third if re.fullmatch(r"-?\d+(\.\d+)?", v))
        if texty and texty >= numeric:
            base["college"] = 2
            base["rate"] = 3 if width > 3 else None
            base["hours"] = 4 if width > 4 else None
            return base
        if numeric and not texty:
            base["amount"] = 2
            base["rate"] = 3 if width > 3 else None
            base["hours"] = 4 if width > 4 else None
            return base

    # 分不清：按最常见的 学号|姓名|学院|标准|工时 处理
    base["college"] = 2
    base["rate"] = 3 if width > 3 else None
    base["hours"] = 4 if width > 4 else None
    return base


def _rows_to_students(rows):
    """rows: 行列表（可能含表头）。返回 (students, header_used)

    如果第一行就是数据（比如用户从 Excel 里选了一块**不含表头**的区域复制），
    要按数据行处理，不能当成表头吃掉——否则整列错位，学院会全空。
    """
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        return [], None

    header_index = 0
    columns = None
    for index in range(min(len(rows), 6)):
        candidate = _normalize_header(rows[index])
        if candidate["id"] is not None or candidate["name"] is not None:
            header_index = index
            columns = candidate
            break

    if columns is None:
        # 认不出表头：看看第一行到底是不是数据
        width = max(len(r) for r in rows)
        if _looks_like_data(rows[0]):
            # 第一行就是数据 -> 无表头模式，整行都要读进来
            header_index = -1
        else:
            # 分不清，按老规矩跳过第一行
            header_index = 0
        columns = _guess_columns(width, rows)
    has_checked_col = columns.get("checked") is not None

    students = []
    for row in rows[header_index + 1:]:
        def value(key):
            col = columns.get(key)
            if col is None or col >= len(row):
                return ""
            cell = row[col]
            # 空单元格：xlrd 给 ''，openpyxl 给 None。都要当空字符串处理，
            # 否则 str('') -> '0.0' 这种会把"空"变成有值，后面的判空就失效了。
            if cell is None or (isinstance(cell, str) and cell.strip() == ""):
                return ""
            if isinstance(cell, float) and cell == int(cell):
                return str(int(cell))
            return str(cell).strip()

        student_id = value("id")
        name = value("name")
        if not student_id and not name:
            continue
        if "合计" in (student_id + name) or "总计" in (student_id + name):
            continue

        # 有「勾选」列就按它判，没有就默认要发
        checked = True
        if has_checked_col:
            mark = value("checked").lower()
            if mark in CHECKED_FALSE or mark == "":
                checked = False
            elif mark in CHECKED_TRUE:
                checked = True
            else:
                checked = True      # 认不出来的写法，宁可选上

        students.append({
            "studentId": student_id,
            "name": name,
            "college": value("college"),
            "rate": fmt_number(value("rate")),
            "hours": fmt_number(value("hours")),
            "amount": fmt_number(value("amount")),
            "checked": checked,
        })
    return students, (rows[header_index] if header_index >= 0 else None)


def read_roster(path):
    """读取名单文件，返回学生列表。"""
    ext = os.path.splitext(path)[1].lower()

    if ext == ".xls":
        import xlrd
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        rows = [[sheet.cell_value(r, c) for c in range(sheet.ncols)]
                for r in range(sheet.nrows)]
        return _rows_to_students(rows)[0]

    if ext == ".xlsx":
        from openpyxl import load_workbook
        book = load_workbook(path, data_only=True)
        sheet = book[book.sheetnames[0]]
        rows = [list(r) for r in sheet.iter_rows(values_only=True)]
        return _rows_to_students(rows)[0]

    if ext in (".csv", ".txt"):
        raw = open(path, "rb").read()
        for encoding in ("utf-8-sig", "gbk", "utf-8", "latin-1"):
            try:
                text = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("无法识别文本编码")
        delimiter = "\t" if "\t" in text.splitlines()[0] else ","
        rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
        return _rows_to_students(rows)[0]

    raise ValueError("不支持的文件类型：%s（请用 .xls/.xlsx/.csv）" % ext)


def read_roster_from_text(text):
    """解析粘贴进来的文本（支持制表符/逗号/多空格分隔）。"""
    lines = [ln for ln in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
             if ln.strip()]
    if not lines:
        return []
    delimiter = "\t"
    if "\t" not in lines[0]:
        if "," in lines[0]:
            delimiter = ","
        elif "，" in lines[0]:
            delimiter = "，"
        else:
            delimiter = None
    rows = []
    for line in lines:
        if delimiter is None:
            rows.append([c for c in re.split(r"\s{2,}|\s+", line.strip()) if c != ""])
        else:
            rows.append([c.strip() for c in line.split(delimiter)])
    return _rows_to_students(rows)[0]


def write_roster_xls(students, out_path, sheet_name="Sheet1"):
    """导出名单 .xls（BIFF8）。

    列：勾选 | 学号 | 姓名 | 所在学院 | 标准（元/时） | 工时 | 助研津贴(三兼费)

    - **按学号从小到大排序**，没填学号的排在最后
    - 勾选列写「是」的才是本次要发的；没勾的也留着，作为下次的备选
    - 这份名单可以直接再导入：勾选状态、学院、标准、工时都会带回来
    """
    import xlwt

    # 未勾选的也一并保留（当备选），按学号升序
    ordered = sorted(list(students or []), key=sort_key_by_id)

    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet(sheet_name or "Sheet1")

    header_style = xlwt.XFStyle()
    header_font = xlwt.Font()
    header_font.name = "宋体"
    header_font.bold = True
    header_style.font = header_font
    header_style.alignment = xlwt.Alignment()
    header_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    header_style.alignment.vert = xlwt.Alignment.VERT_CENTER

    body_style = xlwt.XFStyle()
    body_font = xlwt.Font()
    body_font.name = "宋体"
    body_style.font = body_font
    body_style.alignment = xlwt.Alignment()
    body_style.alignment.vert = xlwt.Alignment.VERT_CENTER

    center_style = xlwt.XFStyle()
    center_style.font = body_font
    center_style.alignment = xlwt.Alignment()
    center_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    center_style.alignment.horz = xlwt.Alignment.HORZ_CENTER

    # 金额列带千分位，和页面上看到的一致；单元格里仍是数字，导回来不影响计算
    money_style = xlwt.XFStyle()
    money_style.font = body_font
    money_style.alignment = xlwt.Alignment()
    money_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    money_style.num_format_str = "#,##0.##"

    headers = ["勾选", "学号", "姓名", "所在学院", "标准（元/时）", "工时",
               "助研津贴(三兼费)"]
    for col, title in enumerate(headers):
        sheet.write(0, col, title, header_style)

    for index, student in enumerate(ordered):
        row = index + 1
        sheet.write(row, 0, "是" if student.get("checked", True) else "",
                    center_style)
        # 学号一律按文本写入，避免长数字被转成科学计数法
        sheet.write(row, 1, str(student.get("studentId", "")), center_style)
        sheet.write(row, 2, str(student.get("name", "")), center_style)
        sheet.write(row, 3, str(student.get("college", "") or ""), center_style)

        for col, key in ((4, "rate"), (5, "hours")):
            value = student.get(key)
            if value in (None, ""):
                sheet.write(row, col, "", center_style)
            else:
                sheet.write(row, col, _to_number(value), center_style)

        amount = student.get("amount")
        if amount in (None, ""):
            amount = _to_number(student.get("rate")) * _to_number(student.get("hours"))
        sheet.write(row, 6, _to_number(amount), money_style)

    for col, width in enumerate((8, 18, 12, 24, 13, 10, 22)):
        sheet.col(col).width = 256 * width

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    book.save(out_path)
    return {"path": out_path, "students": len(ordered),
            "checked": sum(1 for s in ordered if s.get("checked", True))}


def write_submission_xls(students, out_path):
    """Three-column upload form; only checked people, same headers as the supplied XLS."""
    import xlrd
    import xlwt
    template = os.path.join(ROOT, "templates", "助研费用发放列表.xls")
    source = xlrd.open_workbook(template)
    headers = source.sheet_by_index(0).row_values(0)
    if headers != ["学号", "姓名", "助研津贴(三兼费)"]:
        raise ValueError("系统上传 Excel 模板须为学号、姓名、助研津贴三列")
    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet(source.sheet_names()[0])
    for col, value in enumerate(headers):
        sheet.write(0, col, value)
        sheet.col(col).width = 256 * (22 if col != 1 else 16)
    ordered = sorted([s for s in students if s.get("checked", True)], key=sort_key_by_id)
    for index, person in enumerate(ordered, 1):
        sheet.write(index, 0, str(person.get("studentId") or ""))
        sheet.write(index, 1, str(person.get("name") or ""))
        sheet.write(index, 2, _to_number(person.get("amount")))
    book.save(out_path)
    return {"path": out_path, "students": len(ordered)}


def _to_number(value):
    try:
        return float(str(value).replace(",", "").strip() or 0)
    except (TypeError, ValueError):
        return 0.0
