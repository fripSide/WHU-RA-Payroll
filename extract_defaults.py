# -*- coding: utf-8 -*-
"""从现有的两个模板文件（.xls 名单 + .docx 明细表）提取默认配置。

生成 templates/defaults.json，供前端首次打开时作为初始值。
用法：python extract_defaults.py
"""
import os
import sys
import json
import re

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "vendor", "pylib"))

sys.stdout.reconfigure(encoding="utf-8")

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

try:
    import xlrd
except ImportError:  # pragma: no cover
    xlrd = None


def find_first(patterns):
    """在目录中按正则找出第一个匹配的文件。"""
    for name in sorted(os.listdir(HERE)):
        for pat in patterns:
            if re.search(pat, name):
                return os.path.join(HERE, name)
    return None


def cell_text(cell):
    return "".join(node.text or "" for node in cell.iter(qn("w:t"))).strip()


def parse_roster_xls(path):
    """读取名单 .xls：支持 学号/姓名/助研津贴 三列，自动识别表头。"""
    if not path or xlrd is None:
        return []
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)

    id_col, name_col, amount_col = 0, 1, 2
    header_row = -1
    for r in range(min(sheet.nrows, 5)):
        row_vals = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
        joined = "".join(row_vals)
        if "学号" in joined or "姓名" in joined:
            header_row = r
            for c, v in enumerate(row_vals):
                if "学号" in v:
                    id_col = c
                elif "姓名" in v:
                    name_col = c
                elif "金额" in v or "津贴" in v or "补贴" in v or "费" in v:
                    amount_col = c
            break

    students = []
    for r in range(header_row + 1, sheet.nrows):
        def val(col):
            if col is None or col >= sheet.ncols:
                return ""
            v = sheet.cell_value(r, col)
            if isinstance(v, float) and v == int(v):
                return str(int(v))
            return str(v).strip()

        sid = val(id_col)
        name = val(name_col)
        if not sid and not name:
            continue
        if "合计" in str(sid) + str(name):
            continue
        amount = val(amount_col) if amount_col is not None else ""
        students.append({
            "studentId": sid,
            "name": name,
            "amount": amount,
        })
    return students


def parse_docx_template(path):
    """读取明细表模板，抽取头部字段与明细行。"""
    doc = docx.Document(path)
    table = doc.tables[0]._tbl
    rows = table.findall(qn("w:tr"))

    def texts(row_index):
        return [cell_text(tc) for tc in rows[row_index].findall(qn("w:tc"))]

    header = texts(0)          # 发放单位 | 值 | 财务项目编号 | 值 | 兼职时段 | 值
    project_row = texts(1)     # 项目名称 | 值
    note_cell = rows[2].findall(qn("w:tc"))[0]

    settings = {
        "unitName": header[1] if len(header) > 1 else "",
        "projectCode": header[3] if len(header) > 3 else "",
        "period": header[5] if len(header) > 5 else "",
        "projectName": project_row[1] if len(project_row) > 1 else "",
        "note": "",
    }

    # 说明单元格：第 0 段是灰色提示，第 1 段是正文
    note_paras = note_cell.findall(qn("w:p"))
    if len(note_paras) >= 2:
        settings["note"] = "".join(
            node.text or "" for node in note_paras[1].iter(qn("w:t"))
        ).strip()
    elif note_paras:
        settings["note"] = cell_text(note_cell)

    # 明细行：跳过表头行，直到出现“合 计”
    details = []
    for r in range(5, len(rows)):
        cells = [cell_text(tc) for tc in rows[r].findall(qn("w:tc"))]
        joined = "".join(cells).replace(" ", "")
        if not joined:
            continue
        if "合计" in joined:
            break
        # 列顺序：序号 姓名 学号 所在学院 标准 工时 实发金额
        details.append({
            "name": cells[1] if len(cells) > 1 else "",
            "studentId": cells[2] if len(cells) > 2 else "",
            "college": cells[3] if len(cells) > 3 else "",
            "rate": cells[4] if len(cells) > 4 else "",
            "hours": cells[5] if len(cells) > 5 else "",
            "amount": cells[6] if len(cells) > 6 else "",
            "checked": True,
        })

    return settings, details


def main():
    docx_path = find_first([r"发放说明.*\.docx$", r"\.docx$"])
    xls_path = find_first([r"发放模版.*\.xls$", r"模版.*\.xls$", r"\.xls$"])

    if not docx_path:
        print("未找到 .docx 明细表模板")
        return 1

    settings, details = parse_docx_template(docx_path)
    roster = parse_roster_xls(xls_path) if xls_path else []

    # 模板自带的说明是"写死人数和月份"的样例文字，直接用会过期。
    # 换成占位符版本，导出时自动跟着当前名单和时段走。
    if re.search(r"本人项目组.*?参与科研项目", settings.get("note", "")):
        settings["note"] = (
            "本人项目组{第一位学生}等{人数}位本科生、研究生在{时段}参与科研项目"
            "TCP/IP协议栈的FreeBSD/Linux内核协议栈追踪、车规级SoC安全、"
            "鸿蒙SafetyNet兼容等项目的研究。负责相关研发工作。明细见下表："
        )

    # 模板里的时段（如"2026年6月"）很快就会过期，留空让程序默认用当前月份；
    # 用户在第 1 步点一下就能改。想固定的话，把下面这行去掉即可。
    settings["period"] = ""

    # 用名单文件里的金额补全明细行
    amount_by_id = {}
    amount_by_name = {}
    for item in roster:
        if item["studentId"]:
            amount_by_id[item["studentId"]] = item["amount"]
        if item["name"]:
            amount_by_name[item["name"]] = item["amount"]

    for d in details:
        d["amount"] = (
            amount_by_id.get(d["studentId"])
            or amount_by_name.get(d["name"])
            or d["amount"]
        )
        # 名单文件是学号的权威来源：按姓名匹配上时，以名单里的学号为准
        if d["name"] in amount_by_name:
            for item in roster:
                if item["name"] == d["name"] and item["studentId"]:
                    d["studentId"] = item["studentId"]
                    break

    # 名单里有、明细里没有的学生一并并入
    known_ids = {d["studentId"] for d in details}
    known_names = {d["name"] for d in details}
    for item in roster:
        if item["studentId"] in known_ids or item["name"] in known_names:
            continue
        details.append({
            "name": item["name"],
            "studentId": item["studentId"],
            "college": "",
            "rate": details[0]["rate"] if details else "100",
            "hours": "",
            "amount": item["amount"],
            "checked": False,
        })

    students = []
    for i, d in enumerate(details):
        students.append({
            "id": "s%d" % (i + 1),
            "studentId": d["studentId"],
            "name": d["name"],
            "college": d["college"],
            "rate": d["rate"],
            "hours": d["hours"],
            # amount 留空 = 由 标准 × 工时 自动计算；
            # 只有用户自己手填（manual=true）才会固定下来
            "amount": "",
            "manual": False,
            "checked": bool(d["checked"]),
        })

    out = {
        "_comment": ("由 extract_defaults.py 从现有模板文件自动生成，可手工调整。"
                     "period 留空 = 默认用「当前月份」；periodPresets 留空 = 默认本月起 6 个月。"),
        "_source": {
            "docx": os.path.basename(docx_path),
            "xls": os.path.basename(xls_path) if xls_path else None,
        },
        "settings": settings,
        "batch": {
            "rate": "100",
            "hours": "10",
            "college": "",          # 留空 = 用 store.DEFAULT_COLLEGE 兜底
        },
        # 补贴费预设，界面上点一下就能套用
        "feePresets": ["500", "800", "1000", "1200", "1500", "1600", "2000"],
        "periodPresets": [],        # 留空 = 默认本月起往后 6 个月
        "students": students,
    }

    target_dir = os.path.join(HERE, "templates")
    os.makedirs(target_dir, exist_ok=True)
    target = os.path.join(target_dir, "defaults.json")
    with open(target, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=2)

    print("已写入 %s" % target)
    print("  模板文件: %s" % os.path.basename(docx_path))
    print("  名单文件: %s" % (os.path.basename(xls_path) if xls_path else "无"))
    print("  头部字段: 发放单位=%s 项目编号=%s 兼职时段=%s" % (
        settings["unitName"], settings["projectCode"], settings["period"]))
    print("  学生人数: %d（已勾选 %d）" % (
        len(students), sum(1 for s in students if s["checked"])))
    for s in students:
        print("    - %s %s %s 标准=%s 工时=%s 实发=%s" % (
            s["studentId"], s["name"], s["college"], s["rate"], s["hours"], s["amount"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
