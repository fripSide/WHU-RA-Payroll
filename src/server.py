# -*- coding: utf-8 -*-
"""HTTP 服务：上传名单、编辑、导出 Word / PDF / Excel。"""
import os
import sys
import json
import shutil
import urllib.parse
import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# 本文件在 src/ 下：同级模块从这里导入，web/ 等资源在项目根（上一级）
SRC_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SRC_DIR)
sys.path.insert(0, SRC_DIR)
import bootstrap
bootstrap.ensure_dependencies()

import store
import docx_gen
import xls_gen
import pdf_gen

WEB_DIR = os.path.join(ROOT, "web")
OUTPUT_DIR = store.OUTPUT_DIR
PROJECTS_FILE = os.path.join(store.DATA_DIR, "projects.json")
PEOPLE_FILE = os.path.join(store.DATA_DIR, "people.json")
MAX_UPLOAD = 20 * 1024 * 1024


def load_projects():
    """加载项目列表"""
    if os.path.exists(PROJECTS_FILE):
        try:
            with open(PROJECTS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('projects', [])
        except:
            pass
    return []


def save_projects(projects):
    """保存项目列表"""
    store._ensure_dirs()
    data = {
        "_comment": "项目配置文件 - 简化版使用",
        "projects": projects
    }
    with open(PROJECTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_people_data():
    """加载人员数据"""
    if os.path.exists(PEOPLE_FILE):
        try:
            with open(PEOPLE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {
        "_comment": "人员数据文件 - 使用JSON存储",
        "lastUpdate": datetime.datetime.now().isoformat(),
        "allPeople": [],
        "payrolls": []
    }


def save_people_data(data):
    """保存人员数据"""
    store._ensure_dirs()
    data["lastUpdate"] = datetime.datetime.now().isoformat()
    with open(PEOPLE_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def json_response(handler, data, status=200):
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    handler.wfile.write(body)


def parse_multipart(handler):
    """简化的文件上传解析"""
    content_type = handler.headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        raise ValueError("请使用 multipart/form-data")

    boundary = None
    for part in content_type.split(";"):
        part = part.strip()
        if part.startswith("boundary="):
            boundary = part[9:].strip('"')
    if not boundary:
        raise ValueError("缺少 boundary")

    length = int(handler.headers.get("Content-Length") or 0)
    if length > MAX_UPLOAD:
        raise ValueError("文件过大（上限 20MB）")

    raw = handler.rfile.read(length)
    marker = b"--" + boundary.encode("utf-8")

    files = {}
    for chunk in raw.split(marker):
        if not chunk or chunk.startswith(b"--"):
            continue
        chunk = chunk.lstrip(b"\r\n")
        split_at = chunk.find(b"\r\n\r\n")
        if split_at < 0:
            continue

        header_blob = chunk[:split_at].decode("utf-8", "replace")
        content = chunk[split_at + 4:]
        if content.endswith(b"\r\n"):
            content = content[:-2]

        name = filename = None
        for line in header_blob.split("\r\n"):
            if "content-disposition" in line.lower():
                for item in line.split(";"):
                    item = item.strip()
                    if item.startswith("name="):
                        name = item[5:].strip('"')
                    elif item.startswith("filename="):
                        filename = item[9:].strip('"')

        if name and filename:
            files[name] = {"filename": filename, "content": content}

    return files


class SimpleHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        if self.path.startswith("/api/"):
            sys.stderr.write(f"[{self.command}] {self.path}\n")

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        route = parsed.path

        try:
            # 获取项目列表
            if route == "/api/projects":
                projects = load_projects()
                return json_response(self, {"projects": projects})

            # 获取人员数据
            if route == "/api/people":
                data = load_people_data()
                return json_response(self, {
                    "allPeople": data.get("allPeople", []),
                    "payrolls": data.get("payrolls", [])
                })

            # 下载导出的文件
            if route == "/api/export/download":
                query = urllib.parse.parse_qs(parsed.query)
                filename = (query.get("file") or [""])[0]
                filepath = (query.get("path") or [""])[0]

                if not filename or not filepath:
                    return json_response(self, {"error": "缺少文件名或路径"}, 400)

                if not os.path.isfile(filepath):
                    return json_response(self, {"error": "文件不存在"}, 404)

                # 发送文件
                self.send_response(200)
                self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Disposition",
                    f'attachment; filename="{urllib.parse.quote(filename)}"')
                self.send_header("Content-Length", str(os.path.getsize(filepath)))
                self.end_headers()

                with open(filepath, "rb") as f:
                    shutil.copyfileobj(f, self.wfile)

                # 下载完删除临时文件和目录
                try:
                    os.remove(filepath)
                    temp_dir = os.path.dirname(filepath)
                    if not os.listdir(temp_dir):  # 目录空了就删
                        os.rmdir(temp_dir)
                except Exception as e:
                    print(f"清理临时文件失败: {e}")

                return

            # 静态文件
            if route == "/" or route == "":
                route = "/index.html"

            filepath = os.path.join(WEB_DIR, route.lstrip("/"))
            if os.path.isfile(filepath):
                self.send_response(200)
                if filepath.endswith(".html"):
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                elif filepath.endswith(".js"):
                    self.send_header("Content-Type", "application/javascript")
                elif filepath.endswith(".css"):
                    self.send_header("Content-Type", "text/css")
                self.end_headers()
                with open(filepath, "rb") as f:
                    shutil.copyfileobj(f, self.wfile)
                return

            return json_response(self, {"error": "Not Found"}, 404)

        except Exception as e:
            return json_response(self, {"error": str(e)}, 500)

    def do_POST(self):
        route = urllib.parse.urlparse(self.path).path

        try:
            # 上传 Excel
            if route == "/api/roster/upload":
                files = parse_multipart(self)
                file_obj = files.get("file")
                if not file_obj:
                    return json_response(self, {"error": "没有收到文件"}, 400)

                # 保存临时文件
                filename = file_obj["filename"]
                ext = os.path.splitext(filename)[1]
                tmp_path = os.path.join(OUTPUT_DIR, f"temp_upload{ext}")

                with open(tmp_path, "wb") as f:
                    f.write(file_obj["content"])

                try:
                    # 解析 Excel
                    students = parse_excel(tmp_path)
                    return json_response(self, {
                        "students": students,
                        "source": filename
                    })
                finally:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)

            # 简化导出
            if route == "/api/export/simple":
                length = int(self.headers.get("Content-Length") or 0)
                body = json.loads(self.rfile.read(length).decode("utf-8"))

                project_type = body.get("projectType", "research")
                period = body.get("period", "")
                unit_name = body.get("unitName", "")
                project_code = body.get("projectCode", "")
                project_name = body.get("projectName", "")
                note = body.get("note", "")
                students = body.get("students", [])
                list_name = body.get("listName", "")  # 获取上传的文件名

                if not students:
                    return json_response(self, {"error": "没有学生数据"}, 400)

                # 生成文件
                result = export_files(
                    project_type, period, unit_name, project_code,
                    project_name, note, students, list_name
                )

                return json_response(self, result)

            # 保存项目
            if route == "/api/projects/save":
                length = int(self.headers.get("Content-Length") or 0)
                body = json.loads(self.rfile.read(length).decode("utf-8"))

                projects = load_projects()
                project = body.get("project", {})

                if not project.get("name") or not project.get("code"):
                    return json_response(self, {"error": "项目名称和编号不能为空"}, 400)

                # 如果有ID，更新；否则新增
                if project.get("id"):
                    for i, p in enumerate(projects):
                        if p.get("id") == project["id"]:
                            projects[i] = project
                            break
                else:
                    project["id"] = f"proj_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
                    projects.append(project)

                save_projects(projects)
                return json_response(self, {"projects": projects})

            # 删除项目
            if route == "/api/projects/delete":
                length = int(self.headers.get("Content-Length") or 0)
                body = json.loads(self.rfile.read(length).decode("utf-8"))

                project_id = body.get("id")
                projects = load_projects()
                projects = [p for p in projects if p.get("id") != project_id]

                save_projects(projects)
                return json_response(self, {"projects": projects})

            # 保存发放列表
            if route == "/api/payrolls/save":
                length = int(self.headers.get("Content-Length") or 0)
                body = json.loads(self.rfile.read(length).decode("utf-8"))

                data = load_people_data()
                payroll = body.get("payroll", {})

                if not payroll.get("name"):
                    payroll["name"] = f"{payroll.get('period', '')}发放"

                # 生成ID
                if not payroll.get("id"):
                    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                    payroll["id"] = f"payroll_{timestamp}"
                    payroll["createTime"] = datetime.datetime.now().isoformat()

                # 更新或新增
                existing = False
                for i, p in enumerate(data["payrolls"]):
                    if p.get("id") == payroll["id"]:
                        data["payrolls"][i] = payroll
                        existing = True
                        break

                if not existing:
                    data["payrolls"].insert(0, payroll)

                # 同步更新总人员名单
                for person in payroll.get("people", []):
                    student_id = person.get("studentId", "").strip()
                    if not student_id:
                        continue

                    # 查找是否已存在
                    found = False
                    for p in data["allPeople"]:
                        if p.get("studentId") == student_id:
                            # 更新信息
                            p["name"] = person.get("name", p.get("name", ""))
                            p["college"] = person.get("college", p.get("college", ""))
                            found = True
                            break

                    # 不存在则添加
                    if not found:
                        person_id = f"person_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{len(data['allPeople'])}"
                        data["allPeople"].append({
                            "id": person_id,
                            "studentId": student_id,
                            "name": person.get("name", ""),
                            "college": person.get("college", "")
                        })

                save_people_data(data)
                return json_response(self, {"payroll": payroll, "payrolls": data["payrolls"]})

            # 删除发放列表
            if route == "/api/payrolls/delete":
                length = int(self.headers.get("Content-Length") or 0)
                body = json.loads(self.rfile.read(length).decode("utf-8"))

                payroll_id = body.get("id")
                data = load_people_data()
                data["payrolls"] = [p for p in data["payrolls"] if p.get("id") != payroll_id]

                save_people_data(data)
                return json_response(self, {"payrolls": data["payrolls"]})

            return json_response(self, {"error": "未知接口"}, 404)

        except Exception as e:
            import traceback
            traceback.print_exc()
            return json_response(self, {"error": str(e)}, 500)


def parse_excel(filepath):
    """解析 Excel 文件，返回学生列表（复用 xls_gen 模块）"""
    students_raw = xls_gen.read_roster(filepath)

    # 转换为简化格式
    students = []
    for s in students_raw:
        students.append({
            "studentId": s.get("studentId", ""),
            "name": s.get("name", ""),
            "college": s.get("college", ""),
            "amount": s.get("amount", "")
        })

    return students


def export_files(project_type, period, unit_name, project_code,
                 project_name, note, students, list_name=""):
    """生成 Word 和 Excel 文件（存到临时目录，供前端下载后删除）"""
    import tempfile

    # 使用临时目录
    temp_dir = tempfile.mkdtemp(prefix="rapay_export_")

    # 准备数据
    settings = {
        "projectType": project_type,
        "period": period,
        "unitName": unit_name,
        "projectCode": project_code,
        "projectName": project_name,
        "note": note or "",
        "reasonTemplate": "{身份}{姓名}参与了{项目名称}项目工作。",
        "workContent": ""
    }

    # 转换学生数据格式
    student_list = []
    for s in students:
        amount = float(s.get("amount", 0) or 0)

        # 非科研经费：标准=100，工时=金额/100
        if project_type == "non_research":
            rate = "100"
            hours = str(amount / 100) if amount > 0 else ""
        else:
            rate = ""
            hours = ""

        student_list.append({
            "studentId": s.get("studentId", ""),
            "name": s.get("name", ""),
            "college": s.get("college", "") or "国家网络安全学院",
            "amount": str(amount) if amount > 0 else "",
            "rate": rate,
            "hours": hours,
            "manual": True,
            "checked": True,
            "identityName": "学生",
            "reason": note or ""
        })

    # 生成 Word
    template_path = store.current_docx_template(project_type)

    # 调试：打印学生数据
    print(f"\n=== 调试信息 ===")
    print(f"项目类型: {project_type}")
    print(f"学生数量: {len(student_list)}")
    for i, s in enumerate(student_list[:3]):  # 只打印前3个
        print(f"学生{i+1}: 姓名={s.get('name')}, 学号={s.get('studentId')}, "
              f"金额={s.get('amount')}, 标准={s.get('rate')}, 工时={s.get('hours')}")
    print(f"================\n")

    # 使用安全的文件名
    # 从 period 提取月份，格式如 "2026年9月"
    import re
    month_match = re.search(r'(\d+)月', period)
    month_str = f"{month_match.group(1)}月" if month_match else "未知月"

    # 构建文件名：list_name_月份_劳务费明细
    if list_name:
        base_name = f"{store.safe_filename(list_name)}_{month_str}"
    else:
        base_name = month_str

    word_filename = f"{base_name}_劳务费明细.docx"
    word_path = os.path.join(temp_dir, word_filename)

    docx_info = docx_gen.generate_docx(
        {"settings": settings, "students": student_list, "projectName": project_name},
        word_path,
        template_path=template_path
    )

    # 生成 PDF（用项目自带的 pdf_gen，基于 reportlab，不需要装 Office）
    pdf_filename = f"{base_name}_劳务费明细.pdf"
    pdf_path = os.path.join(temp_dir, pdf_filename)

    try:
        pdf_gen.generate_pdf(word_path, pdf_path)
        pdf_success = os.path.exists(pdf_path)
    except Exception as e:
        print(f"PDF 生成失败: {e}")
        pdf_success = False

    if not pdf_success:
        pdf_filename = None
        pdf_path = None

    # 生成简化版 Excel（只包含学号、姓名、金额）
    excel_filename = f"{base_name}_劳务费名单.xls"
    excel_path = os.path.join(temp_dir, excel_filename)

    write_simple_roster_xls(students, excel_path)

    # 计算统计
    total = sum(docx_gen.to_money(s.get("amount", 0)) for s in students)

    result = {
        "word_file": word_filename,
        "word_path": word_path,
        "excel_file": excel_filename,
        "excel_path": excel_path,
        "temp_dir": temp_dir,
        "count": len(students),
        "total": f"{total:.2f}"
    }

    if pdf_success:
        result["pdf_file"] = pdf_filename
        result["pdf_path"] = pdf_path

    return result


def write_simple_roster_xls(students, out_path):
    """导出简化版名单 Excel - 只包含学号、姓名、金额"""
    import xlwt

    # 按学号排序
    ordered = sorted(students, key=lambda s: (
        1 if not s.get("studentId") else 0,
        s.get("studentId", "")
    ))

    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet("Sheet1")

    # 样式
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
    body_style.alignment.horz = xlwt.Alignment.HORZ_CENTER

    # 学号文本样式 - 避免被识别为数字
    text_style = xlwt.XFStyle()
    text_style.font = body_font
    text_style.alignment = xlwt.Alignment()
    text_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    text_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    text_style.num_format_str = '@'  # @ 表示文本格式

    money_style = xlwt.XFStyle()
    money_style.font = body_font
    money_style.alignment = xlwt.Alignment()
    money_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    money_style.num_format_str = "#,##0.##"

    # 表头：学号、姓名、金额
    headers = ["学号", "姓名", "金额"]
    for col, title in enumerate(headers):
        sheet.write(0, col, title, header_style)

    # 数据行
    for index, student in enumerate(ordered):
        row = index + 1
        # 学号使用文本格式
        sheet.write(row, 0, str(student.get("studentId", "")), text_style)
        sheet.write(row, 1, str(student.get("name", "")), body_style)

        amount = student.get("amount", "")
        if amount:
            try:
                sheet.write(row, 2, float(str(amount).replace(",", "")), money_style)
            except:
                sheet.write(row, 2, str(amount), body_style)
        else:
            sheet.write(row, 2, "", body_style)

    # 列宽
    sheet.col(0).width = 256 * 18  # 学号
    sheet.col(1).width = 256 * 12  # 姓名
    sheet.col(2).width = 256 * 15  # 金额

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    book.save(out_path)
    return {"path": out_path, "students": len(ordered)}


def main():
    import socket
    import webbrowser

    port = 8848
    for attempt in range(10):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), SimpleHandler)
            break
        except OSError:
            port += 1
    else:
        print("无法找到可用端口")
        return

    url = f"http://127.0.0.1:{port}/"
    print(f"\n服务已启动: {url}")
    print("按 Ctrl+C 停止\n")

    webbrowser.open(url)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n正在关闭...")
        server.shutdown()


if __name__ == "__main__":
    main()
