# -*- coding: utf-8 -*-
"""本地服务：提供网页界面 + 生成 Word/Excel 文件。

只用 Python 标准库做 HTTP 服务，文件生成交给 docx_gen / xls_gen。
"""
import os
import sys
import json
import shutil
import socket
import tempfile
import threading
import webbrowser
import mimetypes
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "vendor", "pylib"))

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import store            # noqa: E402
import docx_gen         # noqa: E402
import xls_gen          # noqa: E402

WEB_DIR = os.path.join(HERE, "web")
OUTPUT_DIR = store.OUTPUT_DIR
MAX_UPLOAD = 20 * 1024 * 1024


def json_response(handler, data, status=200):
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    try:
        handler.send_response(status)
        handler.send_header("Content-Type", "application/json; charset=utf-8")
        handler.send_header("Content-Length", str(len(body)))
        handler.send_header("Cache-Control", "no-store")
        handler.end_headers()
        handler.wfile.write(body)
    except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError):
        pass


def read_body(handler):
    length = int(handler.headers.get("Content-Length") or 0)
    if length <= 0:
        return b""
    if length > MAX_UPLOAD:
        raise ValueError("上传内容过大（上限 20MB）")
    return handler.rfile.read(length)


def parse_json_body(handler):
    raw = read_body(handler)
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def parse_multipart(handler):
    """极简 multipart/form-data 解析，返回 (fields, files)。"""
    content_type = handler.headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        raise ValueError("请使用 multipart/form-data 上传文件")

    boundary = None
    for part in content_type.split(";"):
        part = part.strip()
        if part.startswith("boundary="):
            boundary = part[len("boundary="):].strip('"')
    if not boundary:
        raise ValueError("缺少 boundary")

    raw = read_body(handler)
    marker = b"--" + boundary.encode("utf-8")
    fields, files = {}, {}
    for chunk in raw.split(marker):
        if not chunk or chunk in (b"--", b"--\r\n", b"\r\n"):
            continue
        chunk = chunk.lstrip(b"\r\n")
        if chunk.startswith(b"--"):
            continue
        split_at = chunk.find(b"\r\n\r\n")
        if split_at < 0:
            continue
        header_blob = chunk[:split_at].decode("utf-8", "replace")
        content = chunk[split_at + 4:]
        if content.endswith(b"\r\n"):
            content = content[:-2]

        name = filename = None
        for line in header_blob.split("\r\n"):
            if line.lower().startswith("content-disposition"):
                for item in line.split(";"):
                    item = item.strip()
                    if item.startswith("name="):
                        name = item[5:].strip('"')
                    elif item.startswith("filename="):
                        filename = item[9:].strip('"')
        if name is None:
            continue
        if filename:
            files[name] = {"filename": filename, "content": content}
        else:
            fields[name] = content.decode("utf-8", "replace")
    return fields, files


class Handler(BaseHTTPRequestHandler):
    server_version = "BaoXiaoHelper/1.0"
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        if self.path.startswith("/api/"):
            sys.stderr.write("[api] %s %s\n" % (self.command, self.path))

    # 浏览器关页面/取消请求时会直接断开连接，属于正常现象，
    # 不要在控制台里抛一堆 ConnectionResetError 吓人。
    def handle_one_request(self):
        try:
            BaseHTTPRequestHandler.handle_one_request(self)
        except (ConnectionResetError, ConnectionAbortedError,
                BrokenPipeError, TimeoutError):
            self.close_connection = True

    def handle(self):
        try:
            BaseHTTPRequestHandler.handle(self)
        except (ConnectionResetError, ConnectionAbortedError,
                BrokenPipeError, TimeoutError):
            pass

    def finish(self):
        try:
            BaseHTTPRequestHandler.finish(self)
        except (ConnectionResetError, ConnectionAbortedError,
                BrokenPipeError, TimeoutError, OSError):
            pass

    # ------------------------------------------------------------- 静态文件
    def _serve_file(self, path, download_name=None):
        if not os.path.isfile(path):
            self.send_error(404, "Not Found")
            return
        ctype = mimetypes.guess_type(path)[0] or "application/octet-stream"
        if path.lower().endswith((".js", ".css", ".html")):
            ctype += "; charset=utf-8"
        size = os.path.getsize(path)
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(size))
        self.send_header("Cache-Control", "no-store")
        if download_name:
            quoted = urllib.parse.quote(download_name)
            self.send_header(
                "Content-Disposition",
                "attachment; filename=\"%s\"; filename*=UTF-8''%s" % (
                    download_name.encode("ascii", "ignore").decode("ascii") or "file",
                    quoted))
        self.end_headers()
        try:
            with open(path, "rb") as fh:
                shutil.copyfileobj(fh, self.wfile)
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError):
            pass

    def _serve_static(self, url_path):
        rel = url_path.lstrip("/") or "index.html"
        target = os.path.normpath(os.path.join(WEB_DIR, rel))
        if not target.startswith(WEB_DIR):
            self.send_error(403, "Forbidden")
            return
        if os.path.isdir(target):
            target = os.path.join(target, "index.html")
        self._serve_file(target)

    # ------------------------------------------------------------------ GET
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        route = parsed.path
        query = urllib.parse.parse_qs(parsed.query)
        try:
            if route == "/api/ping":
                return json_response(self, {"ok": True})
            if route == "/api/settings":
                return json_response(self, store.load_settings())
            if route == "/api/templates":
                return json_response(self, {"templates": store.list_templates()})
            if route == "/api/export/download":
                name = (query.get("file") or [""])[0]
                safe = store.safe_filename(name)
                path = os.path.join(OUTPUT_DIR, safe)
                if not os.path.isfile(path):
                    return json_response(self, {"error": "文件不存在：%s" % safe}, 404)
                return self._serve_file(path, download_name=safe)
            if route.startswith("/api/"):
                return json_response(self, {"error": "未知接口：%s" % route}, 404)
            return self._serve_static(route)
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError):
            return
        except Exception as exc:  # noqa: BLE001
            return json_response(self, {"error": str(exc)}, 500)

    # ----------------------------------------------------------------- POST
    def do_POST(self):
        route = urllib.parse.urlparse(self.path).path
        try:
            if route == "/api/settings":
                return json_response(self, {"ok": True,
                                            "settings": store.save_settings(parse_json_body(self))})
            if route == "/api/settings/reset":
                return json_response(self, {"ok": True, "workspace": store.reset_settings()})
            if route == "/api/templates/save":
                payload = parse_json_body(self)
                info = store.save_template((payload.get("id") or "").strip(), payload)
                return json_response(self, {"ok": True, "template": info,
                                            "templates": store.list_templates()})
            if route == "/api/templates/delete":
                payload = parse_json_body(self)
                removed = store.delete_template((payload.get("id") or "").strip())
                return json_response(self, {"ok": removed, "templates": store.list_templates()})
            if route == "/api/templates/load":
                payload = parse_json_body(self)
                data = store.load_template((payload.get("id") or "").strip())
                return json_response(self, {"ok": True, "template": data})
            if route == "/api/roster/upload":
                return self._roster_upload()
            if route == "/api/roster/import":
                return self._roster_text()
            if route == "/api/template/docx":
                return self._docx_template_upload()
            if route == "/api/batch/preview":
                return self._batch_preview()
            if route == "/api/export":
                return self._export()
            return json_response(self, {"error": "未知接口：%s" % route}, 404)
        except (ConnectionResetError, ConnectionAbortedError, BrokenPipeError):
            return
        except json.JSONDecodeError as exc:
            return json_response(self, {"error": "请求数据不是合法 JSON：%s" % exc}, 400)
        except FileNotFoundError as exc:
            return json_response(self, {"error": str(exc)}, 404)
        except Exception as exc:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            return json_response(self, {"error": "%s: %s" % (type(exc).__name__, exc)}, 500)

    # --------------------------------------------------------------- 各接口
    def _roster_upload(self):
        fields, files = parse_multipart(self)
        item = files.get("file")
        if not item:
            return json_response(self, {"error": "没有收到文件"}, 400)
        suffix = os.path.splitext(item["filename"])[1] or ".dat"
        tmp_dir = tempfile.mkdtemp(prefix="roster-")
        tmp_path = os.path.join(tmp_dir, "upload" + suffix)
        with open(tmp_path, "wb") as fh:
            fh.write(item["content"])
        try:
            students = xls_gen.read_roster(tmp_path)
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return json_response(self, {
            "ok": True,
            "source": item["filename"],
            "students": students,
            "merged": fields.get("merge") == "1",
        })

    def _roster_text(self):
        payload = parse_json_body(self)
        students = xls_gen.read_roster_from_text(payload.get("text") or "")
        return json_response(self, {"ok": True, "source": "粘贴文本", "students": students})

    def _batch_preview(self):
        """批量填入一列：按勾选顺序配对，并算出每人配好之后的金额。

        金额计算复用 docx_gen.batch_fill，所以预览里的数字
        和导出文件里的数字一定一致。
        """
        payload = parse_json_body(self)
        field = payload.get("field") or "hours"
        if field not in ("hours", "rate", "amount"):
            return json_response(self, {"error": "不支持的字段：%s" % field}, 400)

        students = [s for s in (payload.get("students") or [])
                    if s.get("checked", True)]
        if not students:
            return json_response(self, {"error": "还没有勾选学生"}, 400)

        # ---- 总额均摊：填一个总额，平均分给勾选的人 ----
        if field == "amount" and payload.get("total") not in (None, ""):
            total = docx_gen.to_money(payload.get("total"))
            if total <= 0:
                return json_response(self, {"error": "总额要大于 0"}, 400)
            unit = docx_gen.fmt_number(total / len(students))
            mapping = {s.get("studentId"): unit for s in students}
            items = docx_gen.batch_fill(students, "amount", mapping)
            return json_response(self, {
                "ok": True, "field": field, "mode": "total",
                "count": len(students), "unit": unit,
                "total": docx_gen.fmt_number(total),
                "keptCount": 0,
                "filledAmount": docx_gen.fmt_number(total),
                "totalAmount": docx_gen.fmt_number(total),
                "items": items,
            })

        # ---- 逐行一列值 ----
        raw = payload.get("values")
        if not isinstance(raw, list):
            raw = str(payload.get("text") or "").splitlines()

        values = []
        for item in raw:
            text = str(item).strip()
            if text == "":
                continue
            try:
                number = float(text.replace(",", ""))
            except ValueError:
                continue          # 非数值直接跳过，不打断整批
            values.append(docx_gen.fmt_number(number))

        if not values:
            return json_response(self, {"error": "没有解析到有效数值"}, 400)
        if len(values) > len(students):
            return json_response(self, {
                "error": "填了 %d 个值，但只勾选了 %d 位学生"
                         % (len(values), len(students))}, 400)

        mapping = {}
        for student, value in zip(students, values):
            mapping[student.get("studentId")] = value

        items = docx_gen.batch_fill(students, field, mapping)

        # 本次真正填进去的那些人的金额小计（区别于"填完后的总计"）
        filled_sum = sum(docx_gen.to_money(i["amount"]) for i in items
                         if i.get("value") is not None)
        kept = [i for i in items if i.get("value") is None]

        return json_response(self, {
            "ok": True, "field": field, "mode": "list",
            "count": len(values),
            "keptCount": len(kept),
            "filledAmount": docx_gen.fmt_number(filled_sum),
            "totalAmount": docx_gen.fmt_number(sum(
                docx_gen.to_money(i["amount"]) for i in items)),
            "items": items,
        })

    def _docx_template_upload(self):
        fields, files = parse_multipart(self)
        item = files.get("file")
        if not item:
            return json_response(self, {"error": "没有收到文件"}, 400)
        if not item["filename"].lower().endswith(".docx"):
            return json_response(self, {"error": "请上传 .docx 文件"}, 400)
        tmp_dir = tempfile.mkdtemp(prefix="docxtpl-")
        tmp_path = os.path.join(tmp_dir, "t.docx")
        with open(tmp_path, "wb") as fh:
            fh.write(item["content"])
        try:
            # 先试读，确认是合法的 Word 且带表格
            from docx import Document
            probe = Document(tmp_path)
            table_count = len(probe.tables)
            settings, students = _extract_from_docx(probe)
            saved = store.ensure_docx_template(
                tmp_path, item["filename"], fields.get("projectType"))
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return json_response(self, {
            "ok": True,
            "file": os.path.basename(saved),
            "tables": table_count,
            "projectType": store.normalize_project_type(fields.get("projectType")),
            "settings": settings,
            "students": students,
        })

    def _export(self):
        payload = parse_json_body(self)
        settings = payload.get("settings") or {}
        students = payload.get("students") or []
        kinds = payload.get("kinds") or ["docx", "xls"]
        label = store.safe_filename(payload.get("label") or "", "").strip("_")
        project_type = store.normalize_project_type(settings.get("projectType"))
        template_path = store.current_docx_template(project_type)

        results = {}
        if "docx" in kinds:
            out = store.next_output_path("docx", settings, label)
            info = docx_gen.generate_docx(
                {"settings": settings, "students": students,
                 "projectType": project_type},
                out, template_path=template_path)
            store.save_settings({"settings": settings,
                                 "batch": payload.get("batch") or {},
                                 "feePresets": payload.get("feePresets"),
                                 "periodPresets": payload.get("periodPresets"),
                                 "students": students})
            results["docx"] = {"file": os.path.basename(info["path"]),
                               "path": info["path"],
                               "students": info["students"],
                               "total": info["total"]}
        if "xls" in kinds:
            roster = _collect_roster(students)
            out = store.next_output_path("xls", settings, label)
            info = xls_gen.write_roster_xls(roster, out)
            results["xls"] = {"file": os.path.basename(info["path"]),
                              "path": info["path"],
                              "students": info["students"]}
        if not results:
            return json_response(self, {"error": "没有可导出的类型"}, 400)
        return json_response(self, {"ok": True, "results": results,
                                    "outputDir": OUTPUT_DIR})

    # ------------------------------------------------------------------ main
    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Allow", "GET, POST, OPTIONS")
        self.send_header("Content-Length", "0")
        self.end_headers()


def _collect_roster(students):
    """名单表：勾选的正常列出，**未勾选的也保留**，作为下次的备选。

    每行带序号、学院、标准、工时，并按学号升序（排序在 xls_gen 里做），
    这样这份名单再被导入时，勾选状态和各项信息都能原样带回来。
    """
    out = []
    for student in students:
        checked = bool(student.get("checked", True))
        amount = docx_gen.effective_amount(student)
        # 没勾的人本来就没填工时，金额按 0 显示反而难看，留空
        if not checked and not amount:
            amount = ""
        out.append({
            "studentId": student.get("studentId", ""),
            "name": student.get("name", ""),
            "college": student.get("college", ""),
            "rate": student.get("rate", ""),
            "hours": student.get("hours", ""),
            # 与 Word 明细表口径一致：手填优先，否则 标准 × 工时
            "amount": amount,
            "checked": checked,
        })
    return out


def _extract_from_docx(document):
    """从上传的 docx 里读出头部字段和明细行，方便一键套用。"""
    from docx.oxml.ns import qn

    def text_of(tc):
        return "".join(node.text or "" for node in tc.iter(qn("w:t"))).strip()

    settings = {"unitName": "", "projectCode": "", "period": "",
                "projectName": "", "note": ""}
    students = []
    if not document.tables:
        return settings, students

    table = document.tables[0]._tbl
    rows = table.findall(qn("w:tr"))
    if not rows:
        return settings, students

    def compact(text):
        return "".join(str(text or "").split()).replace("\u3000", "")

    head_cells = rows[0].findall(qn("w:tc"))
    head = [text_of(tc) for tc in head_cells]
    for i, label in enumerate(head):
        key = compact(label)
        if i + 1 >= len(head):
            continue
        if "发放单位" in key:
            settings["unitName"] = head[i + 1]
        elif "财务项目编号" in key or "项目编号" in key:
            settings["projectCode"] = head[i + 1]
        elif "兼职时段" in key or "发放月份" in key or "月份" in key:
            settings["period"] = head[i + 1]
    if len(rows) > 1:
        second = [text_of(tc) for tc in rows[1].findall(qn("w:tc"))]
        if len(second) > 1:
            settings["projectName"] = second[1]
    # 表头由“序号 + 姓名”识别，两个内置模板的列数和明细起始行不同。
    header_index = None
    labels = []
    for i, row in enumerate(rows):
        vals = [text_of(tc) for tc in row.findall(qn("w:tc"))]
        compacted = [compact(v) for v in vals]
        if any("序号" in v for v in compacted) and any("姓名" in v for v in compacted):
            header_index, labels = i, compacted
            break
    if header_index is None:
        return settings, students
    cols = {}
    for i, label in enumerate(labels):
        if "姓名" in label: cols["name"] = i
        elif "学号" in label: cols["studentId"] = i
        elif "所在学院" in label or label == "学院": cols["college"] = i
        elif "标准" in label or "单价" in label: cols["rate"] = i
        elif "工时" in label or "时长" in label: cols["hours"] = i
        elif "实发金额" in label or "助研津贴" in label or label == "金额": cols["amount"] = i

    for row in rows[header_index + 1:]:
        cells = [text_of(tc) for tc in row.findall(qn("w:tc"))]
        joined = compact("".join(cells))
        if not joined:
            continue
        if "合计" in joined:
            break
        if "name" not in cols or "studentId" not in cols:
            continue
        name = cells[cols["name"]] if cols["name"] < len(cells) else ""
        sid = cells[cols["studentId"]] if cols["studentId"] < len(cells) else ""
        if not name and not sid:
            continue
        students.append({
            "studentId": sid,
            "name": name,
            "college": cells[cols["college"]] if cols.get("college", 999) < len(cells) else "",
            "rate": cells[cols["rate"]] if cols.get("rate", 999) < len(cells) else "",
            "hours": cells[cols["hours"]] if cols.get("hours", 999) < len(cells) else "",
            "amount": cells[cols["amount"]] if cols.get("amount", 999) < len(cells) else "",
            "checked": True,
        })
    return settings, students


def find_port(start=8765, tries=40):
    for port in range(start, start + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("找不到可用端口（%d-%d）" % (start, start + tries))


def main():
    if not os.path.isdir(WEB_DIR):
        print("找不到 web 目录：%s" % WEB_DIR)
        return 1

    store._ensure_dirs()
    port = find_port(int(os.environ.get("BAOXIAO_PORT") or 8765))
    url = "http://127.0.0.1:%d/" % port

    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    server.daemon_threads = True

    print("=" * 58)
    print("  学生劳务费发放助手")
    print("  界面地址： %s" % url)
    print("  输出目录： %s" % OUTPUT_DIR)
    print("  关闭本窗口即可停止服务")
    print("=" * 58)

    if os.environ.get("BAOXIAO_NO_BROWSER") != "1":
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
