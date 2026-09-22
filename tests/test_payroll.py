import copy
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import threading
import unittest
import urllib.request
import urllib.error
import uuid

ROOT = Path(__file__).resolve().parents[1]
# tmp/vendor/pylib is a regenerable local staging area (see tmp/stage_vendor.py);
# it comes first so a half-written vendor/pylib cannot shadow a working import.
sys.path[:0] = [str(ROOT / "tmp" / "vendor" / "pylib"), str(ROOT), str(ROOT / "vendor" / "pylib")]
import people_store
import store
import docx_gen
import xls_gen
import pdf_gen
import server
from docx import Document
from docx.oxml.ns import qn
import xlrd


def person(pid="p1", sid="20260001", name="测试甲", **extra):
    return dict(id=pid, studentId=sid, name=name, college="计算机学院", rate="100", hours="10",
                amount="800", manual=True, reason="参与科研开发与实验验证", **extra)


def make_temp_dir():
    """Sandboxed runs cannot write to the system temp dir, so stage under the repo.

    The directory is created directly instead of via tempfile.mkdtemp, because the
    0o700 ACL mkdtemp applies to its private directory is not writable here.
    """
    path = ROOT / "tmp" / "tests" / ("case-%d-%s" % (os.getpid(), uuid.uuid4().hex[:8]))
    path.mkdir(parents=True, exist_ok=True)
    return path


def drop_temp_dir(path):
    shutil.rmtree(path, ignore_errors=True)


class PersistentLibraryTests(unittest.TestCase):
    def setUp(self):
        self.temp = make_temp_dir()
        self.path = str(self.temp / "people.sqlite3")
        self.legacy = {"settings": {"period": "2026年9月"}, "students": [person()], "step": 2}
        people_store.initialize(self.path, self.legacy)

    def tearDown(self):
        drop_temp_dir(self.temp)

    def mutate(self, action, **body):
        return people_store.mutate(self.path, action, dict(body, revision=people_store.load(self.path)["revision"]))

    def test_migration_once_and_sequential_additions_survive_reopen(self):
        self.mutate("person.save", person=person("p2", "", "同名"))
        self.mutate("person.save", person=person("p3", "", "同名"))
        people_store.initialize(self.path, {"students": []})
        state = people_store.load(self.path)
        self.assertEqual([p["id"] for p in state["library"]["people"]], ["p1", "p2", "p3"])
        self.mutate("person.save", person={"id": "p2", "name": "改名后", "customFields": {"备注": "长期"}})
        self.assertEqual(people_store.load(self.path)["library"]["people"][1]["name"], "改名后")

    def test_empty_workspace_does_not_delete_people(self):
        self.mutate("workspace", workspace={"students": [], "settings": {"period": "2026年10月"}})
        state = people_store.load(self.path)
        self.assertEqual(state["students"], [])
        self.assertEqual(len(state["library"]["people"]), 1)
        self.mutate("workspace", workspace={"students": [dict(state["library"]["people"][0], checked=True)]})
        self.assertEqual(people_store.load(self.path)["students"][0]["id"], "p1")

    def test_custom_groups_multiple_memberships_rename_delete_keep_people(self):
        state = self.mutate("group.save", name="博士生")
        groups = state["library"]["groups"]
        ids = [groups[0]["id"], groups[-1]["id"]]
        self.mutate("person.save", person={"id": "p1", "groupIds": ids})
        self.mutate("group.save", id=ids[1], name="博士助研")
        state = self.mutate("group.delete", id=ids[1])
        self.assertEqual(state["library"]["people"][0]["groupIds"], [ids[0]])
        self.assertEqual(state["students"][0]["groupIds"], [ids[0]])

    def test_delete_is_explicit_stale_saves_cannot_resurrect(self):
        self.mutate("person.save", person=person("p2", "20260002", "测试乙"))
        stale = people_store.load(self.path)
        state = self.mutate("person.delete", id="p1")
        self.assertEqual([p["id"] for p in state["library"]["people"]], ["p2"])
        self.assertEqual(state["students"], [])
        with self.assertRaises(people_store.ConflictError):
            people_store.mutate(self.path, "workspace", {"revision": stale["revision"], "workspace": stale})
        with self.assertRaises(people_store.ConflictError):
            self.mutate("person.save", person=person())

    def test_failed_transaction_preserves_existing_people_and_revision(self):
        old = people_store.load(self.path)
        bad = person("p2", "20260002", "测试乙", groupIds=["missing"])
        with self.assertRaises(ValueError):
            self.mutate("workspace", workspace={"students": [person(name="不应保存"), bad]})
        self.assertEqual(people_store.load(self.path), old)

    def test_student_id_identity_and_backup_roundtrip(self):
        state = self.mutate("person.save", person=person("new-id", "20260001", "更新姓名"))
        self.assertEqual(len(state["library"]["people"]), 1)
        self.assertEqual(state["library"]["people"][0]["id"], "p1")
        self.mutate("person.save", person={"id": "p1", "groupIds": [state["library"]["groups"][0]["id"]]})
        backup = people_store.backup(self.path)
        other = str(self.temp / "restored.sqlite3")
        people_store.initialize(other, {"students": []})
        restored = people_store.mutate(other, "library.import", {"revision": 0, "backup": backup})
        self.assertEqual(restored["library"]["people"][0]["name"], "更新姓名")
        self.assertEqual(len(restored["library"]["people"][0]["groupIds"]), 1)


class ApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = make_temp_dir()
        cls.old_paths = (store.APP_DIR, store.SETTINGS_PATH, store.DATABASE_PATH, store.OUTPUT_DIR, server.OUTPUT_DIR)
        store.APP_DIR = str(cls.temp)
        store.SETTINGS_PATH = str(cls.temp / "settings.json")
        store.DATABASE_PATH = str(cls.temp / "people.sqlite3")
        store.OUTPUT_DIR = server.OUTPUT_DIR = str(cls.temp / "output")
        cls.http = server.ThreadingHTTPServer(("127.0.0.1", 0), server.Handler)
        cls.thread = threading.Thread(target=cls.http.serve_forever, daemon=True)
        cls.thread.start()
        cls.url = "http://127.0.0.1:%d" % cls.http.server_port

    @classmethod
    def tearDownClass(cls):
        cls.http.shutdown()
        cls.http.server_close()
        cls.thread.join()
        store.APP_DIR, store.SETTINGS_PATH, store.DATABASE_PATH, store.OUTPUT_DIR, server.OUTPUT_DIR = cls.old_paths
        drop_temp_dir(cls.temp)

    def request(self, path, body=None):
        data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
        req = urllib.request.Request(self.url + path, data=data, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req) as response:
            return json.load(response)

    def test_api_persistence_and_reset_keeps_library(self):
        state = self.request("/api/settings")
        response = self.request("/api/people/save", {"revision": state["revision"], "person": person("api-person")})
        state = response["workspace"]
        state["students"] = [person("api-person")]
        saved = self.request("/api/settings", state)["settings"]
        self.assertEqual(saved["students"][0]["name"], "测试甲")
        with self.assertRaises(urllib.error.HTTPError) as error:
            self.request("/api/settings", state)
        self.assertEqual(error.exception.code, 409)
        reset = self.request("/api/settings/reset", {"revision": saved["revision"]})["workspace"]
        self.assertEqual(reset["students"], [])
        self.assertTrue(any(p["name"] == "测试甲" for p in reset["library"]["people"]))

    def test_blank_student_ids_batch_fill_and_exact_total(self):
        students = [person("a", "", "甲"), person("b", "", "乙"), person("c", "", "丙")]
        preview = self.request("/api/batch/preview", {"field": "amount", "students": students, "values": [11,22,33]})
        self.assertEqual([p["amount"] for p in preview["items"]], ["11", "22", "33"])
        preview = self.request("/api/batch/preview", {"field": "amount", "students": students, "total": "100"})
        self.assertEqual([p["amount"] for p in preview["items"]], ["33.34", "33.33", "33.33"])
        self.assertEqual(round(sum(float(p["amount"]) for p in preview["items"]), 2), 100)

    def test_export_both_templates_and_exact_submission_columns(self):
        students = [person("z", "20260003", "测试丙"), person("y", "20260002", "测试乙", checked=False), person("x", "20260001", "测试甲")]
        for kind in ("research", "non_research"):
            with self.subTest(kind=kind):
                payload = {"settings": {"projectType":kind,"period":"2026年9月", "unitName":"测试单位", "projectName":"测试项目"},
                           "students":students, "kinds":["docx","pdf","xls","submission"]}
                results = self.request("/api/export", payload)["results"]
                doc = Document(results["docx"]["path"])
                layout = docx_gen.inspect_template(doc)
                self.assertEqual(layout["research"], kind == "research")
                text = "".join(n.text or "" for n in doc.element.iter(qn("w:t")))
                self.assertIn("测试甲", text)
                self.assertNotIn("测试乙", text)
                self.assertIn("1600", text)
                self.assertLess(text.index("测试甲"), text.index("测试丙"))
                self.assertEqual(Path(results["pdf"]["path"]).read_bytes()[:5], b"%PDF-")
                sheet = xlrd.open_workbook(results["submission"]["path"]).sheet_by_index(0)
                self.assertEqual(sheet.row_values(0), ["学号", "姓名", "助研津贴(三兼费)"])
                self.assertEqual((sheet.nrows,sheet.ncols), (3,3))
                self.assertEqual(sheet.cell_value(1,0), "20260001")
                roster = xlrd.open_workbook(results["xls"]["path"]).sheet_by_index(0)
                self.assertEqual((roster.nrows,roster.ncols), (4,7))
                self.assertEqual(roster.cell_value(2,0), "")

    def test_export_rejects_missing_name_and_missing_submission_id(self):
        for students in ([person(name="")], [person(sid="")]):
            with self.assertRaises(urllib.error.HTTPError) as error:
                self.request("/api/export", {"students":students,"kinds":["submission"]})
            self.assertEqual(error.exception.code,400)


if __name__ == "__main__":
    unittest.main()
