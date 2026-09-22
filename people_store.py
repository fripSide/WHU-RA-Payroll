"""Persistent people, custom groups and workspace, committed in one SQLite transaction.

The workspace contains a selection of people. Omitting a person from a workspace
never deletes them. Only delete_person removes a record from the active library.
"""
import copy
import datetime
import json
import math
import os
import sqlite3
import uuid
from contextlib import contextmanager


class ConflictError(ValueError):
    pass


def _now():
    return datetime.datetime.now().isoformat(timespec="seconds")


def _dump(value):
    return json.dumps(value, ensure_ascii=False, allow_nan=False)


@contextmanager
def _connection(path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    db = sqlite3.connect(path, timeout=20, isolation_level=None)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA foreign_keys=ON")
    try:
        db.execute("BEGIN IMMEDIATE")
        db.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
        db.execute("CREATE TABLE IF NOT EXISTS groups (id TEXT PRIMARY KEY, name TEXT NOT NULL UNIQUE)")
        db.execute("""CREATE TABLE IF NOT EXISTS people (
            id TEXT PRIMARY KEY, student_id TEXT NOT NULL, data TEXT NOT NULL,
            deleted INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL, updated_at TEXT NOT NULL)""")
        db.execute("""CREATE UNIQUE INDEX IF NOT EXISTS active_student_id
            ON people(student_id) WHERE student_id <> '' AND deleted = 0""")
        db.execute("""CREATE TABLE IF NOT EXISTS memberships (
            person_id TEXT REFERENCES people(id), group_id TEXT REFERENCES groups(id) ON DELETE CASCADE,
            PRIMARY KEY(person_id, group_id))""")
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


def _meta(db, key, fallback=None):
    row = db.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return json.loads(row[0]) if row else fallback


def _set_meta(db, key, value):
    db.execute("INSERT OR REPLACE INTO meta(key,value) VALUES (?,?)", (key, _dump(value)))


def _expect(db, revision):
    if type(revision) is not int:
        raise ValueError("请提供读取数据时返回的 revision")
    if revision != _meta(db, "revision", 0):
        raise ConflictError("数据已在其他页面更新。请刷新后继续；本页未保存内容仍保留在浏览器草稿中。")


def _person(db, row):
    data = json.loads(row["data"])
    data.update(id=row["id"], createdAt=row["created_at"], updatedAt=row["updated_at"])
    data["groupIds"] = [r[0] for r in db.execute(
        "SELECT group_id FROM memberships WHERE person_id=? ORDER BY group_id", (row["id"],))]
    return data


def _upsert(db, incoming):
    if not isinstance(incoming, dict):
        raise ValueError("人员必须为对象")
    pid = str(incoming.get("id") or "").strip()
    row = db.execute("SELECT * FROM people WHERE id=?", (pid,)).fetchone() if pid else None
    if row and row["deleted"]:
        raise ConflictError("这位人员已被删除，请刷新后重新添加。")
    sid = str(incoming.get("studentId") or "").strip()
    same_sid = db.execute("SELECT * FROM people WHERE student_id=? AND deleted=0", (sid,)).fetchone() if sid else None
    if row and same_sid and row["id"] != same_sid["id"]:
        raise ValueError("学号 %s 已属于另一位人员" % sid)
    row = row or same_sid
    data = json.loads(row["data"]) if row else {}
    for field in ("studentId", "name", "college", "rate", "hours", "amount", "reason"):
        if field in incoming:
            data[field] = str(incoming[field] if incoming[field] is not None else "").strip()
        else:
            data.setdefault(field, "")
    if not data["name"] and not data["studentId"]:
        raise ValueError("请至少填写姓名或学号")
    for field in ("rate", "hours", "amount"):
        if data[field]:
            try:
                value = float(data[field].replace(",", ""))
            except ValueError:
                raise ValueError("标准、工时和金额必须是数字")
            if not math.isfinite(value) or value < 0:
                raise ValueError("标准、工时和金额必须为有限的非负数")
    data["manual"] = bool(incoming.get("manual", data.get("manual", bool(data["amount"]))))
    data["collegeAuto"] = bool(incoming.get("collegeAuto", data.get("collegeAuto", False)))
    extra = incoming.get("customFields", data.get("customFields", {}))
    if not isinstance(extra, dict):
        raise ValueError("customFields 必须为对象")
    data["customFields"] = extra
    pid = row["id"] if row else (pid or uuid.uuid4().hex)
    if len(pid) > 200:
        raise ValueError("人员 ID 过长")
    stamp = _now()
    if row:
        db.execute("UPDATE people SET student_id=?,data=?,updated_at=? WHERE id=?",
                   (data["studentId"], _dump(data), stamp, pid))
    else:
        db.execute("INSERT INTO people(id,student_id,data,created_at,updated_at) VALUES (?,?,?,?,?)",
                   (pid, data["studentId"], _dump(data), stamp, stamp))
    if "groupIds" in incoming:
        group_ids = incoming["groupIds"]
        if not isinstance(group_ids, list) or any(not isinstance(g, str) for g in group_ids):
            raise ValueError("groupIds 必须是分组 ID 数组")
        group_ids = list(dict.fromkeys(group_ids))
        for gid in group_ids:
            if not db.execute("SELECT 1 FROM groups WHERE id=?", (gid,)).fetchone():
                raise ValueError("分组不存在：%s" % gid)
        db.execute("DELETE FROM memberships WHERE person_id=?", (pid,))
        db.executemany("INSERT INTO memberships VALUES (?,?)", [(pid, gid) for gid in group_ids])
    return _person(db, db.execute("SELECT * FROM people WHERE id=?", (pid,)).fetchone())


def _save_workspace(db, workspace):
    workspace = copy.deepcopy(workspace)
    students, seen = [], set()
    for incoming in workspace.get("students", []):
        # An unfinished blank row remains a draft, not a permanent empty person.
        if not str(incoming.get("name") or "").strip() and not str(incoming.get("studentId") or "").strip():
            if db.execute("SELECT 1 FROM people WHERE id=?", (str(incoming.get("id") or ""),)).fetchone():
                raise ValueError("已有人员的姓名和学号不能同时清空；删除请到人员库操作")
            students.append(incoming)
            continue
        person = _upsert(db, incoming)
        if person["id"] in seen:
            raise ValueError("本次名单中同一位人员出现多次，请按学号合并")
        seen.add(person["id"])
        person["checked"] = incoming.get("checked", True) is not False
        students.append(person)
    workspace["students"] = students
    workspace["savedAt"] = _now()
    _set_meta(db, "workspace", workspace)


def _snapshot(db):
    people = [_person(db, row) for row in db.execute("SELECT * FROM people WHERE deleted=0 ORDER BY rowid")]
    by_id = {p["id"]: p for p in people}
    workspace = _meta(db, "workspace", {})
    students = []
    for draft in workspace.get("students", []):
        person = by_id.get(draft.get("id"))
        if person:
            students.append(dict(person, checked=draft.get("checked", True)))
        elif not draft.get("name") and not draft.get("studentId"):
            students.append(draft)
    workspace["students"] = students
    workspace["revision"] = _meta(db, "revision", 0)
    workspace["apiVersion"] = 1
    workspace["library"] = {
        "groups": [dict(r) for r in db.execute("SELECT id,name FROM groups ORDER BY rowid")],
        "people": people,
    }
    return workspace


def initialize(path, legacy_workspace):
    with _connection(path) as db:
        if _meta(db, "initialized"):
            return
        for name in ("研究生", "本科生", "AI 补充"):
            db.execute("INSERT INTO groups VALUES (?,?)", (uuid.uuid4().hex, name))
        _save_workspace(db, legacy_workspace)
        _set_meta(db, "revision", 0)
        _set_meta(db, "initialized", True)


def load(path):
    with _connection(path) as db:
        return _snapshot(db)


def mutate(path, action, payload):
    with _connection(path) as db:
        _expect(db, payload.get("revision"))
        if action == "workspace":
            _save_workspace(db, payload["workspace"])
        elif action == "person.save":
            _upsert(db, payload.get("person") or {})
        elif action == "person.delete":
            pid = str(payload.get("id") or "")
            count = db.execute("UPDATE people SET deleted=1,updated_at=? WHERE id=? AND deleted=0", (_now(), pid)).rowcount
            if not count:
                raise ValueError("人员不存在或已删除")
            db.execute("DELETE FROM memberships WHERE person_id=?", (pid,))
        elif action == "group.save":
            name = str(payload.get("name") or "").strip()
            if not name or len(name) > 80:
                raise ValueError("分组名称须为 1–80 个字符")
            gid = str(payload.get("id") or "")
            duplicate = db.execute("SELECT id FROM groups WHERE name=?", (name,)).fetchone()
            if duplicate and duplicate[0] != gid:
                raise ValueError("已有同名分组")
            if gid:
                if not db.execute("UPDATE groups SET name=? WHERE id=?", (name, gid)).rowcount:
                    raise ValueError("分组不存在")
            else:
                db.execute("INSERT INTO groups VALUES (?,?)", (uuid.uuid4().hex, name))
        elif action == "group.delete":
            if not db.execute("DELETE FROM groups WHERE id=?", (str(payload.get("id") or ""),)).rowcount:
                raise ValueError("分组不存在")
        elif action == "library.import":
            backup = payload.get("backup") or {}
            if backup.get("format") != "whu-payroll-library" or backup.get("version") != 1:
                raise ValueError("不是受支持的人员库备份")
            group_map = {}
            for group in backup.get("groups", []):
                name = str(group.get("name") or "").strip()
                if not name or len(name) > 80:
                    raise ValueError("备份中的分组名称无效")
                existing = db.execute("SELECT id FROM groups WHERE name=?", (name,)).fetchone()
                gid = existing[0] if existing else uuid.uuid4().hex
                if not existing:
                    db.execute("INSERT INTO groups VALUES (?,?)", (gid, name))
                group_map[group["id"]] = gid
            for person in backup.get("people", []):
                incoming = dict(person)
                old = db.execute("SELECT deleted FROM people WHERE id=?", (str(incoming.get("id") or ""),)).fetchone()
                if old and old[0]:
                    incoming["id"] = uuid.uuid4().hex
                incoming["groupIds"] = [group_map[g] for g in incoming.get("groupIds", [])]
                _upsert(db, incoming)
        else:
            raise ValueError("未知操作")
        _set_meta(db, "revision", _meta(db, "revision", 0) + 1)
        return _snapshot(db)


def backup(path):
    state = load(path)
    return dict(state["library"], format="whu-payroll-library", version=1, exportedAt=_now())
