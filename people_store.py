"""长期人员库、分组与本次名单，全部在一次 SQLite 事务里提交。

分组分两层，规则不一样：

* **身份组**（`kind="identity"`，例如 研究生 / 本科生 / AI 补充）
  每人**有且只有一个**，从根上避免同一个人被重复发放。
* **项目组**（`kind="project"`，例如 国家重点研发 / 横向课题 A）
  可以跟身份组重复，一个人能同时属于多个项目组；项目组只用来筛选和批量勾选。

本次名单（workspace）里选中的人只是"引用"，从名单里移除**永远不会**删除人员档案；
只有 delete_person 才会把人从人员库里拿掉（软删除，过期页面也无法把人复活）。
"""
import copy
import datetime
import json
import math
import os
import sqlite3
import uuid
from contextlib import contextmanager

IDENTITY = "identity"
PROJECT = "project"


class ConflictError(ValueError):
    pass


def _now():
    return datetime.datetime.now().isoformat(timespec="seconds")


def _dump(value):
    return json.dumps(value, ensure_ascii=False, allow_nan=False)


def _text(value):
    return str(value if value is not None else "").strip()


@contextmanager
def _connection(path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    db = sqlite3.connect(path, timeout=20, isolation_level=None)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA foreign_keys=ON")
    try:
        db.execute("BEGIN IMMEDIATE")
        _schema(db)
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


def _schema(db):
    db.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
    db.execute("CREATE TABLE IF NOT EXISTS groups (id TEXT PRIMARY KEY, name TEXT NOT NULL UNIQUE)")
    db.execute("""CREATE TABLE IF NOT EXISTS people (
        id TEXT PRIMARY KEY, student_id TEXT NOT NULL, data TEXT NOT NULL,
        deleted INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL, updated_at TEXT NOT NULL)""")
    db.execute("""CREATE UNIQUE INDEX IF NOT EXISTS active_student_id
        ON people(student_id) WHERE student_id <> '' AND deleted = 0""")
    for statement in ("ALTER TABLE groups ADD COLUMN kind TEXT NOT NULL DEFAULT 'identity'",
                      "ALTER TABLE groups ADD COLUMN position INTEGER NOT NULL DEFAULT 0",
                      "ALTER TABLE people ADD COLUMN identity_id TEXT NOT NULL DEFAULT ''",
                      "ALTER TABLE people ADD COLUMN project_ids TEXT NOT NULL DEFAULT '[]'"):
        try:
            db.execute(statement)
        except sqlite3.OperationalError:
            pass  # 列已存在
    _normalize_groups(db)


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


# ------------------------------------------------------------------ 分组

def _rows(db):
    return db.execute("SELECT * FROM groups ORDER BY position, rowid").fetchall()


def _normalize_groups(db):
    """一次性归位：早期的分组没有身份/项目之分。

    - 用 memberships 把每个人的分组归属搬进 identity_id / project_ids
      （第一个身份组当身份，其余自建分组当项目组）。
    - 为 研究生 / 本科生 / AI 补充 补建身份组并归位，让新库一打开就有固定清单。
    """
    if _meta(db, "groupsNormalized"):
        return
    for row in _rows(db):
        try:
            db.execute("UPDATE groups SET position=? WHERE id=?", (row["rowid"], row["id"]))
        except (sqlite3.OperationalError, IndexError):
            pass
    groups = _rows(db)
    by_name = {row["name"]: row["id"] for row in groups}
    identity_ids = set()
    for name in ("研究生", "本科生", "AI 补充"):
        gid = by_name.get(name)
        if not gid:
            gid = uuid.uuid4().hex
            db.execute("INSERT INTO groups(id,name,kind,position) VALUES (?,?,?,?)",
                       (gid, name, IDENTITY, len(groups)))
        else:
            db.execute("UPDATE groups SET kind=? WHERE id=?", (IDENTITY, gid))
        identity_ids.add(gid)

    memberships = _memberships(db)
    for row in db.execute("SELECT * FROM people WHERE deleted=0").fetchall():
        ids = memberships.get(row["id"], [])
        identity = next((g for g in ids if g in identity_ids), "")
        if not identity and ids:
            identity = ids[0]                      # 自建分组当身份，避免出现无身份的人
        projects = [g for g in ids if g != identity]
        db.execute("UPDATE people SET identity_id=?, project_ids=? WHERE id=?",
                   (identity, _dump(projects), row["id"]))
    _set_meta(db, "groupsNormalized", True)


def _memberships(db):
    """读取旧版的 person_id -> [group_id] 关系（新版不再写入）。"""
    try:
        rows = db.execute("SELECT person_id, group_id FROM memberships").fetchall()
    except sqlite3.OperationalError:
        return {}
    out = {}
    for row in rows:
        out.setdefault(row["person_id"], []).append(row["group_id"])
    return out


def _group_kind(db, gid):
    row = db.execute("SELECT kind FROM groups WHERE id=?", (gid,)).fetchone()
    return row["kind"] if row else None


def _require_kind(db, gid, kind, label):
    if not gid:
        return ""
    actual = _group_kind(db, gid)
    if actual is None:
        raise ValueError("%s不存在：%s" % (label, gid))
    if actual != kind:
        raise ValueError("「%s」不是%s" % (
            db.execute("SELECT name FROM groups WHERE id=?", (gid,)).fetchone()["name"], label))
    return gid


def _assign_groups(db, pid, identity, projects):
    db.execute("UPDATE people SET identity_id=?, project_ids=? WHERE id=?",
               (identity, _dump(projects), pid))


# ------------------------------------------------------------------ 人员

def _person(db, row):
    data = json.loads(row["data"])
    data.update(id=row["id"], createdAt=row["created_at"], updatedAt=row["updated_at"])
    data["identityId"] = row["identity_id"] or ""
    try:
        data["projectIds"] = json.loads(row["project_ids"] or "[]")
    except ValueError:
        data["projectIds"] = []
    return data


def _group_selection(db, incoming, data):
    """把请求里的分组归属规整成 (身份, [项目组])。

    兼容两种写法：新的 identityId / projectIds，以及旧的 groupIds 数组
    （数组里第一个身份组当身份，其余当项目组）。
    """
    legacy = incoming.get("groupIds")
    if legacy is not None and "identityId" not in incoming and "projectIds" not in incoming:
        if not isinstance(legacy, list) or any(not isinstance(g, str) for g in legacy):
            raise ValueError("groupIds 必须是分组 ID 数组")
        identity = ""
        for gid in legacy:
            if _group_kind(db, gid) == IDENTITY:
                identity = gid
                break
        projects = [g for g in legacy if g != identity and _group_kind(db, g) == PROJECT]
        return _require_kind(db, identity, IDENTITY, "身份组"), \
            [_require_kind(db, g, PROJECT, "项目组") for g in dict.fromkeys(projects)]

    identity = _require_kind(db, _text(incoming.get("identityId", data.get("identityId"))),
                             IDENTITY, "身份组")
    if "projectIds" in incoming:
        raw = incoming["projectIds"]
        if not isinstance(raw, list) or any(not isinstance(g, str) for g in raw):
            raise ValueError("projectIds 必须是分组 ID 数组")
        projects = [_require_kind(db, g, PROJECT, "项目组") for g in dict.fromkeys(raw)]
    else:
        projects = list(data.get("projectIds") or [])
    return identity, projects


def _upsert(db, incoming):
    if not isinstance(incoming, dict):
        raise ValueError("人员必须为对象")
    pid = _text(incoming.get("id"))
    row = db.execute("SELECT * FROM people WHERE id=?", (pid,)).fetchone() if pid else None
    if row and row["deleted"]:
        raise ConflictError("这位人员已被删除，请刷新后重新添加。")
    sid = _text(incoming.get("studentId"))
    same_sid = db.execute("SELECT * FROM people WHERE student_id=? AND deleted=0", (sid,)).fetchone() if sid else None
    if row and same_sid and row["id"] != same_sid["id"]:
        raise ValueError("学号 %s 已属于另一位人员" % sid)
    row = row or same_sid
    data = json.loads(row["data"]) if row else {}
    for field in ("studentId", "name", "college", "rate", "hours", "amount", "reason"):
        if field in incoming:
            data[field] = _text(incoming[field])
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

    current = dict(data)
    current["identityId"] = row["identity_id"] if row else ""
    if row:
        try:
            current["projectIds"] = json.loads(row["project_ids"] or "[]")
        except ValueError:
            current["projectIds"] = []
    identity, projects = _group_selection(db, incoming, current)
    data["projectIds"] = projects
    data.pop("identityId", None)

    pid = row["id"] if row else (pid or uuid.uuid4().hex)
    if len(pid) > 200:
        raise ValueError("人员 ID 过长")
    stamp = _now()
    if row:
        db.execute("UPDATE people SET student_id=?,data=?,updated_at=?,identity_id=?,project_ids=? WHERE id=?",
                   (data["studentId"], _dump(data), stamp, identity, _dump(projects), pid))
    else:
        db.execute("""INSERT INTO people(id,student_id,data,deleted,created_at,updated_at,identity_id,project_ids)
                      VALUES (?,?,?,0,?,?,?,?)""",
                   (pid, data["studentId"], _dump(data), stamp, stamp, identity, _dump(projects)))
    return _person(db, db.execute("SELECT * FROM people WHERE id=?", (pid,)).fetchone())


# ------------------------------------------------------------------ 本次名单

def _save_workspace(db, workspace):
    workspace = copy.deepcopy(workspace)
    workspace["students"] = _upsert_many(db, workspace.get("students", []))
    workspace["savedAt"] = _now()
    _set_meta(db, "workspace", workspace)


def _check_identity_import(db, gid, people):
    """导入身份组前先查重：同一个人不能同时出现在两个身份组里。

    这里**不是**静默改归属，而是把整份导入拦下来并说清是谁，
    免得用户以为导进去了、其实人被悄悄挪走了。
    """
    name = db.execute("SELECT name FROM groups WHERE id=?", (gid,)).fetchone()
    label = name["name"] if name else gid
    for item in people:
        sid = _text(item.get("studentId"))
        if not sid:
            continue
        row = db.execute("SELECT p.id, p.identity_id, g.name AS gname FROM people p"
                         " LEFT JOIN groups g ON g.id = p.identity_id"
                         " WHERE p.student_id=? AND p.deleted=0", (sid,)).fetchone()
        if row and row["identity_id"] and row["identity_id"] != gid:
            raise ValueError(
                "学号 %s（%s）已经在「%s」组里，不能同时进「%s」。"
                "请先在人员库把它移出原分组，或用「编辑」改归属。"
                % (sid, _text(item.get("name")) or "未填姓名", row["gname"] or "其他分组", label))


def _upsert_many(db, incoming_list):
    """保存一串人员，返回归一化后的结果；未填完的空行按草稿保留。"""
    students, seen = [], set()
    for incoming in incoming_list:
        # 还没填完的空行只是草稿，不算正式人员。
        if not _text(incoming.get("name")) and not _text(incoming.get("studentId")):
            if db.execute("SELECT 1 FROM people WHERE id=?", (_text(incoming.get("id")),)).fetchone():
                raise ValueError("已有人员的姓名和学号不能同时清空；删除请到人员库操作")
            students.append(incoming)
            continue
        person = _upsert(db, incoming)
        if person["id"] in seen:
            raise ValueError("本次名单中同一位人员出现多次，请按学号合并")
        seen.add(person["id"])
        person["checked"] = incoming.get("checked", True) is not False
        students.append(person)
    return students


def _by_student_id(person):
    """人员库排序：学号从小到大，没填学号的排最后。

    学号按文本比——它常常超过 15 位，转成数字会丢精度。
    """
    sid = _text(person.get("studentId"))
    return (1, "") if sid == "" else (0, sid)


def _snapshot(db):
    people = [_person(db, row) for row in db.execute("SELECT * FROM people WHERE deleted=0 ORDER BY rowid")]
    people.sort(key=_by_student_id)
    by_id = {p["id"]: p for p in people}
    workspace = _meta(db, "workspace", {})
    students = []
    for draft in workspace.get("students", []):
        person = by_id.get(_text(draft.get("id")))
        if person:
            students.append(dict(person, checked=draft.get("checked", True)))
        elif not _text(draft.get("name")) and not _text(draft.get("studentId")):
            students.append(draft)      # 未填完的空白行
    workspace["students"] = students
    workspace["revision"] = _meta(db, "revision", 0)
    workspace["apiVersion"] = 2
    workspace["library"] = {
        "groups": [{"id": r["id"], "name": r["name"], "kind": r["kind"], "position": r["position"]}
                   for r in _rows(db)],
        "people": people,
    }
    return workspace


def initialize(path, legacy_workspace):
    with _connection(path) as db:
        if _meta(db, "initialized"):
            return
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
            pid = _text(payload.get("id"))
            count = db.execute("UPDATE people SET deleted=1,updated_at=? WHERE id=? AND deleted=0",
                               (_now(), pid)).rowcount
            if not count:
                raise ValueError("人员不存在或已删除")
        elif action == "people.save":
            people = payload.get("people")
            if not isinstance(people, list) or not people:
                raise ValueError("请提供要保存的人员列表")
            for item in people:
                _upsert(db, item)
        elif action == "people.delete":
            ids = payload.get("ids")
            if not isinstance(ids, list) or not ids:
                raise ValueError("请提供要删除的人员")
            for pid in ids:
                db.execute("UPDATE people SET deleted=1,updated_at=? WHERE id=? AND deleted=0",
                           (_now(), _text(pid)))
        elif action == "group.roster":
            # 把一整份名单导入某个分组：整批人一起写，中途出错则全部回滚。
            kind = IDENTITY if payload.get("kind") == IDENTITY else PROJECT
            gid = _require_kind(db, _text(payload.get("id")), kind,
                                "身份组" if kind == IDENTITY else "项目组")
            people = payload.get("people")
            if not isinstance(people, list) or not people:
                raise ValueError("请提供要导入的人员名单")
            bulk = [dict(item, **({"identityId": gid} if kind == IDENTITY else {"projectIds": [gid]}))
                    for item in people]
            if kind == IDENTITY:
                _check_identity_import(db, gid, bulk)
            _upsert_many(db, bulk)
            if payload.get("replace"):
                keep = set()
                for item in bulk:
                    row = db.execute("SELECT id FROM people WHERE student_id=? AND deleted=0",
                                     (_text(item.get("studentId")),)).fetchone() if _text(item.get("studentId")) else None
                    if row:
                        keep.add(row["id"])
                column = "identity_id" if kind == IDENTITY else "project_ids"
                for row in db.execute("SELECT id, %s AS g FROM people WHERE deleted=0" % column).fetchall():
                    if kind == IDENTITY:
                        if row["g"] == gid and row["id"] not in keep:
                            db.execute("UPDATE people SET identity_id='' WHERE id=?", (row["id"],))
                    else:
                        try:
                            current = json.loads(row["g"] or "[]")
                        except ValueError:
                            current = []
                        if gid in current and row["id"] not in keep:
                            db.execute("UPDATE people SET project_ids=? WHERE id=?",
                                       (_dump([g for g in current if g != gid]), row["id"]))
        elif action == "group.save":
            name = _text(payload.get("name"))
            if not name or len(name) > 80:
                raise ValueError("分组名称须为 1–80 个字符")
            kind = IDENTITY if payload.get("kind") == IDENTITY else PROJECT
            gid = _text(payload.get("id"))
            duplicate = db.execute("SELECT id FROM groups WHERE name=?", (name,)).fetchone()
            if duplicate and duplicate[0] != gid:
                raise ValueError("已有同名分组")
            if gid:
                if not db.execute("UPDATE groups SET name=?,kind=? WHERE id=?", (name, kind, gid)).rowcount:
                    raise ValueError("分组不存在")
            else:
                position = db.execute("SELECT COALESCE(MAX(position),-1)+1 FROM groups").fetchone()[0]
                db.execute("INSERT INTO groups(id,name,kind,position) VALUES (?,?,?,?)",
                           (uuid.uuid4().hex, name, kind, position))
        elif action == "group.delete":
            gid = _text(payload.get("id"))
            if not db.execute("DELETE FROM groups WHERE id=?", (gid,)).rowcount:
                raise ValueError("分组不存在")
            # 组里的**人**必须留着：只解除归属，不动档案。
            db.execute("UPDATE people SET identity_id='' WHERE identity_id=?", (gid,))
            for row in db.execute("SELECT id, project_ids FROM people WHERE project_ids LIKE ?", ("%" + gid + "%",)).fetchall():
                try:
                    kept = [g for g in json.loads(row["project_ids"] or "[]") if g != gid]
                except ValueError:
                    kept = []
                db.execute("UPDATE people SET project_ids=? WHERE id=?", (_dump(kept), row["id"]))
        elif action == "library.import":
            backup = payload.get("backup") or {}
            if backup.get("format") != "whu-payroll-library" or backup.get("version") not in (1, 2):
                raise ValueError("不是受支持的人员库备份")
            group_map = {}
            for group in backup.get("groups", []):
                name = _text(group.get("name"))
                if not name or len(name) > 80:
                    raise ValueError("备份中的分组名称无效")
                kind = IDENTITY if group.get("kind") == IDENTITY else PROJECT
                existing = db.execute("SELECT id, kind FROM groups WHERE name=?", (name,)).fetchone()
                if existing:
                    gid = existing["id"]
                else:
                    position = db.execute("SELECT COALESCE(MAX(position),-1)+1 FROM groups").fetchone()[0]
                    gid = uuid.uuid4().hex
                    db.execute("INSERT INTO groups(id,name,kind,position) VALUES (?,?,?,?)",
                               (gid, name, kind, position))
                group_map[group["id"]] = gid
            for person in backup.get("people", []):
                incoming = dict(person)
                if "groupIds" in incoming and "identityId" not in incoming:
                    identity = next((g for g in incoming["groupIds"] if _group_kind(db, g) == IDENTITY), "")
                    incoming["identityId"] = group_map.get(identity, identity) if identity else ""
                    incoming["projectIds"] = [group_map.get(g, g) for g in incoming.pop("groupIds")
                                              if _group_kind(db, g) == PROJECT]
                else:
                    incoming["identityId"] = group_map.get(incoming.get("identityId"), incoming.get("identityId") or "")
                    incoming["projectIds"] = [group_map.get(g, g) for g in incoming.get("projectIds", [])]
                old = db.execute("SELECT deleted FROM people WHERE id=?", (_text(incoming.get("id")),)).fetchone()
                if old and old[0]:
                    incoming["id"] = uuid.uuid4().hex
                _upsert(db, incoming)
        else:
            raise ValueError("未知操作")
        _set_meta(db, "revision", _meta(db, "revision", 0) + 1)
        return _snapshot(db)


def backup(path):
    state = load(path)
    return dict(state["library"], format="whu-payroll-library", version=2, exportedAt=_now())
