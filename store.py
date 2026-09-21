# -*- coding: utf-8 -*-
"""模板与配置存储。

- templates/defaults.json 由现有模板文件提取出的内置默认值
- templates/*.json        用户自建的模板（发放信息 + 默认标准/工时/学院）
- app/settings.json       当前工作台状态（自动保存）
"""
import os
import re
import json
import shutil
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(HERE, "templates")
APP_DIR = os.path.join(HERE, "app")
OUTPUT_DIR = os.path.join(HERE, "output")
SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")
DEFAULTS_PATH = os.path.join(TEMPLATE_DIR, "defaults.json")

# 内置的两类 Word 明细表。项目类型保存在工作台设置中，导出时按类型选取。
DOCX_TEMPLATE_FILES = {
    "research": "武汉大学学生劳务费发放明细表（科研经费）.docx",
    "non_research": "武汉大学学生劳务费发放明细表（非科研经费）.docx",
}

SAFE_NAME = re.compile(r'[\\/:*?"<>|\r\n\t]+')

DEFAULT_SETTINGS = {
    "unitName": "",
    "projectCode": "",
    "period": "",
    "projectName": "",
    "note": "",
    "projectType": "research",
}

DEFAULT_BATCH = {"rate": "100", "hours": "10", "college": ""}

# 没填学院时兜底用的学院（换单位只改这一处）
DEFAULT_COLLEGE = "国家网络安全学院"

# 补贴费预设（元）。用户可在界面「补贴费预设」里随意改。
DEFAULT_FEE_PRESETS = ["500", "800", "1000", "1200", "1500", "1600", "2000"]


# 默认兼职时段的起点。首次打开、以及月份预设都从这里开始算。
# 想换起始月份只改这一处，例如 "2026年9月"。
DEFAULT_PERIOD_START = "2026年6月"

# 月份预设生成几个
PERIOD_PRESET_COUNT = 6


def _parse_period(text):
    """把 "2026年6月" 解析成 (年, 月)；解析不出来返回 None。"""
    m = re.match(r"^\s*(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*月?\s*$", str(text or ""))
    if not m:
        return None
    year, month = int(m.group(1)), int(m.group(2))
    if not (1 <= month <= 12) or not (1900 <= year <= 2999):
        return None
    return year, month


def current_period():
    """默认兼职时段 = DEFAULT_PERIOD_START（2026年6月）。

    故意做成固定的、可预测的：基准月份只在上面那一处定义，
    改一处就全改。想换月份的话，第 1 步顶部点一下即可，或者改这个常量。
    """
    start = _parse_period(DEFAULT_PERIOD_START)
    if start is None:
        today = datetime.date.today()
        return "%d年%d月" % (today.year, today.month)
    return "%d年%d月" % start


def period_presets(count=None):
    """默认月份预设：从 DEFAULT_PERIOD_START 起往后 count 个月。"""
    if count is None:
        count = PERIOD_PRESET_COUNT
    start = _parse_period(DEFAULT_PERIOD_START)
    if start is None:
        today = datetime.date.today()
        start = (today.year, today.month)

    year, month = start
    out = []
    for i in range(count):
        y = year + (month - 1 + i) // 12
        m = (month - 1 + i) % 12 + 1
        out.append("%d年%d月" % (y, m))
    return out


def _ensure_dirs():
    for path in (TEMPLATE_DIR, APP_DIR, OUTPUT_DIR):
        os.makedirs(path, exist_ok=True)


def normalize_presets(raw):
    """把前端传来的预设金额规整成字符串列表（去重、保序、丢弃非法值）。"""
    if not isinstance(raw, (list, tuple)):
        return list(DEFAULT_FEE_PRESETS)
    out, seen = [], set()
    for item in raw:
        if isinstance(item, bool) or item is None:
            continue
        text = str(item).strip()
        if not text:
            continue
        try:
            number = float(text.replace(",", ""))
        except ValueError:
            continue
        if number < 0:
            continue
        if number == int(number):
            text = str(int(number))
        else:
            text = ("%.2f" % number).rstrip("0").rstrip(".")
        if text in seen:
            continue
        seen.add(text)
        out.append(text)
    return out or list(DEFAULT_FEE_PRESETS)


def normalize_period_presets(raw):
    """兼职时段的预设月份：纯文本列表，去重保序，不合法就回退到默认。"""
    if not isinstance(raw, (list, tuple)):
        return []
    out, seen = [], set()
    for item in raw:
        if item is None:
            continue
        text = str(item).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        out.append(text)
    return out


def read_json(path, fallback=None):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return fallback


def write_json(path, data):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def safe_filename(name, fallback="未命名"):
    cleaned = SAFE_NAME.sub("_", str(name or "")).strip().strip(".")
    return cleaned or fallback


def normalize_project_type(value):
    """统一项目类型；兼容旧数据以及前端可能传来的中文值。"""
    text = str(value or "").strip().lower()
    if text in ("non_research", "nonresearch", "non-research", "非科研", "非科研项目"):
        return "non_research"
    return "research"


def slug(text):
    """把中文名转成安全的模板 id。"""
    cleaned = SAFE_NAME.sub("", str(text or "")).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:60] or "template"


# ---------------------------------------------------------------- 内置默认值
def load_builtin_defaults():
    data = read_json(DEFAULTS_PATH, None)
    fallback_batch = dict(DEFAULT_BATCH)
    fallback_batch["college"] = DEFAULT_COLLEGE
    if not isinstance(data, dict):
        return {"settings": dict(DEFAULT_SETTINGS),
                "batch": fallback_batch,
                "feePresets": list(DEFAULT_FEE_PRESETS),
                "periodStart": DEFAULT_PERIOD_START,
                "periodPresets": period_presets(),
                "students": []}
    out = {
        "settings": dict(DEFAULT_SETTINGS),
        "batch": fallback_batch,
        "feePresets": normalize_presets(data.get("feePresets")),
        "periodStart": DEFAULT_PERIOD_START,
        "periodPresets": (normalize_period_presets(data.get("periodPresets"))
                          or period_presets()),
        "students": [],
    }
    out["settings"].update(data.get("settings") or {})
    out["settings"]["projectType"] = normalize_project_type(
        out["settings"].get("projectType"))
    out["batch"].update(data.get("batch") or {})
    # 时段留空就补当前月份
    if not str(out["settings"].get("period") or "").strip():
        out["settings"]["period"] = current_period()
    # 学院兜底：默认值里没写就用 DEFAULT_COLLEGE，保证不会留空
    if not str(out["batch"].get("college") or "").strip():
        out["batch"]["college"] = DEFAULT_COLLEGE
    for index, student in enumerate(data.get("students") or []):
        out["students"].append({
            "id": student.get("id") or ("b%d" % (index + 1)),
            "studentId": str(student.get("studentId", "")),
            "name": str(student.get("name", "")),
            "college": str(student.get("college", "")) or DEFAULT_COLLEGE,
            "rate": student.get("rate", ""),
            "hours": student.get("hours", ""),
            "amount": student.get("amount", ""),
            "checked": bool(student.get("checked", True)),
        })
    return out


# ------------------------------------------------------------------ 用户模板
def template_file(template_id):
    return os.path.join(TEMPLATE_DIR, safe_filename(template_id) + ".json")


def list_templates():
    _ensure_dirs()
    result = []
    for name in sorted(os.listdir(TEMPLATE_DIR)):
        if not name.lower().endswith(".json"):
            continue
        if name in ("defaults.json", "builtin.json"):
            continue
        data = read_json(os.path.join(TEMPLATE_DIR, name), None)
        if not isinstance(data, dict):
            continue
        result.append({
            "id": os.path.splitext(name)[0],
            "name": data.get("name") or os.path.splitext(name)[0],
            "updatedAt": data.get("updatedAt", ""),
            "studentCount": len(data.get("students") or []),
            "unitName": (data.get("settings") or {}).get("unitName", ""),
            "period": (data.get("settings") or {}).get("period", ""),
        })
    result.sort(key=lambda item: item.get("updatedAt", ""), reverse=True)
    return result


def load_template(template_id):
    path = template_file(template_id)
    data = read_json(path, None)
    if not isinstance(data, dict):
        raise FileNotFoundError("模板不存在：%s" % template_id)

    payload = {
        "name": data.get("name") or template_id,
        "settings": dict(DEFAULT_SETTINGS),
        "batch": dict(DEFAULT_BATCH),
        "feePresets": normalize_presets(data.get("feePresets")),
        "periodPresets": normalize_period_presets(data.get("periodPresets")),
        "students": [],
    }
    payload["settings"].update(data.get("settings") or {})
    payload["batch"].update(data.get("batch") or {})
    # 模板默认不启用「记住学生」，只带默认值
    if data.get("includeStudents"):
        payload["students"] = data.get("students") or []
    return payload


def save_template(template_id, payload):
    _ensure_dirs()
    name = (payload.get("name") or template_id or "").strip()
    if not template_id:
        template_id = slug(name) or ("tpl_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S"))
    record = {
        "name": name or template_id,
        "updatedAt": datetime.datetime.now().isoformat(timespec="seconds"),
        "settings": payload.get("settings") or {},
        "batch": payload.get("batch") or {},
        "feePresets": normalize_presets(payload.get("feePresets")),
        "periodPresets": normalize_period_presets(payload.get("periodPresets")),
        "includeStudents": bool(payload.get("includeStudents")),
        "students": payload.get("students") or [],
    }
    write_json(template_file(template_id), record)
    return {"id": template_id, "name": record["name"]}


def delete_template(template_id):
    path = template_file(template_id)
    if os.path.isfile(path):
        os.remove(path)
        return True
    return False


# -------------------------------------------------------------- 工作台状态
def load_settings():
    _ensure_dirs()
    data = read_json(SETTINGS_PATH, None)
    if not isinstance(data, dict):
        seed = load_builtin_defaults()
        seed["step"] = 0          # 首次打开：按有没有名单自动决定停在第几步
        seed["savedAt"] = datetime.datetime.now().isoformat(timespec="seconds")
        write_json(SETTINGS_PATH, seed)
        return seed

    try:
        step = int(data.get("step") or 0)
    except (TypeError, ValueError):
        step = 0

    out = {
        "step": step if 1 <= step <= 3 else 0,
        "settings": dict(DEFAULT_SETTINGS),
        "batch": dict(DEFAULT_BATCH),
        "feePresets": normalize_presets(data.get("feePresets")),
        "periodStart": DEFAULT_PERIOD_START,
        "periodPresets": (normalize_period_presets(data.get("periodPresets"))
                          or period_presets()),
        "students": data.get("students") or [],
        "savedAt": data.get("savedAt", ""),
    }
    out["settings"].update(data.get("settings") or {})
    out["settings"]["projectType"] = normalize_project_type(
        out["settings"].get("projectType"))
    out["batch"].update(data.get("batch") or {})
    # 时段留空就补当前月份，别让用户看到一个空着或很旧的默认值
    if not str(out["settings"].get("period") or "").strip():
        out["settings"]["period"] = current_period()
    # 学院兜底
    if not str(out["batch"].get("college") or "").strip():
        out["batch"]["college"] = DEFAULT_COLLEGE
    return out


def save_settings(payload):
    _ensure_dirs()
    try:
        step = int(payload.get("step") or 0)
    except (TypeError, ValueError):
        step = 0
    record = {
        "step": step if 1 <= step <= 3 else 0,
        "settings": payload.get("settings") or {},
        "batch": payload.get("batch") or {},
        "feePresets": normalize_presets(payload.get("feePresets")),
        "periodPresets": normalize_period_presets(payload.get("periodPresets")),
        "students": payload.get("students") or [],
        "savedAt": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    record["settings"]["projectType"] = normalize_project_type(
        record["settings"].get("projectType"))
    write_json(SETTINGS_PATH, record)
    return record


def reset_settings():
    if os.path.isfile(SETTINGS_PATH):
        os.remove(SETTINGS_PATH)
    return load_settings()


# ------------------------------------------------------------------ 输出目录
def next_output_path(kind, settings=None, label=""):
    """生成 output/ 下的文件名，若已存在则自动加序号。"""
    _ensure_dirs()
    settings = settings or {}
    period = safe_filename(settings.get("period") or "", "").strip("_")
    stamp = datetime.datetime.now().strftime("%Y%m%d")
    parts = [p for p in (period, label, stamp) if p]
    base = "_".join(parts) or ("导出_" + stamp)
    if kind == "docx":
        base += "_劳务费发放明细表"
    else:
        base += "_助研津贴名单"

    candidate = os.path.join(OUTPUT_DIR, base + "." + kind)
    index = 2
    while os.path.exists(candidate):
        candidate = os.path.join(OUTPUT_DIR, "%s(%d).%s" % (base, index, kind))
        index += 1
    return candidate


def ensure_docx_template(uploaded_path, original_name, project_type=None):
    """把用户上传的 Word 模板存到 templates/，并按项目类型记录底板。"""
    _ensure_dirs()
    name = safe_filename(original_name or "明细表模板")
    if not name.lower().endswith(".docx"):
        name += ".docx"
    target = os.path.join(TEMPLATE_DIR, name)
    shutil.copyfile(uploaded_path, target)
    kind = normalize_project_type(project_type)
    meta_path = os.path.join(TEMPLATE_DIR, "docx_base.json")
    meta = read_json(meta_path, {})
    if not isinstance(meta, dict):
        meta = {}
    # 兼容旧版单一底板记录：旧的 file 仍作为科研模板使用。
    meta.setdefault("files", {})
    if meta.get("file") and "research" not in meta["files"]:
        meta["files"]["research"] = meta["file"]
    meta["files"][kind] = name
    meta["file"] = name
    meta["updatedAt"] = datetime.datetime.now().isoformat(timespec="seconds")
    write_json(meta_path, meta)
    return target


def current_docx_template(project_type=None):
    """返回指定项目类型的 Word 模板；没有自定义底板时使用内置模板。"""
    kind = normalize_project_type(project_type)
    meta = read_json(os.path.join(TEMPLATE_DIR, "docx_base.json"), None)
    if isinstance(meta, dict):
        filename = (meta.get("files") or {}).get(kind)
        if not filename and kind == "research":
            filename = meta.get("file")
        if filename:
            candidate = os.path.join(TEMPLATE_DIR, safe_filename(filename))
            if os.path.isfile(candidate):
                return candidate
    bundled = os.path.join(TEMPLATE_DIR, DOCX_TEMPLATE_FILES[kind])
    if os.path.isfile(bundled):
        return bundled
    # 兼容旧目录结构中的科研模板。
    legacy = os.path.join(HERE, "助研费用发放说明(1).docx")
    if kind == "research" and os.path.isfile(legacy):
        return legacy
    return None
