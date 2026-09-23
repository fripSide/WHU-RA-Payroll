# -*- coding: utf-8 -*-
"""路径与 Word 模板选取。

目录约定（本文件在 src/ 下，所以路径都相对**项目根**，即上一级）：

- templates/*.docx        内置的两类明细表底板（科研 / 非科研）
- templates/docx_base.json 记录用户自定义底板（可选）
- data/                   projects.json、people.json 等本地数据
"""
import os
import re
import json

# 本文件在 src/ 下，项目根是它的上一级
SRC_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SRC_DIR)

TEMPLATE_DIR = os.path.join(ROOT, "templates")
DATA_DIR = os.path.abspath(os.environ.get("BAOXIAO_DATA_DIR") or os.path.join(ROOT, "data"))
OUTPUT_DIR = os.path.abspath(os.environ.get("BAOXIAO_OUTPUT_DIR") or os.path.expanduser("~/Downloads"))

# 兼容旧名字（外部脚本可能还在用 store.APP_DIR）
APP_DIR = DATA_DIR

# 内置的两类 Word 明细表。导出时按项目类型选取。
DOCX_TEMPLATE_FILES = {
    "research": "武汉大学学生劳务费发放明细表（科研经费）.docx",
    "non_research": "武汉大学学生劳务费发放明细表（非科研经费）.docx",
}

SAFE_NAME = re.compile(r'[\\/:*?"<>|\r\n\t]+')


def _ensure_dirs():
    for path in (TEMPLATE_DIR, APP_DIR, OUTPUT_DIR):
        os.makedirs(path, exist_ok=True)


def read_json(path, fallback=None):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return fallback


def safe_filename(name, fallback="未命名"):
    cleaned = SAFE_NAME.sub("_", str(name or "")).strip().strip(".")
    return cleaned or fallback


def normalize_project_type(value):
    """统一项目类型；兼容旧数据以及前端可能传来的中文值。"""
    text = str(value or "").strip().lower()
    if text in ("non_research", "nonresearch", "non-research", "非科研", "非科研项目"):
        return "non_research"
    return "research"


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
    return None
