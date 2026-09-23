# -*- coding: utf-8 -*-
"""依赖引导：找得到就用，找不到（或坏了）就装，绝不让残缺的目录挡住可用的包。

为什么需要这个文件
------------------
`vendor/pylib` 里如果留下一个**空的、而且读不了**的同名目录（pip 装到一半失败、
权限被改坏时就会这样），Python 会把它当成"命名空间包"，于是

    import xlwt            -> 成功，但 xlwt.Workbook 不存在
    AttributeError: module 'xlwt' has no attribute 'Workbook'

也就是说"能 import"并不代表"能用"。这个模块会检查每个依赖是否真的可用，
把坏的目录从搜索路径里摘掉，必要时换一个干净的目录重装。

依赖目录的优先级：
    BAOXIAO_VENDOR_DIR 指定的目录  ->  vendor/pylib  ->  vendor/pylib2
重装优先往 vendor/pylib2 装，因为 vendor/pylib 往往正是因为坏掉才需要重装。
"""
import importlib.util
import os
import sys
import sysconfig

HERE = os.path.dirname(os.path.abspath(__file__))

# (pip 包名, import 名)
DEPENDENCIES = [
    ("python-docx", "docx"),
    ("xlrd", "xlrd"),
    ("xlwt", "xlwt"),
    ("openpyxl", "openpyxl"),
    ("reportlab", "reportlab"),
]

# 每个包都要求有这些属性，专门用来识别"空壳"包（namespace package）
REQUIRED_ATTRS = {
    "xlwt": ("Workbook",),
    "xlrd": ("open_workbook",),
    "docx": ("Document",),
    "openpyxl": ("Workbook",),
    "reportlab": ("__version__",),
}


def vendor_candidates():
    """按优先级列出依赖目录。"""
    out = []
    override = os.environ.get("BAOXIAO_VENDOR_DIR")
    if override:
        out.append(os.path.abspath(override))
    for name in ("pylib", "pylib2"):
        out.append(os.path.join(HERE, "vendor", name))
    return out


def _usable(module, name):
    """这个模块真能用吗？（空壳的 namespace package 不算）"""
    if module is None:
        return False
    if getattr(module, "__file__", None):
        pass
    elif not getattr(module, "__path__", None):
        return False           # 既没有文件也没有包路径，肯定不对
    else:
        # 有 __path__ 但没 __file__：可能是命名空间包，得有真东西才算数
        if not any(os.path.isfile(os.path.join(p, "__init__.py")) for p in module.__path__):
            return False
    return all(hasattr(module, attr) for attr in REQUIRED_ATTRS.get(name, ()))


def broken_module(name):
    """已导入的模块是不是坏的空壳。"""
    module = sys.modules.get(name)
    if module is None:
        return False
    return not _usable(module, name)


def probe(name):
    """在不污染 sys.modules 的前提下，看这个依赖能不能真的用。"""
    existing = sys.modules.get(name)
    if existing is not None and _usable(existing, name):
        return True
    if existing is not None and not _usable(existing, name):
        return False
    try:
        spec = importlib.util.find_spec(name)
    except (ImportError, ValueError, AttributeError):
        return False
    if spec is None:
        return False
    # 命名空间包（没有 loader/origin 且目录里没有 __init__.py）直接判为不可用
    if spec.origin is None and spec.submodule_search_locations:
        return any(os.path.isfile(os.path.join(p, "__init__.py"))
                   for p in spec.submodule_search_locations)
    try:
        module = importlib.import_module(name)
    except Exception:  # noqa: BLE001 - 导入期任何异常都算不可用
        return False
    return _usable(module, name)


def add_vendor_paths():
    """把依赖目录放到 sys.path 最前面，并摘掉会遮挡的坏目录。

    这里刻意**保留**坏目录在 sys.path 上，只是排到最后：
    - 后面的好目录只有排在它前面才找得到（比如 vendor/pylib2）
    - 万一两个目录都没有这个包，还能落到系统已装的版本
    """
    good, bad = [], []
    for path in vendor_candidates():
        (good if _healthy_dir(path) else bad).append(path)
    # 逐个 insert(0) 会把顺序倒过来，所以倒着插，保证 good[0] 排在最前面
    for path in reversed(good):
        if os.path.isdir(path):
            sys.path.insert(0, path)
    for path in bad:
        if os.path.isdir(path):
            sys.path.append(path)
    return good


def _healthy_dir(path):
    """这个依赖目录里有没有"读不了的"顶层子目录（会导致同一个包被遮挡）。"""
    if not os.path.isdir(path):
        return False
    try:
        entries = os.listdir(path)
    except OSError:
        return False
    for entry in entries:
        target = os.path.join(path, entry)
        if not os.path.isdir(target):
            continue
        try:
            os.listdir(target)
        except OSError:
            return False       # 读不了 -> 会变成空壳包，整个目录判为不健康
    return True


def missing():
    """返回还不能用的 (pip名, import名)。"""
    out = []
    for pip_name, module in DEPENDENCIES:
        if not probe(module):
            out.append((pip_name, module))
    return out


def forget(names):
    """把坏掉的模块从缓存里清掉，让下一次 import 重新找。"""
    for name in names:
        sys.modules.pop(name, None)


def install_target():
    """优先往 vendor/pylib2 装——vendor/pylib 往往正是坏掉的那个。"""
    override = os.environ.get("BAOXIAO_VENDOR_DIR")
    if override:
        return os.path.abspath(override)
    second = os.path.join(HERE, "vendor", "pylib2")
    first = os.path.join(HERE, "vendor", "pylib")
    if _healthy_dir(first) and not os.listdir(first):
        return first           # 又干净又是空的：直接用
    return second


def install_into_vendor(names, target=None):
    """把缺的依赖装进依赖目录；pip 不可用时再退回逐个下载 wheel。"""
    import subprocess

    target = target or install_target()
    os.makedirs(target, exist_ok=True)
    print("  缺少依赖：%s" % ", ".join(names))
    print("  正在安装到 %s ..." % os.path.relpath(target, HERE))
    print()

    env = dict(os.environ)
    # pip 要在临时目录里解包；受限账户下默认临时目录常常写不了
    tmp = os.path.join(target, ".pip-tmp")
    os.makedirs(tmp, exist_ok=True)
    env.update(TMP=tmp, TEMP=tmp, TMPDIR=tmp)

    commands = [
        [sys.executable, "-m", "pip", "install", "--target", target,
         "--disable-pip-version-check", "--no-warn-script-location"] + names,
        [sys.executable, "-m", "pip", "install", "--user",
         "--disable-pip-version-check"] + names,
    ]
    for command in commands:
        try:
            result = subprocess.run(command, env=env)
        except OSError as exc:
            print("  安装失败：%s" % exc)
            continue
        if result.returncode == 0:
            return target

    if download_wheels(names, target):
        return target
    return None


def download_wheels(names, target):
    """pip 用不了时的兜底：直接从 PyPI 下 wheel 解包（只用标准库）。"""
    import urllib.request
    import zipfile

    print("  pip 不可用，改为直接下载安装包 ...")
    ok = True
    for name in names:
        try:
            _clear_partial(target, name)
            url = _wheel_url(name)
            print("    %s" % os.path.basename(url.split("?", 1)[0]))
            with urllib.request.urlopen(url, timeout=120) as response:
                blob = response.read()
            local = os.path.join(target, ".download-%s.whl" % name)
            with open(local, "wb") as handle:
                handle.write(blob)
            try:
                with zipfile.ZipFile(local) as archive:
                    archive.extractall(target)
            finally:
                if os.path.exists(local):
                    os.remove(local)
        except Exception as exc:  # noqa: BLE001 - 兜底路径，失败就如实报告
            print("    失败：%s" % exc)
            ok = False
    return ok


def _clear_partial(target, name):
    """pip 装到一半留下的空壳目录会让解包失败，先清掉。

    同样地，如果有个**读不了**的同名目录（权限坏掉），这里也会尽力绕开：
    清不掉就如实报错，由调用方决定怎么办。
    """
    import shutil

    for entry in (name, _module_dir_name(name)):
        path = os.path.join(target, entry)
        if not os.path.isdir(path):
            continue
        try:
            if os.listdir(path):
                continue           # 有内容，不是空壳
        except OSError as exc:
            raise RuntimeError("目录不可读，无法安装：%s（%s）" % (path, exc))
        shutil.rmtree(path, ignore_errors=True)


def _module_dir_name(pip_name):
    return pip_name.replace("-", "_")


def _wheel_url(name):
    """从 PyPI 取这个包最新版的 wheel 地址。"""
    import json
    import urllib.request

    with urllib.request.urlopen("https://pypi.org/pypi/%s/json" % name, timeout=60) as response:
        data = json.load(response)
    for item in data.get("urls", []):
        if item.get("packagetype") == "bdist_wheel" and item["filename"].endswith(".whl"):
            return item["url"]
    raise RuntimeError("PyPI 上没有 %s 的 wheel" % name)


def ensure_dependencies(auto_install=True):
    """确保依赖可用。返回 (ok, 说明文字)。"""
    add_vendor_paths()
    todo = missing()
    if not todo:
        return True, ""

    names = [pip_name for pip_name, _ in todo]
    # 先清掉空壳，免得重新 import 时又拿到缓存里的坏模块
    forget(module for _, module in todo)
    add_vendor_paths()
    if not missing():
        return True, ""

    if not auto_install:
        return False, "缺少依赖：%s" % ", ".join(names)

    target = install_into_vendor(names)
    if target:
        forget(module for _, module in todo)
        if os.path.abspath(target) not in [os.path.abspath(p) for p in sys.path]:
            sys.path.insert(0, os.path.abspath(target))
        if not missing():
            return True, ""

    still = [pip_name for pip_name, _ in missing()]
    return False, ("依赖装不上：%s\n"
                   "请手动执行：%s -m pip install --target vendor/pylib2 %s"
                   % (", ".join(still), os.path.basename(sys.executable), " ".join(still)))
