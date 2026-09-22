#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""学生劳务费发放助手 —— 统一启动入口（Windows / macOS / Linux 通用）。

用法：
    python3 run.py            # 自动找端口并打开浏览器
    python3 run.py --no-open  # 不自动打开浏览器
    python3 run.py --port 9000
    ./run.sh                  # macOS / Linux 便捷脚本
    run.cmd                   # Windows 双击
"""
import os
import sys
import argparse

HERE = os.path.dirname(os.path.abspath(__file__))
VENDOR = os.path.abspath(os.environ.get("BAOXIAO_VENDOR_DIR")
                         or os.path.join(HERE, "vendor", "pylib"))

MIN_PYTHON = (3, 8)

# 依赖名 -> import 名 -> pip 包名
DEPENDENCIES = [
    ("python-docx", "docx"),
    ("xlrd", "xlrd"),
    ("xlwt", "xlwt"),
    ("openpyxl", "openpyxl"),
    ("reportlab", "reportlab"),
]


def _force_utf8_output():
    """保证中文不乱码，并且输出实时可见（重定向到文件时也不会卡在缓冲区）。

    配套 run.cmd / run.sh 会设置 PYTHONUTF8=1（Python 的 UTF-8 模式），
    这样中文文件名和输出在三个平台上都不会出问题。
    """
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace",
                               line_buffering=True)
        except (AttributeError, ValueError, TypeError):
            pass


def banner(lines):
    width = 60
    print("=" * width)
    for line in lines:
        print("  " + line)
    print("=" * width)
    sys.stdout.flush()


def missing_dependencies():
    """返回还没装上的 (pip名, import名)。"""
    missing = []
    for pip_name, module in DEPENDENCIES:
        try:
            __import__(module)
        except ImportError:
            missing.append((pip_name, module))
    return missing


def install_into_vendor(missing):
    """把缺的依赖装到 vendor/pylib（可随目录一起拷贝，不污染系统环境）。"""
    import subprocess

    os.makedirs(VENDOR, exist_ok=True)
    names = [pip_name for pip_name, _ in missing]
    print("  缺少依赖：%s" % ", ".join(names))
    print("  正在安装到 vendor/pylib ...")
    print()

    commands = [
        [sys.executable, "-m", "pip", "install", "--target", VENDOR,
         "--disable-pip-version-check", "--no-warn-script-location"] + names,
        [sys.executable, "-m", "pip", "install", "--user",
         "--disable-pip-version-check"] + names,
    ]
    for command in commands:
        try:
            result = subprocess.run(command)
        except OSError as exc:
            print("  安装失败：%s" % exc)
            continue
        if result.returncode == 0:
            print()
            print("  安装完成。")
            print()
            return True
        print()

    return False


def preflight():
    """检查 Python 版本与依赖；必要时自动安装。返回 True 表示可以启动。"""
    if sys.version_info < MIN_PYTHON:
        banner([
            "需要 Python %d.%d 或更高版本" % MIN_PYTHON,
            "当前版本：%s" % sys.version.split()[0],
            "",
            "请升级 Python 后重试： https://www.python.org/downloads/",
        ])
        return False

    # 优先使用随项目附带的依赖目录
    if os.path.isdir(VENDOR) and VENDOR not in sys.path:
        sys.path.insert(0, VENDOR)

    missing = missing_dependencies()
    if missing:
        if not install_into_vendor(missing):
            if VENDOR not in sys.path:
                sys.path.insert(0, VENDOR)
            still = missing_dependencies()
            if still:
                banner([
                    "依赖安装失败，请手动执行：",
                    "",
                    "    %s -m pip install --target vendor/pylib %s"
                    % (os.path.basename(sys.executable),
                       " ".join(n for n, _ in still)),
                    "",
                    "若提示权限不足，可改为：",
                    "    %s -m pip install --user %s"
                    % (os.path.basename(sys.executable),
                       " ".join(n for n, _ in still)),
                ])
                return False

    # 安装后重新加载路径
    if VENDOR not in sys.path:
        sys.path.insert(0, VENDOR)
    return True


def main():
    _force_utf8_output()

    parser = argparse.ArgumentParser(
        description="学生劳务费发放助手（本地服务）", add_help=True)
    parser.add_argument("--port", type=int, default=None,
                        help="指定端口，默认从 8765 起自动寻找可用端口")
    parser.add_argument("--no-open", action="store_true",
                        help="不要自动打开浏览器")
    args = parser.parse_args()

    if not preflight():
        return 1

    # 依赖就绪后再导入，避免在缺少依赖时抛出难懂的栈
    import server

    if args.no_open:
        os.environ["BAOXIAO_NO_BROWSER"] = "1"
    if args.port:
        os.environ["BAOXIAO_PORT"] = str(args.port)

    return server.main()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n已停止。")
        raise SystemExit(0)
