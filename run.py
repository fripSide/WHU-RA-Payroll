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
MIN_PYTHON = (3, 8)


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
    """返回还没装上的 (pip名, import名)。判断逻辑在 bootstrap 里（含空壳包识别）。"""
    import bootstrap

    return bootstrap.missing()


def preflight():
    """检查 Python 版本与依赖；必要时自动安装。返回 True 表示可以启动。

    依赖是否可用交给 bootstrap 判断：它会识别"能 import 但其实是空壳"的情况
    （vendor/pylib 里残留的坏目录会这样），把坏目录摘掉再重装到干净目录。
    """
    if sys.version_info < MIN_PYTHON:
        banner([
            "需要 Python %d.%d 或更高版本" % MIN_PYTHON,
            "当前版本：%s" % sys.version.split()[0],
            "",
            "请升级 Python 后重试： https://www.python.org/downloads/",
        ])
        return False

    import bootstrap

    ok, note = bootstrap.ensure_dependencies()
    if ok:
        return True

    manual = [n for n, _ in bootstrap.missing()]
    banner([
        "依赖不可用：",
        "",
        "    %s" % note.replace("\n", "\n    "),
        "",
        "若提示权限不足，可改为：",
        "    %s -m pip install --user %s"
        % (os.path.basename(sys.executable), " ".join(manual)),
        "",
        "装好后重新运行即可。",
    ])
    return False


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
