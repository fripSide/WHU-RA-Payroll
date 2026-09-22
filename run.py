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


def self_check():
    """自检：不动任何数据，生成一份 Word 和 PDF，确认环境正常。

    用法： python run.py --selfcheck
    """
    import bootstrap

    ok, note = bootstrap.ensure_dependencies()
    if not ok:
        print("依赖不可用：%s" % note)
        return 1

    import docx_gen
    import pdf_gen

    out = os.path.join(HERE, "自检输出")
    os.makedirs(out, exist_ok=True)
    payload = {
        "settings": {"period": "2026年7月", "unitName": "自检单位",
                     "projectCode": "0000-000000", "projectName": "自检项目",
                     "workContent": "自检工作"},
        "students": [{"id": "t1", "studentId": "2024000000001", "name": "自检同学",
                      "college": "计算机学院", "rate": "100", "hours": "10",
                      "amount": "1000", "manual": True, "checked": True,
                      "identityName": "本科生"}],
        "projectName": "自检项目",
    }
    docx_path = docx_gen.generate_docx(payload, os.path.join(out, "自检_明细表.docx"))["path"]
    pdf_gen.generate_pdf(docx_path, os.path.join(out, "自检_明细表.pdf"))
    print("\n自检通过，环境没问题。已生成：")
    for name in sorted(os.listdir(out)):
        print("   %s" % os.path.join("自检输出", name))
    print("\n（自检不会改动你的任何数据。）")
    return 0


def main():
    _force_utf8_output()

    parser = argparse.ArgumentParser(
        description="学生劳务费发放助手（本地服务）", add_help=True)
    parser.add_argument("--port", type=int, default=None,
                        help="指定端口，默认从 8765 起自动寻找可用端口")
    parser.add_argument("--no-open", action="store_true",
                        help="不要自动打开浏览器")
    parser.add_argument("--selfcheck", action="store_true",
                        help="自检：生成一份样例 Word/PDF，确认环境正常（不改动数据）")
    args = parser.parse_args()

    if args.selfcheck:
        return self_check()

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
