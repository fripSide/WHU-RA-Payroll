#!/usr/bin/env bash
# 学生劳务费发放助手 —— macOS / Linux 启动脚本
# 用法：
#     ./run.sh                 自动找端口并打开浏览器
#     ./run.sh --no-open       不自动打开浏览器
#     ./run.sh --port 9000     指定端口
set -euo pipefail

cd "$(dirname "$0")"

# UTF-8 模式：中文文件名/输出在 Linux、macOS 上都不出错
export PYTHONUTF8=1

# 按优先级找一个可用的 Python 3
PY=""
for candidate in python3 python; do
  if command -v "$candidate" >/dev/null 2>&1; then
    if "$candidate" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 8) else 1)' 2>/dev/null; then
      PY="$candidate"
      break
    fi
  fi
done

if [ -z "$PY" ]; then
  echo
  echo "  [错误] 没有找到 Python 3.8 或更高版本。"
  echo
  echo "  macOS:   brew install python3"
  echo "  Ubuntu:  sudo apt install python3 python3-pip"
  echo "  CentOS:  sudo yum install python3 python3-pip"
  echo
  exit 1
fi

exec "$PY" run.py "$@"
