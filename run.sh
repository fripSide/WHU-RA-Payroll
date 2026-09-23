#!/bin/bash
# 启动：./run.sh
#
# 代码在 src/，数据在 data/，Word 底板在 templates/，界面在 web/。
# 依赖由 src/bootstrap.py 在启动时自动补齐（vendor/pylib、vendor/pylib2）。

cd "$(dirname "$0")"
exec python3 src/server.py "$@"
