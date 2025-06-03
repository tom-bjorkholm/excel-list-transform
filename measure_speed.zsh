#! /bin/zsh
#
# Copyright (c) 2024-2025 Tom Björkholm
# MIT License
#
set -v
PYTHON=`./bestInstalledPython.zsh`
./doBuild.zsh
. ./venv/bin/activate
${PYTHON} ./test/test_excel_list_transform/measure_speed.py
