#! /bin/zsh
#
# Copyright (c) 2024-2026 Tom Björkholm
# MIT License
#
set -v
cd test
../venv/bin/python -m test_excel_list_transform.measure_speed
