#! /usr/local/bin/python3
"""Function for common indata checks.."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from typing import Callable, Optional
from excel_list_transform.commontypes import Data, Row


def check_indata_common(indata: Data[Row],
                        fix_indata_empty_rows: Callable[[Data[Row]], None]) \
                            -> None:
    """Check that the indata is well formed common part."""
    assert isinstance(indata, list), 'Internal error. ' + \
        f'Expected list of rows but got {type(indata).__name__}'
    if len(indata) < 1:
        msg = 'No rows in input data'
        print(msg, file=sys.stderr)
        sys.exit(1)
    fix_indata_empty_rows(indata)
    if len(indata) < 1:
        msg = 'No columns in input data'
        print(msg, file=sys.stderr)
        sys.exit(1)
    num_cols: Optional[int] = None
    for row in indata:
        rowlen: int = len(row)
        if num_cols is None:
            num_cols = rowlen
        else:
            assert num_cols is not None
            if num_cols != rowlen:
                msg = 'Number of columns different between lines. '
                msg += f'Found {num_cols} and {rowlen}. Aborting.'
                print(msg, file=sys.stderr)
                sys.exit(1)
