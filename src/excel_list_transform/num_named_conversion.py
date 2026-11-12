#! /usr/local/bin/python3
"""Convert data between numbered columns and named columns."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


import sys
from excel_list_transform.commontypes import NameData, NumData, \
    NumDataSeq, NumRow, num_row_to_str_list


def _names_must_be_unique(names: list[str], filemsg: str) -> None:
    """Stop with error report if a column name occurs more than once."""
    seen: set[str] = set()
    for name in names:
        if name in seen:
            print(f'Duplicate input column name "{name}"' + filemsg,
                  file=sys.stderr)
            sys.exit(1)
        seen.add(name)


def named_cols_from_num_cols(data: NumData | NumDataSeq,
                             filename: str) -> NameData:
    """Convert data with numbered columns to data with named columns.

    If rows have too few columns, pad as needed with None.
    Column names (title row) must be present and unique.
    """
    filemsg = f' in file {filename}.'
    for i, name in enumerate(data[0]):
        if name is None:
            print(f'Cannot handle input column {i} without name' + filemsg,
                  file=sys.stderr)
            sys.exit(1)
    names: list[str] = num_row_to_str_list(data[0])
    _names_must_be_unique(names=names, filemsg=filemsg)
    num_cols = len(names)
    ret: NameData = []
    for row_in in data[1:]:
        nrow: NumRow = list(row_in)
        if len(nrow) > num_cols:
            print('Data row(s) have more columns than title row' + filemsg,
                  file=sys.stderr)
            sys.exit(1)
        while len(nrow) < num_cols:
            nrow.append(None)
        ret.append(dict(zip(names, nrow)))
    return ret
