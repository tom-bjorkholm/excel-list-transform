#! /usr/local/bin/python3
"""Convert data between numbered columns and named columns."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from excel_list_transform.commontypes import NameData, NumData, \
    NumDataSeq, NumRow, NameDataMap, \
    num_row_to_str_list, str_list_to_num_row


def named_cols_from_num_cols(data: NumData | NumDataSeq,
                             filename: str) -> NameData:
    """Convert data with numbered columns to data with named columns.

    If rows have too few columns, pad as needed with None.
    """
    filemsg = f' in file {filename}.'
    for i, name in enumerate(data[0]):
        if name is None:
            print(f'Cannot handle input column {i} without name' + filemsg,
                  file=sys.stderr)
            sys.exit(1)
    names: list[str] = num_row_to_str_list(data[0])
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


def num_cols_from_named_cols(data: NameData | NameDataMap,
                             column_order: list[str]) -> NumData:
    """Convert data with named columns to data with numbered columns.

    Silently ignore data with keys not in column_order.
    If key is missing flag it as error.
    """
    ret: NumData = []
    ret.append(str_list_to_num_row(column_order))
    for rownum, row in enumerate(data):
        outrow: NumRow = []
        for key in column_order:
            if key in row:
                outrow.append(row[key])
            else:
                msg = f'Data row {rownum} is missing data for column {key}'
                print(msg, file=sys.stderr)
                sys.exit(1)
        ret.append(outrow)
    return ret
