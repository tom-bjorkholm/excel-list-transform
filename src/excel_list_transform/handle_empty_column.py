#! /usr/local/bin/python3
"""Handle empty column in input."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


from excel_list_transform.commontypes import NumData, NumRow, NumDataSeq


def handle_empty_column_in_lst(input_data: NumData | NumDataSeq) -> NumData:
    """Handle empty column in list of lists of input data."""
    ret: NumData = []
    for row in input_data:
        nrow: NumRow = [i if i != '' else None for i in row]
        ret.append(nrow)
    return ret
