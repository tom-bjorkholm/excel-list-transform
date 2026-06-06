#! /usr/local/bin/python3
"""Common type definitions used for type hints."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


from typing import Optional, Sequence, Mapping, TypeVar
from datetime import datetime


# types used to describe input and output data
type Value = Optional[str | int | float | datetime]
type NumRow = list[Value]
type NumRowSeq = Sequence[Value]
type NameRow = dict[str, Value]
Row = TypeVar('Row', NumRow, NameRow)
type NameRowMap = Mapping[str, Value]
type NumData = list[NumRow]
type NumDataSeq = list[NumRowSeq]
type NameData = list[NameRow]
type NameDataMap = Sequence[NameRowMap]
type Data[Row] = list[Row]

# helper functions


def num_row_to_str_list(row: NumRowSeq) -> list[str]:
    """Convert NumRow to list of str."""
    ret: list[str] = []
    for i in row:
        if isinstance(i, str):
            ret.append(i)
        elif i is None:
            raise TypeError('Found None when expecting str.')
        else:
            ret.append(str(i))
    return ret


T = TypeVar('T')


def get_checked_type(value: object, istype: type[T]) -> T:
    """Get value narrowed to be of type istype."""
    assert value is not None
    assert isinstance(value, istype)
    return value
