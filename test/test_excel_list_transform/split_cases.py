#! /usr/local/bin/python3
"""Shared test cases for empty parts in row splits (name and num)."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from typing import NamedTuple, Optional
from excel_list_transform.config_enums import EmptyParts


class RowModeCase(NamedTuple):
    """Value split at ',' with an empty parts mode, and resulting parts."""

    value: Optional[str]
    mode: EmptyParts
    parts: list[Optional[str]]


ROW_MODE_CASES: list[RowModeCase] = [
    RowModeCase('a,,b,', EmptyParts.KEEP_ALL, ['a', '', 'b', '']),
    RowModeCase('a,,b,', EmptyParts.DROP_ALL, ['a', 'b']),
    RowModeCase('a,,b,', EmptyParts.DROP_TRAILING, ['a', '', 'b']),
    RowModeCase('a,b,,', EmptyParts.DROP_TRAILING, ['a', 'b']),
    RowModeCase(',a', EmptyParts.DROP_TRAILING, ['', 'a']),
    RowModeCase(',a', EmptyParts.DROP_ALL, ['a']),
    RowModeCase('a, ,b', EmptyParts.DROP_ALL, ['a', 'b']),
    RowModeCase('a, b', EmptyParts.DROP_ALL, ['a', ' b']),
    RowModeCase('a,', EmptyParts.DROP_TRAILING, ['a']),
    RowModeCase('a,', EmptyParts.KEEP_ALL, ['a', '']),
    RowModeCase(',', EmptyParts.KEEP_ALL, ['', '']),
    RowModeCase(',', EmptyParts.DROP_ALL, ['']),
    RowModeCase(' , ', EmptyParts.DROP_TRAILING, ['']),
    RowModeCase('', EmptyParts.KEEP_ALL, ['']),
    RowModeCase('', EmptyParts.DROP_ALL, ['']),
    RowModeCase('   ', EmptyParts.DROP_ALL, ['']),
    RowModeCase('   ', EmptyParts.KEEP_ALL, ['   ']),
    RowModeCase('abc', EmptyParts.DROP_TRAILING, ['abc']),
    RowModeCase(None, EmptyParts.KEEP_ALL, [None]),
    RowModeCase(None, EmptyParts.DROP_ALL, [None]),
    RowModeCase(None, EmptyParts.DROP_TRAILING, [None])]
"""Cases for splitting one cell at the separator ','."""
