#! /usr/local/bin/python3
"""Functions for manipulating and rewriting a value."""

# Copyright (c) 2025 Tom Björkholm
# MIT License

import sys
from copy import deepcopy
from excel_list_transform.commontypes import NameData, NameRow
from excel_list_transform.config_excel_list_transform import RuleRowSplit


def get_nosep_pos(instr: str,
                  not_separators: list[str]) -> list[tuple[int, int]]:
    """Find begin and end position of all not separators in input."""
    nosep_pos: list[tuple[int, int]] = []
    start = 0
    found = True
    lenstr = len(instr)
    while found:
        found = False
        nsep_pos = (lenstr, 0)
        for nsep in not_separators:
            beg = instr.find(nsep, start)
            if beg >= 0:
                found = True
                end = beg + len(nsep)
                if beg < nsep_pos[0] or (beg == nsep_pos[0] and
                                         end > nsep_pos[1]):
                    nsep_pos = (beg, end)
        if not found:
            break
        nosep_pos.append(nsep_pos)
        start = nsep_pos[1]
    return nosep_pos


def in_nosep_pos(pos: int, nosep_pos: list[tuple[int, int]]) -> bool:
    """Is this pos inside any of the nosep_pos intervals."""
    for nsepbeg, nsepend in nosep_pos:
        if nsepbeg <= pos < nsepend:
            return True
    return False


def split_one_str(instr: str, separators: list[str],
                  not_separators: list[str]) -> list[str]:
    """Split one string based on separators and not separators."""
    nosep_pos = get_nosep_pos(instr=instr, not_separators=not_separators)
    start = 0
    ret: list[str] = []
    while start < len(instr):
        pos = len(instr)
        seplen = 0
        for sep in separators:
            tpos = instr.find(sep, start)
            if tpos < 0:
                continue
            if in_nosep_pos(pos=tpos, nosep_pos=nosep_pos):
                continue
            if tpos < pos:
                pos = tpos
                seplen = len(sep)
        if pos < len(instr):
            ret.append(instr[start:pos])
            start = pos + seplen
        else:
            ret.append(instr[start:])
            start = len(instr)
    return ret


def one_split_one_row_name(inrow: NameRow, column: str, separators: list[str],
                           not_separators: list[str]) -> NameData:
    """Handle one split row directive for one row."""
    instr = inrow[column]
    if not isinstance(instr, str):
        print(f'Trying to split rows based on column "{column}".\n' +
              f'But that column has value of type {type(instr).__name__}:' +
              str(instr) + '\nCan onbly split strings\n',
              file=sys.stderr)
        sys.exit(1)
    assert isinstance(instr, str)
    splitted: list[str] = split_one_str(instr=instr,
                                        separators=separators,
                                        not_separators=not_separators)
    if len(splitted) <= 1:
        return [inrow]
    result: NameData = []
    for part in splitted:
        row: NameRow = deepcopy(inrow)
        row[column] = part
        result.append(row)
    return result


def one_split_name(indata: NameData, column: str, separators: list[str],
                   not_separators: list[str]) -> NameData:
    """Handle one of the split row directives."""
    if column not in indata[0]:
        print(f'Trying to split lines based on column "{column},"' +
              ' but no such column in data.\n',
              file=sys.stderr)
        sys.exit(1)
    result: NameData = []
    for row in indata:
        result += one_split_one_row_name(inrow=row, column=column,
                                         separators=separators,
                                         not_separators=not_separators)
    return result


def split_rows_name(indata: NameData,
                    directives: RuleRowSplit[str]) -> NameData:
    """Split rows according to configuration."""
    result: NameData = indata
    data: NameData = indata
    for directive in directives:
        col = directive['column']
        assert isinstance(col, str)
        sep = directive['separators']
        assert isinstance(sep, list)
        notsep = directive['not_separators']
        assert isinstance(notsep, list)
        result = one_split_name(indata=data, column=col, separators=sep,
                                not_separators=notsep)
        data = result
    return result
