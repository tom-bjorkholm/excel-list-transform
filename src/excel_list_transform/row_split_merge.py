#! /usr/local/bin/python3
"""Functions for manipulating and rewriting a value."""

# Copyright (c) 2025 Tom Björkholm
# MIT License

import sys
from typing import overload, cast
from copy import deepcopy
from excel_list_transform.commontypes import Value, \
    Row, NumRow, NameRow, Data
from excel_list_transform.config_excel_list_transform import RuleRowSplit, \
    SingleRuleMerge, RuleMerge, Column
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum


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


@overload
def one_split_one_row(inrow: NumRow, column: int,
                      separators: list[str],
                      not_separators: list[str]) -> Data[NumRow]: ...


@overload
def one_split_one_row(inrow: NameRow, column: str,
                      separators: list[str],
                      not_separators: list[str]) -> Data[NameRow]: ...


def one_split_one_row(inrow: Row, column: Column,
                      separators: list[str],
                      not_separators: list[str]) -> Data[Row]:
    """Handle one split row directive for one row."""
    assert (isinstance(inrow, list) and isinstance(column, int)) or \
        (isinstance(inrow, dict) and isinstance(column, str))
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
    result: Data[Row] = []
    for part in splitted:
        row: Row = deepcopy(inrow)
        row[column] = part
        result.append(row)
    return result


@overload
def column_in_row(col: int, row: NumRow) -> bool: ...


@overload
def column_in_row(col: str, row: NameRow) -> bool: ...


def column_in_row(col: Column, row: Row) -> bool:
    """Test if column exists in row."""
    if isinstance(row, dict):
        assert isinstance(row, dict)
        assert isinstance(col, str)
        return col in row
    assert isinstance(row, list)
    assert isinstance(col, int)
    if col < 0:
        return False
    return col < len(row)


@overload
def one_split(indata: Data[NameRow], column: str, separators: list[str],
              not_separators: list[str]) -> Data[NameRow]: ...


@overload
def one_split(indata: Data[NumRow], column: int, separators: list[str],
              not_separators: list[str]) -> Data[NumRow]: ...


def one_split(indata: Data[Row], column: Column, separators: list[str],
              not_separators: list[str]) -> Data[Row]:
    """Handle one of the split row directives."""
    if not indata:
        return indata
    assert (isinstance(indata[0], dict) and isinstance(column, str)) or \
        (isinstance(indata[0], list) and isinstance(column, int))
    if not column_in_row(column, indata[0]):
        print(f'Trying to split lines based on column "{column}",' +
              ' but no such column in data.\n',
              file=sys.stderr)
        sys.exit(1)
    result: Data[Row] = []
    for row in indata:
        result += one_split_one_row(inrow=row, column=column,
                                    separators=separators,
                                    not_separators=not_separators)
    return result


@overload
def split_rows(indata: Data[NameRow],
               directives: RuleRowSplit[str]) -> Data[NameRow]: ...


@overload
def split_rows(indata: Data[NumRow],
               directives: RuleRowSplit[int]) -> Data[NumRow]: ...


def split_rows(indata: Data[Row],
               directives: RuleRowSplit[Column]) -> Data[Row]:
    """Split rows according to configuration."""
    if not indata:
        return indata
    result: Data[Row] = indata
    data: Data[Row] = indata
    for directive in directives:
        col = directive['column']
        assert (isinstance(indata[0], list) and isinstance(col, int)) or\
            (isinstance(indata[0], dict) and isinstance(col, str))
        sep = directive['separators']
        assert isinstance(sep, list)
        notsep = directive['not_separators']
        assert isinstance(notsep, list)
        result = one_split(indata=data, column=col, separators=sep,
                           not_separators=notsep)
        data = result
    return result


@overload
def split_rows_cfg(indata: Data[NameRow],
                   cfg: ConfigXlsListTransfName,
                   tinfo: str) -> Data[NameRow]: ...


@overload
def split_rows_cfg(indata: Data[NumRow],
                   cfg: ConfigXlsListTransfNum,
                   tinfo: int) -> Data[NumRow]: ...


def split_rows_cfg(indata: Data[Row],
                   cfg: ConfigXlsListTransfName | ConfigXlsListTransfNum,
                   tinfo: Column) -> Data[Row]:
    """Split rows according to configuration."""
    if not indata:
        return indata
    if not cfg.s01_split_rows:
        return indata
    assert (isinstance(indata[0], dict) and isinstance(tinfo, str)) or \
           (isinstance(indata[0], list) and isinstance(tinfo, int))
    assert (isinstance(tinfo, str) and
            isinstance(cfg.s01_split_rows[0]['column'], str)) or \
           (isinstance(tinfo, int) and
            isinstance(cfg.s01_split_rows[0]['column'], int))
    direc: RuleRowSplit[Column] = \
        cast(RuleRowSplit[Column], cfg.s01_split_rows)
    return split_rows(indata=indata, directives=direc)


def merge_strings(to_merge: list[Value], sep: str) -> Value:
    """Merge list of strings to one cell value.

    If two or more values are on several position it will be
    merged only once. If two positions have different values
    the values will be converted to str and concatenated with
    the separator sep.
    @param to_merge A list of the cell values for one column for the rows
                    that are subject to the merge.
    @param sep      The separator between values when merged.
    """
    uniq: list[Value] = []
    for i in to_merge:
        if i not in uniq:
            uniq.append(deepcopy(i))
    if not uniq:
        return None
    if len(uniq) == 1:
        return uniq[0]
    ret = str(uniq[0])
    for i in uniq[1:]:
        if ret:
            ret += sep
        ret += str(i)
    return ret


def columns_in_row(row: Row, tinfo: Column) -> list[Column]:
    """Get list of columns in row."""
    if isinstance(row, dict):
        assert isinstance(row, dict)
        assert isinstance(tinfo, str)
        return list(row.keys())
    assert isinstance(row, list)
    assert isinstance(tinfo, int)
    return list(range(0, len(row)))


@overload
def merge_identified_rows(rows: Data[NameRow], row_numbers: list[list[int]],
                          separator: str, tinfo: str) -> Data[NameRow]: ...


@overload
def merge_identified_rows(rows: Data[NumRow], row_numbers: list[list[int]],
                          separator: str, tinfo: int) -> Data[NumRow]: ...


def merge_identified_rows(rows: Data[Row], row_numbers: list[list[int]],
                          separator: str, tinfo: Column) -> Data[Row]:
    """Merge rows identified by row numbers."""
    if len(rows) <= 1:
        return rows
    if len(row_numbers) < 1:
        return rows
    assert (isinstance(tinfo, int) and isinstance(rows[0], list)) or \
        (isinstance(tinfo, str) and isinstance(rows[0], dict))
    for colname in columns_in_row(row=rows[0], tinfo=tinfo):
        for merge_set in row_numbers:
            if len(merge_set) <= 1:
                continue
            to_merge: list[Value] = []
            for rownum in merge_set:
                to_merge.append(deepcopy(rows[rownum][colname]))
            rows[merge_set[0]][colname] = merge_strings(to_merge=to_merge,
                                                        sep=separator)
    rows_to_del: list[int] = []
    for merge_set in row_numbers:
        rows_to_del += merge_set[1:]
    for rownum in sorted(rows_to_del, reverse=True):
        del rows[rownum]
    return rows


@overload
def identify_rows_to_merge(rows: Data[NameRow], columns_to_cmp: list[str],
                           tinfo: str) -> list[list[int]]: ...


@overload
def identify_rows_to_merge(rows: Data[NumRow], columns_to_cmp: list[int],
                           tinfo: int) -> list[list[int]]: ...


def identify_rows_to_merge(rows: Data[Row], columns_to_cmp: list[Column],
                           tinfo: Column) -> list[list[int]]:
    """Identify rows to merge.

    Find rows that have identical values in the columns to compare.
    @param columns_to_compare  List of column names for columns to
                               compare values in..
    @param rows  The data rows to look for values in.
    @return List of lists of columns with identical values in
            columns being compared.
    """
    numrows = len(rows)
    ret: list[list[int]] = []
    if numrows < 2:
        return ret
    assert (isinstance(tinfo, int) and isinstance(rows[0], list)) or \
        (isinstance(tinfo, str) and isinstance(rows[0], dict))
    used_in_merge: list[int] = []
    for startnum, startrow in enumerate(rows):
        if startnum in used_in_merge:
            continue
        candidates = [startnum]
        for othernum in range(startnum+1, numrows):
            if othernum in used_in_merge:
                continue
            otherrow = rows[othernum]
            differs = False
            for col in columns_to_cmp:
                if startrow[col] != otherrow[col]:
                    differs = True
            if not differs:
                candidates.append(othernum)
        if len(candidates) >= 2:
            ret.append(candidates)
            used_in_merge += candidates
    return ret


@overload
def one_merge_rows(indata: Data[NameRow],
                   columns_to_cmp: list[str],
                   separator: str, tinfo: str) -> Data[NameRow]: ...


@overload
def one_merge_rows(indata: Data[NumRow],
                   columns_to_cmp: list[int],
                   separator: str, tinfo: int) -> Data[NumRow]: ...


def one_merge_rows(indata: Data[Row],
                   columns_to_cmp: list[Column],
                   separator: str, tinfo: Column) -> Data[Row]:
    """Merge rows based on content of single rule."""
    if len(indata) < 2:
        return indata
    assert (isinstance(tinfo, int) and isinstance(indata[0], list)) or \
        (isinstance(tinfo, str) and isinstance(indata[0], dict))
    rows_to_merge = identify_rows_to_merge(rows=indata,
                                           columns_to_cmp=columns_to_cmp,
                                           tinfo=tinfo)
    data = deepcopy(indata)
    data = merge_identified_rows(rows=data, row_numbers=rows_to_merge,
                                 separator=separator, tinfo=tinfo)
    return data


def one_rule_merge_rows(indata: Data[Row],
                        rule: SingleRuleMerge[Column],
                        tinfo: Column) -> Data[Row]:
    """Merge rows based on a single rule."""
    if len(indata) < 2:
        return indata
    assert (isinstance(tinfo, int) and isinstance(indata[0], list)) or \
        (isinstance(tinfo, str) and isinstance(indata[0], dict))
    assert 'columns' in rule
    columns = rule['columns']
    assert isinstance(columns, list)
    assert 'separator' in rule
    separator = rule['separator']
    assert isinstance(separator, str)
    return one_merge_rows(indata=indata,
                          columns_to_cmp=columns,
                          separator=separator, tinfo=tinfo)


def merge_rows(indata: Data[Row], rules: RuleMerge[Column],
               tinfo: Column) -> Data[Row]:
    """Merge rows based on rules."""
    if len(indata) < 2:
        return indata
    assert (isinstance(tinfo, int) and isinstance(indata[0], list)) or \
        (isinstance(tinfo, str) and isinstance(indata[0], dict))
    if len(indata) < 2:
        return indata
    assert (isinstance(tinfo, int) and isinstance(indata[0], list)) or \
        (isinstance(tinfo, str) and isinstance(indata[0], dict))
    data = deepcopy(indata)
    for rule in rules:
        data = one_rule_merge_rows(indata=data, rule=rule, tinfo=tinfo)
    return data


def merge_rows_cfg(indata: Data[Row],
                   cfg: ConfigXlsListTransfName | ConfigXlsListTransfNum,
                   tinfo: Column) -> Data[Row]:
    """Merge rows based on configuration."""
    assert (isinstance(tinfo, int) and
            isinstance(cfg, ConfigXlsListTransfNum)) or \
        (isinstance(tinfo, str) and isinstance(cfg, ConfigXlsListTransfName))
    if len(indata) < 2:
        return indata
    assert (isinstance(tinfo, int) and isinstance(indata[0], list)) or \
        (isinstance(tinfo, str) and isinstance(indata[0], dict))
    assert (isinstance(tinfo, int) and isinstance(indata[0], list)) or \
        (isinstance(tinfo, str) and isinstance(indata[0], dict))
    return merge_rows(indata=indata, rules=cfg.s02_merge_rows, tinfo=tinfo)
