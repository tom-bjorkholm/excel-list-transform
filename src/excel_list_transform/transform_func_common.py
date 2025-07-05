#! /usr/local/bin/python3
"""Functions for doing transform of lists excel files."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from copy import deepcopy
from excel_list_transform.config_enums import SplitWhere
from excel_list_transform.config_excel_list_transform import Column, \
    Rule, RuleMerge, RuleRewrite, RuleSplit, SingleRuleSplit, \
    ConfigExcelListTransform
from excel_list_transform.rewrite_value import rewrite_value
from excel_list_transform.commontypes import  \
    NumRow, NameRow, NumRowSeq, NameRowMap, Data, Row, Value


def col_must_exist_num(col: int, row: NumRowSeq, param: str) -> None:
    """Stop with error report if column does not exist (num case)."""
    if col < 0 or col >= len(row):
        msg = f'{param}: column index {col} out of range [0, {len(row)-1}].'
        print(msg, file=sys.stderr)
        sys.exit(1)


def col_must_exist_name(col: str, row: NameRowMap, param: str) -> None:
    """Stop with error report if column does not exist (name case)."""
    if col not in row:
        msg: str = f'{param}: no column named "{col}" in data row.'
        print(msg, file=sys.stderr)
        sys.exit(1)


def col_must_exist(col: Column, row: NumRowSeq | NameRowMap,
                   param: str) -> None:
    """Stop with error report if column does not exist."""
    if isinstance(col, str):
        assert isinstance(col, str)
        assert isinstance(row, dict)
        col_must_exist_name(col=col, row=row, param=param)
    else:
        assert isinstance(col, int)
        assert isinstance(row, list)
        col_must_exist_num(col=col, row=row, param=param)


def cols_must_exist_lst(cols: list[Column], row: NumRowSeq | NameRowMap,
                        param: str, tinfo: Column) -> None:
    """Stop with error report all if columns does not exist."""
    for col in cols:
        assert isinstance(col, type(tinfo))
        col_must_exist(col=col, row=row, param=param)


def cols_must_exist_dict(rule:  Rule[Column] | RuleSplit[Column] |
                         RuleRewrite[Column],
                         row: NumRowSeq | NameRowMap,
                         param: str, tinfo: Column) -> None:
    """Stop with error report if columns does not exist."""
    cols = ConfigExcelListTransform.get_cols_single(rule=rule, tinfo=tinfo)
    cols_must_exist_lst(cols=cols, row=row, param=param, tinfo=tinfo)


def cols_must_exist_dictlst(rule: RuleMerge[Column],
                            row: NumRowSeq | NameRowMap,
                            param: str, tinfo: Column) -> None:
    """Stop with error report if columns does not exist."""
    cols = ConfigExcelListTransform.get_cols_multi(rule=rule, tinfo=tinfo)
    cols_must_exist_lst(cols=cols, row=row, param=param, tinfo=tinfo)


def store_col_split_num(row: NumRow, colref: int, val: list[str],
                        singlerule: SingleRuleSplit[int]) -> None:
    """Store result of column split for number refs."""
    if len(val) == 0:
        for _ in range(2):
            row.insert(colref, None)
        return
    if len(val) == 1:
        if singlerule['store_single'] != SplitWhere.RIGHTMOST:
            row.insert(colref, None)
            row.insert(colref, val[0])
        else:
            row.insert(colref, val[0])
            row.insert(colref, None)
        return
    for j in reversed(val):
        row.insert(colref, j)


def store_col_split_name(row: NameRow, colref: str, val: list[str],
                         singlerule: SingleRuleSplit[str]) -> None:
    """Store result of column split for number refs."""
    rightname = singlerule['right_name']
    assert isinstance(rightname, str)
    row[rightname] = None
    if len(val) == 0:
        row[colref] = None
    elif len(val) == 1:
        row[colref] = val[0]
    else:
        row[colref] = val[0]
        row[rightname] = val[1]


def pop_from_row(row: NumRow | NameRow, colref: Column) -> Value:
    """Keep mypy happy with row.pop."""
    if isinstance(colref, int):
        assert isinstance(colref, int)
        assert isinstance(row, list)
        return row.pop(colref)
    assert isinstance(colref, str)
    assert isinstance(row, dict)
    return row.pop(colref)


def split_columns(indata: Data[Row], cfg: ConfigExcelListTransform[Column],
                  tinfo: Column) -> Data[Row]:
    """Split columns in the in indata."""
    if len(cfg.s03_split_columns) == 0:
        return indata
    ret = deepcopy(indata)
    for row in ret:
        cols_must_exist_dict(rule=cfg.s03_split_columns, row=row,
                             param='s03_split_columns', tinfo=tinfo)
        for i in reversed(cfg.s03_split_columns):
            colref = i['column']
            assert isinstance(colref, type(tinfo))
            sep = i['separator']
            assert isinstance(sep, str)
            inval = pop_from_row(row=row, colref=colref)
            val: list[str] = []
            if inval is not None:
                if not isinstance(inval, str):
                    msg = 's03_split_columns: '
                    msg += 'Can only split columns with string values.\n'
                    msg += f'Column "{colref}" has value of '
                    msg += f'type {type(inval).__name__}'
                    print(msg, file=sys.stderr)
                    sys.exit(1)
                assert isinstance(inval, str)
                if i['where'] == SplitWhere.RIGHTMOST:
                    val = inval.rsplit(sep=sep, maxsplit=1)
                else:
                    val = inval.split(sep=sep, maxsplit=1)
            if isinstance(tinfo, str):
                assert isinstance(row, dict)
                assert isinstance(colref, str)
                store_col_split_name(row=row, colref=colref, val=val,
                                     singlerule=i)
            else:
                assert isinstance(row, list)
                assert isinstance(colref, int)
                store_col_split_num(row=row, colref=colref, val=val,
                                    singlerule=i)
    return ret


def insert_into_row(row: NumRow | NameRow, colref: Column, val: Value) -> None:
    """Keep mypy happy insert into row."""
    if isinstance(colref, int):
        assert isinstance(colref, int)
        assert isinstance(row, list)
        row.insert(colref, val)
        return
    assert isinstance(colref, str)
    assert isinstance(row, dict)
    row[colref] = val


def merge_columns(indata: Data[Row], cfg: ConfigExcelListTransform[Column],
                  tinfo: Column) -> Data[Row]:
    """Merge columns in the list in indata with column number refs."""
    if len(cfg.s05_merge_columns) == 0:
        return indata
    ret = deepcopy(indata)
    for row in ret:
        cols_must_exist_dictlst(rule=cfg.s05_merge_columns, row=row,
                                param='s05_merge_columns', tinfo=tinfo)
        for i in reversed(cfg.s05_merge_columns):
            colspec = i['columns']
            assert isinstance(colspec, list)
            if len(colspec) < 2:
                continue
            colref = colspec[0]
            assert isinstance(colref, type(tinfo))
            vals: list[str] = []
            for j in reversed(colspec):
                assert isinstance(j, type(tinfo))
                popval = pop_from_row(row=row, colref=j)
                if popval is not None:
                    vals.insert(0, str(popval))
            if len(vals) > 0:
                sep = i['separator']
                assert isinstance(sep, str)
                insert_into_row(row=row, colref=colref,
                                val=sep.join(vals))
            else:
                insert_into_row(row=row, colref=colref, val=None)
    return ret


def row_element(row: Row, idx: Column, tinfo: Column) -> Value:
    """Keep mypy happy with accessing element in dict or list."""
    assert isinstance(idx, type(tinfo))
    if isinstance(idx, int):
        assert isinstance(idx, int)
        assert isinstance(row, list)
        return row[idx]
    assert isinstance(idx, str)
    assert isinstance(row, dict)
    return row[idx]


def set_row_element(row: Row, idx: Column, tinfo: Column,
                    val: Value) -> None:
    """Keep mypy happy with accessing element in dict or list."""
    assert isinstance(idx, type(tinfo))
    if isinstance(idx, int):
        assert isinstance(idx, int)
        assert isinstance(row, list)
        row[idx] = val
        return
    assert isinstance(idx, str)
    assert isinstance(row, dict)
    row[idx] = val


def rewrite_columns(indata: Data[Row], cfg: ConfigExcelListTransform[Column],
                    tinfo: Column) -> Data[Row]:
    """Rewrite columns in the list in indata with column number refs."""
    if len(cfg.s09_rewrite_columns) == 0:
        return indata
    ret = deepcopy(indata)
    for row in ret:
        cols_must_exist_dict(rule=cfg.s09_rewrite_columns, row=row,
                             param='s09_rewrite_columns', tinfo=tinfo)
        for i in cfg.s09_rewrite_columns:
            colref = i['column']
            assert isinstance(colref, type(tinfo))
            oldvalue = row_element(row=row, idx=colref, tinfo=tinfo)
            if oldvalue is None:
                continue
            assert oldvalue is not None
            newvalue = rewrite_value(value=str(oldvalue), spec=i, tinfo=tinfo)
            set_row_element(row=row, idx=colref, tinfo=tinfo, val=newvalue)
    return ret
