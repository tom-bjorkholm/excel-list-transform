#! /usr/local/bin/python3
"""Test the excel_list_transform common functions functionality."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from typing import cast
from copy import deepcopy
import pytest
from pytest import CaptureFixture
from excel_list_transform.config_excel_list_transform import Rule, RuleMerge
from excel_list_transform.commontypes import NameRow, NumRow, Value
from excel_list_transform.transform_func_common import col_must_exist, \
    cols_must_exist_lst, cols_must_exist_dict, cols_must_exist_multi, \
    pop_from_row, insert_into_row


type TestColumn = int | str
type TestRow = NumRow | NameRow
type TestRule = Rule[int] | Rule[str]
type TestRuleMulti = RuleMerge[int] | RuleMerge[str]


def _col_must_exist(col: TestColumn, row: TestRow, param: str) -> None:
    """Call col_must_exist with narrowed matching column and row types."""
    if isinstance(col, int):
        assert isinstance(row, list)
        col_must_exist(col=col, row=row, param=param)
        return
    assert isinstance(row, dict)
    col_must_exist(col=col, row=row, param=param)


def _cols_must_exist_lst(cols: list[int] | list[str], row: TestRow, param: str,
                         tinfo: TestColumn) -> None:
    """Call cols_must_exist_lst with a narrowed column-list type."""
    if isinstance(tinfo, int):
        assert isinstance(row, list)
        cols_must_exist_lst(cols=cast(list[int], cols), row=row, param=param,
                            tinfo=tinfo)
        return
    assert isinstance(row, dict)
    cols_must_exist_lst(cols=cast(list[str], cols), row=row, param=param,
                        tinfo=tinfo)


def _cols_must_exist_dict(rule: TestRule, row: TestRow, param: str,
                          tinfo: TestColumn) -> None:
    """Call cols_must_exist_dict with a narrowed rule type."""
    if isinstance(tinfo, int):
        assert isinstance(row, list)
        cols_must_exist_dict(rule=cast(Rule[int], rule), row=row, param=param,
                             tinfo=tinfo)
        return
    assert isinstance(row, dict)
    cols_must_exist_dict(rule=cast(Rule[str], rule), row=row, param=param,
                         tinfo=tinfo)


def _cols_must_exist_multi(rule: TestRuleMulti, row: TestRow, param: str,
                           tinfo: TestColumn) -> None:
    """Call cols_must_exist_multi with a narrowed merge-rule type."""
    if isinstance(tinfo, int):
        assert isinstance(row, list)
        cols_must_exist_multi(rule=cast(RuleMerge[int], rule), row=row,
                              param=param, tinfo=tinfo)
        return
    assert isinstance(row, dict)
    cols_must_exist_multi(rule=cast(RuleMerge[str], rule), row=row,
                          param=param, tinfo=tinfo)


def _pop_from_row(row: TestRow, colref: TestColumn) -> Value:
    """Call pop_from_row with a narrowed column-reference type."""
    if isinstance(colref, int):
        return pop_from_row(row=row, colref=colref)
    return pop_from_row(row=row, colref=colref)


def _insert_into_row(row: TestRow, colref: TestColumn, val: Value) -> None:
    """Call insert_into_row with a narrowed column-reference type."""
    if isinstance(colref, int):
        insert_into_row(row=row, colref=colref, val=val)
        return
    insert_into_row(row=row, colref=colref, val=val)


@pytest.mark.parametrize('col, row, par',
                         [(1, ['a', 'b', 'c'], 'test1'),
                          (0, ['a', 'b', 'c'], 'test2'),
                          (2, ['a', 'b', 'c'], 'test3'),
                          ('x', {'x': 'a', 'y': 'b', 'z': 'c'}, 'test4'),
                          ('y', {'x': 'a', 'y': 'b', 'z': 'c'}, 'test5'),
                          ('z', {'x': 'a', 'y': 'b', 'z': 'c'}, 'test6')])
def test_col_must_exist_ok(capsys: CaptureFixture[str], col: TestColumn,
                           row: TestRow, par: str) -> None:
    """Test OK cases of col_must_exist."""
    _col_must_exist(col=col, row=row, param=par)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('col, row, par, msg',
                         [(-1, ['a', 'b', 'c'], 'test1',
                           'test1: column index -1 out of range [0, 2]'),
                          (3, ['a', 'b', 'c'], 'test2',
                           'test2: column index 3 out of range [0, 2]'),
                          (7, ['a', 'b', 'c', 'd'], 'test3',
                           'test3: column index 7 out of range [0, 3]'),
                          ('q', {'x': 'a', 'y': 'b', 'z': 'c'}, 'test4',
                           'test4: no column named "q" in data row'),
                          ('a', {'x': 'a', 'y': 'b', 'z': 'c'}, 'test5',
                           'test5: no column named "a" in data row'),
                          ('c', {'x': 'a', 'y': 'b', 'z': 'c'}, 'test6',
                           'test6: no column named "c" in data row')])
def test_col_must_exist_nok(capsys: CaptureFixture[str], col: TestColumn,
                            row: TestRow, par: str, msg: str) -> None:
    """Test not OK cases of col_must_exist."""
    with pytest.raises(SystemExit):
        _col_must_exist(col=col, row=row, param=par)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('collst, row, par, tinf',
                         [([2, 1, 3], ['a', 'b', 'c', 'd'], 'test1', 2),
                          (['z', 'x', 'y'],
                           {'x': 'a', 'y': 'b', 'z': 'c', 'q': 'd'},
                           'test2', 'a')])
def test_cols_must_exist_ok(capsys: CaptureFixture[str],
                            collst: list[int] | list[str], row: TestRow,
                            par: str, tinf: TestColumn) -> None:
    """Test OK cases of cols_must_exist_lst."""
    _cols_must_exist_lst(cols=collst, row=row, param=par, tinfo=tinf)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('collst, row, par, tinf, msg',
                         [([2, 10, 3], ['a', 'b', 'c', 'd'], 'test1', 2,
                           'test1: column index 10 out of range [0, 3]'),
                          (['z', 'aa', 'y'],
                           {'x': 'a', 'y': 'b', 'z': 'c', 'q': 'd'},
                           'test2', 'a',
                           'test2: no column named "aa" in data row')])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_cols_must_exist_nok(capsys: CaptureFixture[str],
                             collst: list[int] | list[str], row: TestRow,
                             par: str, tinf: TestColumn, msg: str) -> None:
    """Test OK cases of cols_must_exist_lst."""
    with pytest.raises(SystemExit):
        _cols_must_exist_lst(cols=collst, row=row, param=par, tinfo=tinf)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('rule, row, par, tinf',
                         [([{'column': 2, 'name': 'abc'},
                            {'column': 1, 'name': 'def'},
                            {'column': 3, 'name': 'ghi'}],
                           ['a', 'b', 'c', 'd'], 'test1', 2),
                          ([{'column': 'z', 'name': 'abc'},
                            {'column': 'x', 'name': 'abc'},
                            {'column': 'y', 'name': 'abc'}],
                           {'x': 'a', 'y': 'b', 'z': 'c', 'q': 'd'},
                           'test2', 'a')])
def test_cols_must_exst_dct_o(capsys: CaptureFixture[str], rule: TestRule,
                              row: TestRow, par: str,
                              tinf: TestColumn) -> None:
    """Test OK cases of cols_must_exist_dict."""
    _cols_must_exist_dict(rule=rule, row=row, param=par, tinfo=tinf)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('rule, row, par, tinf, msg',
                         [([{'column': 2, 'name': 'abc'},
                            {'column': 10, 'name': 'def'},
                            {'column': 3, 'name': 'ghi'}],
                           ['a', 'b', 'c', 'd'], 'test1', 2,
                           'test1: column index 10 out of range [0, 3]'),
                          ([{'column': 'z', 'name': 'abc'},
                            {'column': 'xx', 'name': 'abc'},
                            {'column': 'y', 'name': 'abc'}],
                           {'x': 'a', 'y': 'b', 'z': 'c', 'q': 'd'},
                           'test2', 'a',
                           'test2: no column named "xx" in data row')])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_cols_must_exst_dct_n(capsys: CaptureFixture[str], rule: TestRule,
                              row: TestRow, par: str, tinf: TestColumn,
                              msg: str) -> None:
    """Test not OK cases of cols_must_exist_dict."""
    with pytest.raises(SystemExit):
        _cols_must_exist_dict(rule=rule, row=row, param=par, tinfo=tinf)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('rule, row, par, tinf',
                         [([{'columns': [2, 1], 'separator': ' '},
                            {'columns': [3, 0], 'separator': ' '}],
                           ['a', 'b', 'c', 'd'], 'test1', 2),
                          ([{'columns': ['y', 'x'], 'separator': ' '},
                            {'columns': ['z', 'q'], 'separator': ' '}],
                           {'x': 'a', 'y': 'b', 'z': 'c', 'q': 'd'},
                           'test2', 'a')])
def test_cols_must_exst_dlst(capsys: CaptureFixture[str], rule: TestRuleMulti,
                             row: TestRow, par: str, tinf: TestColumn) -> None:
    """Test OK cases of cols_must_exist_multi."""
    _cols_must_exist_multi(rule=rule, row=row, param=par, tinfo=tinf)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('rule, row, par, tinf, msg',
                         [([{'columns': [2, 1], 'separator': ' '},
                            {'columns': [-1, 0], 'separator': ' '}],
                           ['a', 'b', 'c', 'd'], 'test1', 2,
                           'test1: column index -1 out of range [0, 3]'),
                          ([{'columns': ['y', 'x'], 'separator': ' '},
                            {'columns': ['z', 'aa'], 'separator': ' '}],
                           {'x': 'a', 'y': 'b', 'z': 'c', 'q': 'd'},
                           'test2', 'a',
                           'test2: no column named "aa" in data row')])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_cols_must_exst_dls_2(capsys: CaptureFixture[str], rule: TestRuleMulti,
                              row: TestRow, par: str, tinf: TestColumn,
                              msg: str) -> None:
    """Test not OK cases of cols_must_exist_multi."""
    with pytest.raises(SystemExit):
        _cols_must_exist_multi(rule=rule, row=row, param=par, tinfo=tinf)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('inrow, idx, resval, resrow',
                         [(['a', 'b', 'c'], 1, 'b', ['a', 'c']),
                          (['a', 'b', 'c'], 0, 'a', ['b', 'c']),
                          (['a', 'b', 'c'], 2, 'c', ['a', 'b']),
                          ({'x': 'a', 'y': 'b', 'z': 'c'},
                           'y', 'b', {'x': 'a', 'z': 'c'})])
def test_pop_from_row_ok(capsys: CaptureFixture[str], inrow: TestRow,
                         idx: TestColumn, resval: Value,
                         resrow: TestRow) -> None:
    """Test OK cases for pop_from_row."""
    row = deepcopy(inrow)
    ret = _pop_from_row(row=row, colref=idx)
    out, err = capsys.readouterr()
    assert ret == resval
    assert row == resrow
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('inrow, idx, exc',
                         [(['a', 'b', 'c'], 'a', AssertionError),
                          (['a', 'b', 'c'], 4, IndexError),
                          ({'x': 'a', 'y': 'b', 'z': 'c'},
                           1, AssertionError)])
def test_pop_from_row_nok(capsys: CaptureFixture[str], inrow: TestRow,
                          idx: TestColumn, exc: type[BaseException]) -> None:
    """Test not OK cases for pop_from_row."""
    row = deepcopy(inrow)
    with pytest.raises(exc):
        _ = _pop_from_row(row=row, colref=idx)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('inrow, idx, val, resrow',
                         [(['a', 'b', 'c'], 1, 'd',
                           ['a', 'd', 'b', 'c']),
                          (['a', 'b', 'c'], 0, 'd',
                           ['d', 'a', 'b', 'c']),
                          (['a', 'b', 'c'], 3, 'd',
                           ['a', 'b', 'c', 'd']),
                          ({'x': 'a', 'y': 'b', 'z': 'c'},
                           'q', 'p',
                           {'x': 'a', 'y': 'b', 'z': 'c', 'q': 'p'})])
def test_insert_into_row_ok(capsys: CaptureFixture[str], inrow: TestRow,
                            idx: TestColumn, val: Value,
                            resrow: TestRow) -> None:
    """Test OK cases for pop_from_row."""
    row = deepcopy(inrow)
    _insert_into_row(row=row, colref=idx, val=val)
    out, err = capsys.readouterr()
    assert row == resrow
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('inrow, idx, val, exc',
                         [(['a', 'b', 'c'], 'a', 'd', AssertionError),
                          ({'x': 'a', 'y': 'b', 'z': 'c'},
                           1, 'd', AssertionError)])
def test_insert_into_row_nok(capsys: CaptureFixture[str], inrow: TestRow,
                             idx: TestColumn, val: Value,
                             exc: type[BaseException]) -> None:
    """Test not OK cases for pop_from_row."""
    row = deepcopy(inrow)
    with pytest.raises(exc):
        _insert_into_row(row=row, colref=idx, val=val)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
