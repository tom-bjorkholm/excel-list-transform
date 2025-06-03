#! /usr/local/bin/python3
"""Test the excel_list_transform common functions functionality."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from copy import deepcopy
import pytest
from excel_list_transform.transform_func_common import col_must_exist, \
    cols_must_exist_lst, cols_must_exist_dict, cols_must_exist_dictlst, \
    pop_from_row, insert_into_row


@pytest.mark.parametrize('col, row, par',
                         [(1, ['a', 'b', 'c'], 'test1'),
                          (0, ['a', 'b', 'c'], 'test2'),
                          (2, ['a', 'b', 'c'], 'test3'),
                          ('x', {'x': 'a', 'y': 'b', 'z': 'c'}, 'test4'),
                          ('y', {'x': 'a', 'y': 'b', 'z': 'c'}, 'test5'),
                          ('z', {'x': 'a', 'y': 'b', 'z': 'c'}, 'test6')])
def test_col_must_exist_ok(capsys, col, row, par):
    """Test OK cases of col_must_exist."""
    col_must_exist(col=col, row=row, param=par)
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
def test_col_must_exist_nok(capsys, col, row, par, msg):
    """Test not OK cases of col_must_exist."""
    with pytest.raises(SystemExit):
        col_must_exist(col=col, row=row, param=par)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('collst, row, par, tinf',
                         [([2, 1, 3], ['a', 'b', 'c', 'd'], 'test1', 2),
                          (['z', 'x', 'y'],
                           {'x': 'a', 'y': 'b', 'z': 'c', 'q': 'd'},
                           'test2', 'a')])
def test_cols_must_exist_lst_ok(capsys, collst, row, par, tinf):
    """Test OK cases of cols_must_exist_lst."""
    cols_must_exist_lst(cols=collst, row=row, param=par, tinfo=tinf)
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
def test_cols_must_exist_lst_nok(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments  # noqa: E501
                                 collst, row, par, tinf, msg):
    """Test OK cases of cols_must_exist_lst."""
    with pytest.raises(SystemExit):
        cols_must_exist_lst(cols=collst, row=row, param=par, tinfo=tinf)
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
def test_cols_must_exist_dict_ok(capsys, rule, row, par, tinf):
    """Test OK cases of cols_must_exist_dict."""
    cols_must_exist_dict(rule=rule, row=row, param=par, tinfo=tinf)
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
def test_cols_must_exist_dict_nok(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments  # noqa: E501
                                  rule, row, par, tinf, msg):
    """Test not OK cases of cols_must_exist_dict."""
    with pytest.raises(SystemExit):
        cols_must_exist_dict(rule=rule, row=row, param=par, tinfo=tinf)
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
def test_cols_must_exist_dlst_ok(capsys, rule, row, par, tinf):
    """Test OK cases of cols_must_exist_dictlst."""
    cols_must_exist_dictlst(rule=rule, row=row, param=par, tinfo=tinf)
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
def test_cols_must_exist_dlst_nok(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments  # noqa: E501
                                  rule, row, par, tinf, msg):
    """Test not OK cases of cols_must_exist_dictlst."""
    with pytest.raises(SystemExit):
        cols_must_exist_dictlst(rule=rule, row=row, param=par, tinfo=tinf)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('inrow, idx, resval, resrow',
                         [(['a', 'b', 'c'], 1, 'b', ['a', 'c']),
                          (['a', 'b', 'c'], 0, 'a', ['b', 'c']),
                          (['a', 'b', 'c'], 2, 'c', ['a', 'b']),
                          ({'x': 'a', 'y': 'b', 'z': 'c'},
                           'y', 'b', {'x': 'a', 'z': 'c'})])
def test_pop_from_row_ok(capsys, inrow, idx, resval, resrow):
    """Test OK cases for pop_from_row."""
    row = deepcopy(inrow)
    ret = pop_from_row(row=row, colref=idx)
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
def test_pop_from_row_nok(capsys, inrow, idx, exc):
    """Test not OK cases for pop_from_row."""
    row = deepcopy(inrow)
    with pytest.raises(exc):
        _ = pop_from_row(row=row, colref=idx)
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
def test_insert_into_row_ok(capsys, inrow, idx, val, resrow):
    """Test OK cases for pop_from_row."""
    row = deepcopy(inrow)
    insert_into_row(row=row, colref=idx, val=val)
    out, err = capsys.readouterr()
    assert row == resrow
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('inrow, idx, val, exc',
                         [(['a', 'b', 'c'], 'a', 'd', AssertionError),
                          ({'x': 'a', 'y': 'b', 'z': 'c'},
                           1, 'd', AssertionError)])
def test_insert_into_row_nok(capsys, inrow, idx, val, exc):
    """Test not OK cases for pop_from_row."""
    row = deepcopy(inrow)
    with pytest.raises(exc):
        insert_into_row(row=row, colref=idx, val=val)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
