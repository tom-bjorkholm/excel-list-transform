#! /usr/local/bin/python3
"""Test data conversion between numbered columns and named columns."""

# Copyright (c) 2024 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from copy import deepcopy
import pytest
from excel_list_transform.num_named_conversion import \
    named_cols_from_num_cols, num_cols_from_named_cols


@pytest.mark.parametrize('ind, outd',
                         [([['ab', 'cd', 'ef'],
                            [1, 2, 3], [4, 5, 6]],
                           [{'ab': 1, 'cd': 2, 'ef': 3},
                            {'ab': 4, 'cd': 5, 'ef': 6}]),
                          ([['ab', 'cd', 'ef'],
                            [1, None, 3], ['a', 'x y', 7]],
                           [{'ab': 1, 'cd': None, 'ef': 3},
                            {'ab': 'a', 'cd': 'x y', 'ef': 7}]),
                          ([['ab', 'cd', 'ef'],
                            [1, None], ['a', 'x y', 7]],
                           [{'ab': 1, 'cd': None, 'ef': None},
                            {'ab': 'a', 'cd': 'x y', 'ef': 7}])])
def test_nam_col_num_cols_ok(capsys, ind, outd):
    """Test OK cases of named_cols_from_num_cols."""
    ret = named_cols_from_num_cols(ind, 'somfile.csv')
    out, err = capsys.readouterr()
    assert outd == ret
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('ind, fname, msg',
                         [([['ab', None, 'ef'],
                            [1, 2, 3], [4, 5, 6]],
                           'a.csv', 'Cannot handle input column 1 without ' +
                           'name in file a.csv'),
                          ([['ab', 'cd', 'ef'],
                            [1, 2, 3], [4, 5, 6, 7]],
                           'b.dat', 'Data row(s) have more columns than ' +
                           'title row in file b.dat')])
def test_nam_col_num_cols_nok(capsys, ind, fname, msg):
    """Test not OK cases of named_cols_from_num_cols."""
    with pytest.raises(SystemExit):
        _ = named_cols_from_num_cols(ind, fname)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('ind, outd, corder',
                         [([{'ab': 1, 'cd': 2, 'ef': 3},
                            {'ab': 4, 'cd': 5, 'ef': 6}],
                           [['ef', 'ab', 'cd'],
                            [3, 1, 2], [6, 4, 5]],
                           ['ef', 'ab', 'cd']),
                          ([{'ab': 'x', 'cd': None, 'ef': 'b r'},
                            {'cd': 5, 'ef': 6, 'ab': None}],
                           [['ef', 'ab', 'cd'],
                            ['b r', 'x', None], [6, None, 5]],
                           ['ef', 'ab', 'cd']),
                          ([{'ab': 1, 'cd': 2, 'ef': 3},
                            {'ab': 4, 'cd': 5, 'ef': 6}],
                           [['ab', 'ef', 'cd'],
                            [1, 3, 2], [4, 6, 5]],
                           ['ab', 'ef', 'cd'])])
def test_num_col_nam_cols_ok(capsys, ind, outd, corder):
    """Test OK cases of num_cols_from_named_cols."""
    ret = num_cols_from_named_cols(ind, column_order=corder)
    out, err = capsys.readouterr()
    assert outd == ret
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('ind, corder, msg',
                         [([{'ab': 'x', 'cd': None, 'ef': 'b r'},
                            {'cd': 5, 'ef': 6}],
                           ['ef', 'ab', 'cd'],
                           'Data row 1 is missing data for column ab')])
def test_num_col_nam_cols_nok(capsys, ind, corder, msg):
    """Test not OK cases of named_cols_from_num_cols."""
    with pytest.raises(SystemExit):
        _ = num_cols_from_named_cols(ind, column_order=corder)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('ind',
                         [[['ab', 'cd', 'ef'],
                           [1, 2, 3], [4, 5, 6]],
                          [['ab', 'cd', 'ef'],
                           [1, None, 3], ['a', 'x y', 7]],
                          [['ab', 'cd', 'ef'],
                           [1, None, None], ['a', 'x y', 7]]])
def test_nam_col_num_cols_rev(capsys, ind):
    """Test OK cases of named_cols_from_num_cols."""
    exp = deepcopy(ind)
    ret1 = named_cols_from_num_cols(ind, 'somfile.csv')
    ret2 = num_cols_from_named_cols(ret1, column_order=ind[0])
    out, err = capsys.readouterr()
    assert exp == ret2
    assert '' == out
    assert '' == err
