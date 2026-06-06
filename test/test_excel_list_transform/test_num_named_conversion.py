#! /usr/local/bin/python3
"""Test data conversion between numbered columns and named columns."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License

import pytest
from pytest import CaptureFixture
from excel_list_transform.num_named_conversion import named_cols_from_num_cols
from excel_list_transform.commontypes import NumData, NameData


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
def test_nam_col_num_cols_ok(capsys: CaptureFixture[str], ind: NumData,
                             outd: NameData) -> None:
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
def test_nam_col_num_cols_nok(capsys: CaptureFixture[str], ind: NumData,
                              fname: str, msg: str) -> None:
    """Test not OK cases of named_cols_from_num_cols."""
    with pytest.raises(SystemExit):
        _ = named_cols_from_num_cols(ind, fname)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err
