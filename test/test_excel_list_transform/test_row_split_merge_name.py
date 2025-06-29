#! /usr/local/bin/python3
"""Test the rewriting of a single value."""

# Copyright (c) 2025 Tom Björkholm
# MIT License

import pytest
from excel_list_transform.row_split_merge_name import get_nosep_pos, \
    in_nosep_pos, split_one_str, one_split_one_row_name, one_split_name


@pytest.mark.parametrize('instr,nseps,res',
                         [('abc', [], []),
                          ('abc def ddd', [' '],
                           [(3, 4), (7, 8)]),
                          ('abc def ddd', [' d', 'bc'],
                           [(1, 3), (3, 5), (7, 9)]),
                          ('abcdefghijk', ['bc', 'abc', 'gh', 'ghi'],
                           [(0, 3), (6, 9)])])
def test_get_nosep_pos1(capsys, instr, nseps, res):
    """Test normal cases for get_nosep_pos."""
    ret = get_nosep_pos(instr=instr, not_separators=nseps)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('pos,nseps,res',
                         [(4, [(2, 3), (7, 8)], False),
                          (15, [], False),
                          (4, [(2, 5)], True),
                          (5, [(1, 3), (5, 7), (10, 12)], True),
                          (5, [(1, 3), (4, 5), (7, 10)], False)])
def test_in_nosep_pos1(capsys, pos, nseps, res):
    """Test normal cases for in_nosep_pos."""
    ret = in_nosep_pos(pos=pos, nosep_pos=nseps)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('instr,sep,nosep,res',
                         [('abcdef', ['b', 'd'], [], ['a', 'c', 'ef']),
                          ('abcdef', ['b', 'd'], ['bc'], ['abc', 'ef']),
                          ('abcdef', ['b', 'd'], ['cd'], ['a', 'cdef']),
                          ('abcabcdabcdeabcdefabcdefg', ['bc', 'e'],
                           ['fff', 'efg'],
                           ['a', 'a', 'da', 'd', 'a', 'd', 'fa', 'defg'])])
def test_split_one_str(capsys, instr, sep, nosep, res):
    """Test normal cases for split_one_str."""
    ret = split_one_str(instr=instr, separators=sep, not_separators=nosep)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('inrow,col,seps,noseps,res',
                         [({'a': 'b', 'c': 'd e', 'f': 'g'}, 'c',
                           ['+', ' ', '-'], ['  ', '++'],
                          [{'a': 'b', 'c': 'd', 'f': 'g'},
                           {'a': 'b', 'c': 'e', 'f': 'g'}]),
                          ({'a': 'b', 'c': 'd e+x', 'f': 'g'}, 'c',
                           ['+', ' ', '-'], ['  ', '++'],
                          [{'a': 'b', 'c': 'd', 'f': 'g'},
                           {'a': 'b', 'c': 'e', 'f': 'g'},
                           {'a': 'b', 'c': 'x', 'f': 'g'}]),
                          ({'a': 'b', 'c': 'de', 'f': 'g'}, 'c',
                           ['+', ' ', '-'], ['  ', '++'],
                           [{'a': 'b', 'c': 'de', 'f': 'g'}])])
def test_one_split_one_na_ok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
                              inrow, col, seps, noseps, res):
    """Test OK cases of one_split_one_row_name."""
    ret = one_split_one_row_name(inrow=inrow, column=col,
                                 separators=seps, not_separators=noseps)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('inrow,col,seps,noseps,msgs',
                         [({'a': 'b', 'c': 'd e', 'f': 2}, 'f',
                           ['+', ' ', '-'], ['  ', '++'],
                          ['Trying to split rows based on column "f".',
                           'But that column has value of type int']),
                          ({'a': 'b', 'c': ['a'], 'f': 'g'}, 'c',
                           ['+', ' ', '-'], ['  ', '++'],
                          ['Trying to split rows based on column "c".',
                           'But that column has value of type list'])])
def test_one_split_one_na_nok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
                               inrow, col, seps, noseps, msgs):
    """Test not OK cases of one_split_one_row_name."""
    with pytest.raises(SystemExit):
        _ = one_split_one_row_name(inrow=inrow, column=col,
                                   separators=seps, not_separators=noseps)
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err


@pytest.mark.parametrize('indata,col,seps,noseps,res',
                         [([{'a': 'b', 'c': 'd e', 'f': 'g'}], 'c',
                           ['+', ' ', '-'], ['  ', '++'],
                          [{'a': 'b', 'c': 'd', 'f': 'g'},
                           {'a': 'b', 'c': 'e', 'f': 'g'}]),
                          ([{'a': 'b', 'c': 'd e+x', 'f': 'g'},
                            {'a': 'h', 'c': 'i+j', 'f': 'k'}], 'c',
                           ['+', ' ', '-'], ['  ', '++'],
                          [{'a': 'b', 'c': 'd', 'f': 'g'},
                           {'a': 'b', 'c': 'e', 'f': 'g'},
                           {'a': 'b', 'c': 'x', 'f': 'g'},
                           {'a': 'h', 'c': 'i', 'f': 'k'},
                           {'a': 'h', 'c': 'j', 'f': 'k'}]),
                          ([{'a': 'b', 'c': 'de', 'f': 'g'}], 'c',
                           ['+', ' ', '-'], ['  ', '++'],
                           [{'a': 'b', 'c': 'de', 'f': 'g'}])])
def test_one_split_name_ok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
                            indata, col, seps, noseps, res):
    """Test OK cases of one_split_name."""
    ret = one_split_name(indata=indata, column=col,
                         separators=seps, not_separators=noseps)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('indata,col,seps,noseps,msgs',
                         [([{'a': 'b', 'c': 'd e', 'f': 2}], 'k',
                           ['+', ' ', '-'], ['  ', '++'],
                           ['Trying to split lines based on column "k"',
                            'but no such column in data'])])
def test_one_split_name_nok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
                             indata, col, seps, noseps, msgs):
    """Test not OK cases of one_split_name."""
    with pytest.raises(SystemExit):
        _ = one_split_name(indata=indata, column=col,
                           separators=seps, not_separators=noseps)
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err
