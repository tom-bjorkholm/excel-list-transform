#! /usr/local/bin/python3
"""Test the rewriting of a single value."""

# Copyright (c) 2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from copy import deepcopy
from datetime import datetime
import pytest
from excel_list_transform.row_split_merge import get_nosep_pos, \
    in_nosep_pos, split_one_str, one_split_one_row, one_split, \
    split_rows, split_rows_cfg, merge_strings, \
    merge_identified_rows, identify_rows_to_merge, \
    one_merge_rows, one_rule_merge_rows, merge_rows, \
    merge_rows_cfg
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName


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
    ret = get_nosep_pos(instr=deepcopy(instr), not_separators=nseps)
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
    ret = in_nosep_pos(pos=deepcopy(pos), nosep_pos=deepcopy(nseps))
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
    ret = split_one_str(instr=deepcopy(instr), separators=deepcopy(sep),
                        not_separators=deepcopy(nosep))
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
    ret = one_split_one_row(inrow=deepcopy(inrow), column=deepcopy(col),
                            separators=deepcopy(seps),
                            not_separators=deepcopy(noseps))
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
        _ = one_split_one_row(inrow=deepcopy(inrow),
                              column=deepcopy(col),
                              separators=deepcopy(seps),
                              not_separators=deepcopy(noseps))
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err


@pytest.mark.parametrize('indata,col,seps,noseps,res',
                         [([], 'b', [' ', ';'], ['  ', '\\;'], []),
                          ([{'a': 'b', 'c': 'd e', 'f': 'g'}], 'c',
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
    ret = one_split(indata=deepcopy(indata), column=deepcopy(col),
                    separators=deepcopy(seps),
                    not_separators=deepcopy(noseps))
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
        _ = one_split(indata=deepcopy(indata), column=deepcopy(col),
                      separators=deepcopy(seps),
                      not_separators=deepcopy(noseps))
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err


@pytest.mark.parametrize('indata,direc,res',
                         [([], [], []),
                          ([], [{'column': 'c',
                                 'separators': [' ', ';'],
                                 'not_separators': ['\\;', '  ']}], []),
                          ([{'a': 'b', 'c': 'd e', 'f': 'g'},
                            {'a': 'h+j', 'c': 'k', 'f': 'l;m'},
                            {'a': 'n', 'c': 'o', 'f': 'p'}],
                           [{'column': 'c', 'separators': [' ', 'x'],
                             'not_separators': ['   ']},
                            {'column': 'a', 'separators': ['+', 'x'],
                             'not_separators': ['xx', 'xy']},
                            {'column': 'f', 'separators': [';', 'x'],
                             'not_separators': ['\\;']}],
                           [{'a': 'b', 'c': 'd', 'f': 'g'},
                            {'a': 'b', 'c': 'e', 'f': 'g'},
                            {'a': 'h', 'c': 'k', 'f': 'l'},
                            {'a': 'h', 'c': 'k', 'f': 'm'},
                            {'a': 'j', 'c': 'k', 'f': 'l'},
                            {'a': 'j', 'c': 'k', 'f': 'm'},
                            {'a': 'n', 'c': 'o', 'f': 'p'}])])
def test_split_rows_name_ok1(capsys, indata, direc, res):
    """Test OK cases for split_rows_name."""
    ret = split_rows(indata=deepcopy(indata), directives=deepcopy(direc))
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('direc,err,msg',
                         [(2, TypeError, 'object is not iterable'),
                          ('not list', TypeError,
                           'string indices must be integers'),
                          (['not dict'], TypeError,
                           'string indices must be integers'),
                          ([{'colum': 'c', 'separators': ['a', 'b'],
                             'not_separators': ['aa', 'bb']}],
                           KeyError, "KeyError('column')"),
                          ([{'column': 'c', 'separator': ['a', 'b'],
                             'not_separators': ['aa', 'bb']}],
                           KeyError, "KeyError('separators')"),
                          ([{'column': 'c', 'separators': ['a', 'b'],
                             'not_separator': ['aa', 'bb']}],
                           KeyError, "KeyError('not_separators')"),
                          ([{'column': 1, 'separators': ['a', 'b'],
                             'not_separators': ['aa', 'bb']}],
                           AssertionError, 'ExceptionInfo AssertionError'),
                          ([{'column': 'a', 'separators': ['a', 'b'],
                             'not_separators': 'bb'}],
                           AssertionError, 'ExceptionInfo AssertionError'),
                          ([{'column': 1, 'separators': 'a',
                             'not_separators': ['aa', 'bb']}],
                           AssertionError, 'ExceptionInfo AssertionError')])
def test_slit_rows_name_nok1(capsys, direc, err, msg):
    """Test not OK cases for split_rows_name."""
    indata = [{'a': 'b', 'c': 'd'}]
    with pytest.raises(err) as exc:
        _ = split_rows(indata=deepcopy(indata), directives=direc)
    out, err = capsys.readouterr()
    assert msg in str(exc)
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('indata,direc,res',
                         [([], [], []),
                          ([], [{'column': 'c',
                                 'separators': [' ', ';'],
                                 'not_separators': ['\\;', '  ']}], []),
                          ([{'a': 'b', 'c': 'd e', 'f': 'g'},
                            {'a': 'h+j', 'c': 'k', 'f': 'l;m'},
                            {'a': 'n', 'c': 'o', 'f': 'p'}],
                           [{'column': 'c', 'separators': [' ', 'x'],
                             'not_separators': ['   ']},
                            {'column': 'a', 'separators': ['+', 'x'],
                             'not_separators': ['xx', 'xy']},
                            {'column': 'f', 'separators': [';', 'x'],
                             'not_separators': ['\\;']}],
                           [{'a': 'b', 'c': 'd', 'f': 'g'},
                            {'a': 'b', 'c': 'e', 'f': 'g'},
                            {'a': 'h', 'c': 'k', 'f': 'l'},
                            {'a': 'h', 'c': 'k', 'f': 'm'},
                            {'a': 'j', 'c': 'k', 'f': 'l'},
                            {'a': 'j', 'c': 'k', 'f': 'm'},
                            {'a': 'n', 'c': 'o', 'f': 'p'}])])
def test_split_rows_namecfg_ok1(capsys, indata, direc, res):
    """Test OK cases for split_rows_namecfg."""
    cfg1 = ConfigXlsListTransfName()
    cfg1.s01_split_rows = direc
    jsontxt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfName(from_json_text=jsontxt)
    ret = split_rows_cfg(indata=deepcopy(indata), cfg=cfg2, tinfo='a')
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('inlst,sep,res',
                         [(['a', 'b', 2, 'b', 'a'], ' ', 'a b 2'),
                          ([], ';', None),
                          ([1, 1, 1, 1], '+', 1),
                          ([1.23, 'a', 1.23, datetime(year=2025, month=6,
                                                      day=1, hour=12,
                                                      minute=13, second=14),
                            1.23, 'a', 1.23, datetime(year=2025, month=6,
                                                      day=1, hour=12,
                                                      minute=13, second=14)],
                           '@', '1.23@a@2025-06-01 12:13:14')])
def test_merge_strings(capsys, inlst, sep, res):
    """Test OK cases of merge_strings."""
    ret = merge_strings(to_merge=inlst, sep=sep)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('data,rowidxs,sep,res',
                         [([], [], '+', []),
                          ([], [[]], '+', []),
                          ([], [[2, 4]], '+', []),
                          ([{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3}],
                           [], '+',
                           [{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3}]),
                          ([{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3}],
                           [[]], '+',
                           [{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3}]),
                          ([{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3}],
                           [[1]], '+',
                           [{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3}]),
                          ([{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3},
                            {'a': 'x1', 'b': 'y', 'c': 'z1'},
                            {'a': 4, 'b': 5, 'c': 6}],
                           [[0, 2, 3]], '+',
                           [{'a': 'x+x1+4', 'b': 'y+5', 'c': 'z+z1+6'},
                            {'a': 1, 'b': 2, 'c': 3}]),
                          ([{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3},
                            {'a': 'x1', 'b': 'y', 'c': 'z1'},
                            {'a': 4, 'b': 5, 'c': 6}],
                           [[2, 0]], '-',
                           [{'a': 1, 'b': 2, 'c': 3},
                            {'a': 'x1-x', 'b': 'y', 'c': 'z1-z'},
                            {'a': 4, 'b': 5, 'c': 6}]),
                          ([{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3},
                            {'a': 'x1', 'b': 'y', 'c': 'z1'},
                            {'a': 4, 'b': 5, 'c': 6}],
                           [[0, 2], [1]], '-',
                           [{'a': 'x-x1', 'b': 'y', 'c': 'z-z1'},
                            {'a': 1, 'b': 2, 'c': 3},
                            {'a': 4, 'b': 5, 'c': 6}]),
                          ([{'a': 'x', 'b': 'y', 'c': 'z'},
                            {'a': 1, 'b': 2, 'c': 3},
                            {'a': 'x1', 'b': 'y', 'c': 'z1'},
                            {'a': 4, 'b': 5, 'c': 6}],
                           [[0, 2], [1, 3]], '-',
                           [{'a': 'x-x1', 'b': 'y', 'c': 'z-z1'},
                            {'a': '1-4', 'b': '2-5', 'c': '3-6'}])])
def test_merge_ident_rows_name(capsys, data, rowidxs, sep, res):
    """Test OK cases of merge_identified_rows."""
    ret = merge_identified_rows(rows=deepcopy(data),
                                row_numbers=deepcopy(rowidxs),
                                separator=sep, tinfo='a')
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


DATA1 = [{'a': 'aa', 'b': 'bb', 'c': 3},
         {'a': 'ff', 'b': 'bc', 'c': 5},
         {'a': 'aa', 'b': 'bd', 'c': 3},
         {'a': 'ff', 'b': 'be', 'c': 5},
         {'a': 'gg', 'b': 'bf', 'c': 5},
         {'a': 'aa', 'b': 'bg', 'c': 7},
         {'a': 'aa', 'b': 'bh', 'c': 3}]


@pytest.mark.parametrize('data,cols,res',
                         [(DATA1, ['a', 'c'],
                           [[0, 2, 6], [1, 3]]),
                          (DATA1, ['b'], []),
                          (DATA1, [],
                           [[0, 1, 2, 3, 4, 5, 6]]),
                          (DATA1, ['c'],
                           [[0, 2, 6], [1, 3, 4]])])
def test_iden_rows_to_merge1(capsys, data, cols, res):
    """Test OK cases of merge identified rows."""
    ret = identify_rows_to_merge(rows=deepcopy(data),
                                 columns_to_cmp=deepcopy(cols),
                                 tinfo='a')
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


DATA1ACP = [{'a': 'aa', 'b': 'bb+bd+bh', 'c': 3},
            {'a': 'ff', 'b': 'bc+be', 'c': 5},
            {'a': 'gg', 'b': 'bf', 'c': 5},
            {'a': 'aa', 'b': 'bg', 'c': 7}]

DATA2ACCPM = [{'a': 'aa', 'b': 'bb+bd+bh', 'c': 3},
              {'a': 'ff-gg', 'b': 'bc+be-bf', 'c': 5},
              {'a': 'aa', 'b': 'bg', 'c': 7}]

DATA1NP = [{'a': 'aa+ff+gg', 'b': 'bb+bc+bd+be+bf+bg+bh', 'c': '3+5+7'}]

DATA1CP = [{'a': 'aa', 'b': 'bb+bd+bh', 'c': 3},
           {'a': 'ff+gg', 'b': 'bc+be+bf', 'c': 5},
           {'a': 'aa', 'b': 'bg', 'c': 7}]

DATA1ACS = [{'a': 'aa', 'b': 'bb bd bh', 'c': 3},
            {'a': 'ff', 'b': 'bc be', 'c': 5},
            {'a': 'gg', 'b': 'bf', 'c': 5},
            {'a': 'aa', 'b': 'bg', 'c': 7}]

DATA1NS = [{'a': 'aa ff gg', 'b': 'bb bc bd be bf bg bh', 'c': '3 5 7'}]

DATA1CS = [{'a': 'aa', 'b': 'bb bd bh', 'c': 3},
           {'a': 'ff gg', 'b': 'bc be bf', 'c': 5},
           {'a': 'aa', 'b': 'bg', 'c': 7}]


@pytest.mark.parametrize('data,cols,sep,res',
                         [(DATA1, ['a', 'c'], '+',
                           DATA1ACP),
                          (DATA1, ['b'], '+', DATA1),
                          (DATA1, [], '+', DATA1NP),
                          (DATA1, ['c'], '+', DATA1CP),
                          (DATA1, ['a', 'c'], ' ',
                           DATA1ACS),
                          (DATA1, ['b'], ' ', DATA1),
                          (DATA1, [], ' ', DATA1NS),
                          (DATA1, ['c'], ' ', DATA1CS)])
def test_one_merge_rows_name_ok1(capsys, data, cols, sep, res):
    """Test OK cases of one merge rows name."""
    ret = one_merge_rows(indata=deepcopy(data),
                         columns_to_cmp=deepcopy(cols),
                         separator=deepcopy(sep), tinfo='a')
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('data,cols,sep,res',
                         [(DATA1, ['a', 'c'], '+',
                           DATA1ACP),
                          (DATA1, ['b'], '+', DATA1),
                          (DATA1, [], '+', DATA1NP),
                          (DATA1, ['c'], '+', DATA1CP),
                          (DATA1, ['a', 'c'], ' ',
                           DATA1ACS),
                          (DATA1, ['b'], ' ', DATA1),
                          (DATA1, [], ' ', DATA1NS),
                          (DATA1, ['c'], ' ', DATA1CS)])
def test_one_rule_merge_rows_na_ok1(capsys, data, cols, sep, res):
    """Test OK cases of one merge rows name."""
    rule = {'columns': deepcopy(cols), 'separator': deepcopy(sep)}
    ret = one_rule_merge_rows(indata=deepcopy(data), rule=rule, tinfo='a')
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('data,cols,sep,res',
                         [(DATA1, ['a', 'c'], '+',
                           DATA1ACP),
                          (DATA1, ['b'], '+', DATA1),
                          (DATA1, ['c'], '+', DATA1CP),
                          (DATA1, ['a', 'c'], ' ',
                           DATA1ACS),
                          (DATA1, ['b'], ' ', DATA1),
                          (DATA1, ['c'], ' ', DATA1CS)])
def test_merge_rows_name_ok1(capsys, data, cols, sep, res):
    """Test OK cases of merge rows name."""
    rule = {'columns': deepcopy(cols), 'separator': deepcopy(sep)}
    cfg1 = ConfigXlsListTransfName()
    cfg1.s02_merge_rows = [rule]
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfName(from_json_text=txt)
    ret = merge_rows(indata=deepcopy(data), rules=cfg2.s02_merge_rows,
                     tinfo='a')
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('data,rule,res',
                         [(DATA1, [{'columns': ['a', 'c'],
                                    'separator': '+'},
                                   {'columns': ['c'],
                                    'separator': '-'}],
                           DATA2ACCPM)])
def test_merge_rows_name_ok2(capsys, data, rule, res):
    """Test OK cases of merge rows name more than one rule."""
    cfg1 = ConfigXlsListTransfName()
    cfg1.s02_merge_rows = deepcopy(rule)
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfName(from_json_text=txt)
    ret = merge_rows(indata=deepcopy(data), rules=cfg2.s02_merge_rows,
                     tinfo='a')
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('data,cols,sep,res',
                         [(DATA1, ['a', 'c'], '+',
                           DATA1ACP),
                          (DATA1, ['b'], '+', DATA1),
                          (DATA1, ['c'], '+', DATA1CP),
                          (DATA1, ['a', 'c'], ' ',
                           DATA1ACS),
                          (DATA1, ['b'], ' ', DATA1),
                          (DATA1, ['c'], ' ', DATA1CS)])
def test_merge_rows_name_ok3(capsys, data, cols, sep, res):
    """Test OK cases of merge rows name."""
    rule = {'columns': deepcopy(cols), 'separator': deepcopy(sep)}
    cfg1 = ConfigXlsListTransfName()
    cfg1.s02_merge_rows = [rule]
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfName(from_json_text=txt)
    ret = merge_rows_cfg(indata=deepcopy(data), cfg=cfg2, tinfo='a')
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('data,rule,res',
                         [(DATA1, [{'columns': ['a', 'c'],
                                    'separator': '+'},
                                   {'columns': ['c'],
                                    'separator': '-'}],
                           DATA2ACCPM)])
def test_merge_rows_name_ok4(capsys, data, rule, res):
    """Test OK cases of merge rows name more than one rule."""
    cfg1 = ConfigXlsListTransfName()
    cfg1.s02_merge_rows = deepcopy(rule)
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfName(from_json_text=txt)
    ret = merge_rows_cfg(indata=deepcopy(data), cfg=cfg2, tinfo='a')
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err
