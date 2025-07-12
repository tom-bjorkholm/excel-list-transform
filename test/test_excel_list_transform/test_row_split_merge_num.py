#! /usr/local/bin/python3
"""Test the rewriting of a single value."""

# Copyright (c) 2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from copy import deepcopy
from datetime import datetime
import pytest
from excel_list_transform.row_split_merge_name import \
    one_split_one_row, one_split, \
    split_rows, split_rows_cfg, merge_strings, \
    merge_identified_rows, identify_rows_to_merge, \
    one_merge_rows, one_rule_merge_rows, merge_rows, \
    merge_rows_cfg
from excel_list_transform.config_xls_list_transf_num import \
   ConfigXlsListTransfNum


@pytest.mark.parametrize('inrow,col,seps,noseps,res',
                         [(['b', 'd e', 'g'], 1,
                           ['+', ' ', '-'], ['  ', '++'],
                          [['b', 'd', 'g'],
                           ['b', 'e', 'g']]),
                          (['b', 'd e+x', 'g'], 1,
                           ['+', ' ', '-'], ['  ', '++'],
                          [['b', 'd', 'g'],
                           ['b', 'e', 'g'],
                           ['b', 'x', 'g']]),
                          (['b', 'de', 'g'], 1,
                           ['+', ' ', '-'], ['  ', '++'],
                           [['b', 'de', 'g']])])
def test_one_split_one_nu_ok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
                              inrow, col, seps, noseps, res):
    """Test OK cases of one_split_one_row num."""
    ret = one_split_one_row(inrow=deepcopy(inrow), column=deepcopy(col),
                            separators=deepcopy(seps),
                            not_separators=deepcopy(noseps))
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('inrow,col,seps,noseps,msgs',
                         [(['b', 'd e', 2], 2,
                           ['+', ' ', '-'], ['  ', '++'],
                          ['Trying to split rows based on column "2".',
                           'But that column has value of type int']),
                          (['b', ['a'], 'g'], 1,
                           ['+', ' ', '-'], ['  ', '++'],
                          ['Trying to split rows based on column "1".',
                           'But that column has value of type list'])])
def test_one_split_one_nu_nok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
                               inrow, col, seps, noseps, msgs):
    """Test not OK cases of one_split_one_row num."""
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
                          ([['b', 'd e', 'g']], 1,
                           ['+', ' ', '-'], ['  ', '++'],
                          [['b', 'd', 'g'],
                           ['b', 'e', 'g']]),
                          ([['b', 'd e+x', 'g'],
                            ['h', 'i+j', 'k']], 1,
                           ['+', ' ', '-'], ['  ', '++'],
                          [['b', 'd', 'g'],
                           ['b', 'e', 'g'],
                           ['b', 'x', 'g'],
                           ['h', 'i', 'k'],
                           ['h', 'j', 'k']]),
                          ([['b', 'de', 'g']], 1,
                           ['+', ' ', '-'], ['  ', '++'],
                           [['b', 'de', 'g']])])
def test_one_split_num_ok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
                           indata, col, seps, noseps, res):
    """Test OK cases of one_split num."""
    ret = one_split(indata=deepcopy(indata), column=deepcopy(col),
                    separators=deepcopy(seps),
                    not_separators=deepcopy(noseps))
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('indata,col,seps,noseps,msgs',
                         [([['b', 'd e', 2]], 7,
                           ['+', ' ', '-'], ['  ', '++'],
                           ['Trying to split lines based on column "7"',
                            'but no such column in data']),
                          ([['b', 'd e', 2]], -1,
                           ['+', ' ', '-'], ['  ', '++'],
                           ['Trying to split lines based on column "-1"',
                            'but no such column in data'])])
def test_one_split_num_nok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
                            indata, col, seps, noseps, msgs):
    """Test not OK cases of one_split num."""
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
                          ([], [{'column': 1,
                                 'separators': [' ', ';'],
                                 'not_separators': ['\\;', '  ']}], []),
                          ([['b', 'd e', 'g'],
                            ['h+j', 'k', 'l;m'],
                            ['n', 'o', 'p']],
                           [{'column': 1, 'separators': [' ', 'x'],
                             'not_separators': ['   ']},
                            {'column': 0, 'separators': ['+', 'x'],
                             'not_separators': ['xx', 'xy']},
                            {'column': 2, 'separators': [';', 'x'],
                             'not_separators': ['\\;']}],
                           [['b', 'd', 'g'],
                            ['b', 'e', 'g'],
                            ['h', 'k', 'l'],
                            ['h', 'k', 'm'],
                            ['j', 'k', 'l'],
                            ['j', 'k', 'm'],
                            ['n', 'o', 'p']])])
def test_split_rows_num_ok1(capsys, indata, direc, res):
    """Test OK cases for split_rows num."""
    ret = split_rows(indata=deepcopy(indata), directives=deepcopy(direc))
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('direc,err,msg',
                         [(2, TypeError, 'object is not iterable'),
                          ('not list', TypeError,
                           'string indices must be integers, not'),
                          (['not dict'], TypeError,
                           'string indices must be integers, not '),
                          ([{'colum': 1, 'separators': ['a', 'b'],
                             'not_separators': ['aa', 'bb']}],
                           KeyError, "KeyError('column')"),
                          ([{'column': 2, 'separator': ['a', 'b'],
                             'not_separators': ['aa', 'bb']}],
                           KeyError, "KeyError('separators')"),
                          ([{'column': 3, 'separators': ['a', 'b'],
                             'not_separator': ['aa', 'bb']}],
                           KeyError, "KeyError('not_separators')"),
                          ([{'column': 'a', 'separators': ['a', 'b'],
                             'not_separators': ['aa', 'bb']}],
                           SystemExit,
                           'Value for key column expected to be of type int'),
                          ([{'column': 4, 'separators': ['a', 'b'],
                             'not_separators': 'bb'}],
                           AssertionError, 'ExceptionInfo AssertionError'),
                          ([{'column': 5, 'separators': 'a',
                             'not_separators': ['aa', 'bb']}],
                           AssertionError, 'ExceptionInfo AssertionError')])
@pytest.mark.skip
def test_split_rows_num_nok1(capsys, direc, err, msg):
    """Test not OK cases for split_rows num."""
    indata = [['b', 'd']]
    with pytest.raises(err) as exc:
        _ = split_rows(indata=deepcopy(indata), directives=direc)
    out, err = capsys.readouterr()
    assert msg in str(exc)
    assert '' == out


@pytest.mark.parametrize('indata,direc,res',
                         [([], [], []),
                          ([['b', 'b']], [], [['b', 'b']]),
                          ([], [{'column': 1,
                                 'separators': [' ', ';'],
                                 'not_separators': ['\\;', '  ']}], []),
                          ([['b', 'd e', 'g'],
                            ['h+j', 'k', 'l;m'],
                            ['n', 'o', 'p']],
                           [{'column': 1, 'separators': [' ', 'x'],
                             'not_separators': ['   ']},
                            {'column': 0, 'separators': ['+', 'x'],
                             'not_separators': ['xx', 'xy']},
                            {'column': 2, 'separators': [';', 'x'],
                             'not_separators': ['\\;']}],
                           [['b', 'd', 'g'],
                            ['b', 'e', 'g'],
                            ['h', 'k', 'l'],
                            ['h', 'k', 'm'],
                            ['j', 'k', 'l'],
                            ['j', 'k', 'm'],
                            ['n', 'o', 'p']])])
def test_split_rows_numcfg_ok1(capsys, indata, direc, res):
    """Test OK cases for split_rows numcfg."""
    cfg1 = ConfigXlsListTransfNum()
    cfg1.s01_split_rows = direc
    jsontxt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfNum(from_json_text=jsontxt)
    ret = split_rows_cfg(indata=deepcopy(indata), cfg=cfg2, tinfo=1)
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
def test_merge_strings2(capsys, inlst, sep, res):
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
                          ([['x', 'y', 'z'],
                            [1, 2, 3]],
                           [], '+',
                           [['x', 'y', 'z'],
                            [1, 2, 3]]),
                          ([['x', 'y', 'z'],
                            [1, 2, 3]],
                           [[]], '+',
                           [['x', 'y', 'z'],
                            [1, 2, 3]]),
                          ([['x', 'y', 'z'],
                            [1, 2, 3]],
                           [[1]], '+',
                           [['x', 'y', 'z'],
                            [1, 2, 3]]),
                          ([['x', 'y', 'z'],
                            [1, 2, 3],
                            ['x1', 'y', 'z1'],
                            [4, 5, 6]],
                           [[0, 2, 3]], '+',
                           [['x+x1+4', 'y+5', 'z+z1+6'],
                            [1, 2, 3]]),
                          ([['x', 'y', 'z'],
                            [1, 2, 3],
                            ['x1', 'y', 'z1'],
                            [4, 5, 6]],
                           [[2, 0]], '-',
                           [[1, 2, 3],
                            ['x1-x', 'y', 'z1-z'],
                            [4, 5, 6]]),
                          ([['x', 'y', 'z'],
                            [1, 2, 3],
                            ['x1', 'y', 'z1'],
                            [4, 5, 6]],
                           [[0, 2], [1]], '-',
                           [['x-x1', 'y', 'z-z1'],
                            [1, 2, 3],
                            [4, 5, 6]]),
                          ([['x', 'y', 'z'],
                            [1, 2, 3],
                            ['x1', 'y', 'z1'],
                            [4, 5, 6]],
                           [[0, 2], [1, 3]], '-',
                           [['x-x1', 'y', 'z-z1'],
                            ['1-4', '2-5', '3-6']])])
def test_merge_ident_rows_num(capsys, data, rowidxs, sep, res):
    """Test OK cases of merge_identified_rows."""
    ret = merge_identified_rows(rows=deepcopy(data),
                                row_numbers=deepcopy(rowidxs),
                                separator=sep, tinfo=1)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


DATA1 = [['aa', 'bb', 3],
         ['ff', 'bc', 5],
         ['aa', 'bd', 3],
         ['ff', 'be', 5],
         ['gg', 'bf', 5],
         ['aa', 'bg', 7],
         ['aa', 'bh', 3]]


@pytest.mark.parametrize('data,cols,res',
                         [(DATA1, [0, 2],
                           [[0, 2, 6], [1, 3]]),
                          (DATA1, [1], []),
                          (DATA1, [],
                           [[0, 1, 2, 3, 4, 5, 6]]),
                          (DATA1, [2],
                           [[0, 2, 6], [1, 3, 4]]),
                          ([['b', 'a']], [0], [])])
def test_iden_rows_to_merge1nu(capsys, data, cols, res):
    """Test OK cases of merge identified rows."""
    ret = identify_rows_to_merge(rows=deepcopy(data),
                                 columns_to_cmp=deepcopy(cols),
                                 tinfo=1)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


DATA1ACP = [['aa', 'bb+bd+bh', 3],
            ['ff', 'bc+be', 5],
            ['gg', 'bf', 5],
            ['aa', 'bg', 7]]

DATA2ACCPM = [['aa', 'bb+bd+bh', 3],
              ['ff-gg', 'bc+be-bf', 5],
              ['aa', 'bg', 7]]

DATA1NP = [['aa+ff+gg', 'bb+bc+bd+be+bf+bg+bh', '3+5+7']]

DATA1CP = [['aa', 'bb+bd+bh', 3],
           ['ff+gg', 'bc+be+bf', 5],
           ['aa', 'bg', 7]]

DATA1ACS = [['aa', 'bb bd bh', 3],
            ['ff', 'bc be', 5],
            ['gg', 'bf', 5],
            ['aa', 'bg', 7]]

DATA1NS = [['aa ff gg', 'bb bc bd be bf bg bh', '3 5 7']]

DATA1CS = [['aa', 'bb bd bh', 3],
           ['ff gg', 'bc be bf', 5],
           ['aa', 'bg', 7]]


@pytest.mark.parametrize('data,cols,sep,res',
                         [(DATA1, [0, 2], '+',
                           DATA1ACP),
                          (DATA1, [1], '+', DATA1),
                          (DATA1, [], '+', DATA1NP),
                          (DATA1, [2], '+', DATA1CP),
                          (DATA1, [0, 2], ' ',
                           DATA1ACS),
                          (DATA1, [1], ' ', DATA1),
                          (DATA1, [], ' ', DATA1NS),
                          (DATA1, [2], ' ', DATA1CS)])
def test_one_merge_rows_num_ok1(capsys, data, cols, sep, res):
    """Test OK cases of one merge rows num."""
    ret = one_merge_rows(indata=deepcopy(data),
                         columns_to_cmp=deepcopy(cols),
                         separator=deepcopy(sep), tinfo=1)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
@pytest.mark.parametrize('data,cols,sep,res',
                         [(DATA1, [0, 2], '+',
                           DATA1ACP),
                          (DATA1, [1], '+', DATA1),
                          (DATA1, [], '+', DATA1NP),
                          (DATA1, [3], '+', DATA1CP),
                          (DATA1, [0, 2], ' ',
                           DATA1ACS),
                          (DATA1, [1], ' ', DATA1),
                          (DATA1, [], ' ', DATA1NS),
                          (DATA1, [2], ' ', DATA1CS)])
def test_one_rule_merge_rows_nu_ok1(capsys, data, cols, sep, res):
    """Test OK cases of one merge rows num."""
    rule = {'columns': deepcopy(cols), 'separator': deepcopy(sep)}
    ret = one_rule_merge_rows(indata=deepcopy(data), rule=rule,
                              tinfo=1)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
@pytest.mark.parametrize('data,cols,sep,res',
                         [(DATA1, [0, 2], '+',
                           DATA1ACP),
                          (DATA1, [1], '+', DATA1),
                          (DATA1, [2], '+', DATA1CP),
                          (DATA1, [0, 2], ' ',
                           DATA1ACS),
                          (DATA1, [1], ' ', DATA1),
                          (DATA1, [2], ' ', DATA1CS)])
def test_merge_rows_num_ok1(capsys, data, cols, sep, res):
    """Test OK cases of merge rows num."""
    rule = {'columns': deepcopy(cols), 'separator': deepcopy(sep)}
    cfg1 = ConfigXlsListTransfNum()
    cfg1.s02_merge_rows = [rule]
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfNum(from_json_text=txt)
    ret = merge_rows(indata=deepcopy(data), rules=cfg2.s02_merge_rows,
                     tinfo=1)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
@pytest.mark.parametrize('data,rule,res',
                         [(DATA1, [{'columns': [0, 2],
                                    'separator': '+'},
                                   {'columns': [3],
                                    'separator': '-'}],
                           DATA2ACCPM)])
def test_merge_rows_num_ok2(capsys, data, rule, res):
    """Test OK cases of merge rows num more than one rule."""
    cfg1 = ConfigXlsListTransfNum()
    cfg1.s02_merge_rows = deepcopy(rule)
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfNum(from_json_text=txt)
    ret = merge_rows(indata=deepcopy(data), rules=cfg2.s02_merge_rows,
                     tinfo=1)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
@pytest.mark.parametrize('data,cols,sep,res',
                         [(DATA1, [0, 2], '+',
                           DATA1ACP),
                          (DATA1, [1], '+', DATA1),
                          (DATA1, [2], '+', DATA1CP),
                          (DATA1, [0, 2], ' ',
                           DATA1ACS),
                          (DATA1, [1], ' ', DATA1),
                          (DATA1, [2], ' ', DATA1CS)])
def test_merge_rows_num_ok3(capsys, data, cols, sep, res):
    """Test OK cases of merge rows num."""
    rule = {'columns': deepcopy(cols), 'separator': deepcopy(sep)}
    cfg1 = ConfigXlsListTransfNum()
    cfg1.s02_merge_rows = [rule]
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfNum(from_json_text=txt)
    ret = merge_rows_cfg(indata=deepcopy(data), cfg=cfg2, tinfo=1)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
@pytest.mark.parametrize('data,rule,res',
                         [(DATA1, [{'columns': [0, 2],
                                    'separator': '+'},
                                   {'columns': ['2'],
                                    'separator': '-'}],
                           DATA2ACCPM)])
def test_merge_rows_num_ok4(capsys, data, rule, res):
    """Test OK cases of merge rows num more than one rule."""
    cfg1 = ConfigXlsListTransfNum()
    cfg1.s02_merge_rows = deepcopy(rule)
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfNum(from_json_text=txt)
    ret = merge_rows_cfg(indata=deepcopy(data), cfg=cfg2, tinfo=1)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err
