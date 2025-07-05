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
    merge_identified_rows_name, identify_rows_to_merge_name, \
    one_merge_rows_name, one_rule_merge_rows_name, merge_rows_name, \
    merge_rows_namecfg
from excel_list_transform.config_xls_list_transf_num import \
   ConfigXlsListTransfNum
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName


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
    """Test OK cases of one_split_one_row_name."""
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
    """Test OK cases of one_split_name."""
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
                            'but no such column in data'])])
def test_one_split_num_nok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
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
    """Test OK cases for split_rows_name."""
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
def test_slit_rows_num_nok1(capsys, direc, err, msg):
    """Test not OK cases for split_rows_name."""
    indata = [{'a': 'b', 'c': 'd'}]
    with pytest.raises(err) as exc:
        _ = split_rows(indata=deepcopy(indata), directives=direc)
    out, err = capsys.readouterr()
    assert msg in str(exc)
    assert '' == out


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
def test_split_rows_numcfg_ok1(capsys, indata, direc, res):
    """Test OK cases for split_rows_namecfg."""
    cfg1 = ConfigXlsListTransfNum()
    cfg1.s01_split_rows = direc
    jsontxt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfNum(from_json_text=jsontxt)
    ret = split_rows_cfg(indata=deepcopy(indata), cfg=cfg2, tinfo=1)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
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


@pytest.mark.skip
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
    """Test OK cases of merge_identified_rows_name."""
    ret = merge_identified_rows_name(rows=deepcopy(data),
                                     row_numbers=deepcopy(rowidxs),
                                     separator=sep)
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


@pytest.mark.skip
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
    ret = identify_rows_to_merge_name(rows=deepcopy(data),
                                      columns_to_cmp=deepcopy(cols))
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


@pytest.mark.skip
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
    ret = one_merge_rows_name(indata=deepcopy(data),
                              columns_to_cmp=deepcopy(cols),
                              separator=deepcopy(sep))
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
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
    ret = one_rule_merge_rows_name(indata=deepcopy(data), rule=rule)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
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
    ret = merge_rows_name(indata=deepcopy(data), rules=cfg2.s02_merge_rows)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
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
    ret = merge_rows_name(indata=deepcopy(data), rules=cfg2.s02_merge_rows)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
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
    ret = merge_rows_namecfg(indata=deepcopy(data), cfg=cfg2)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err


@pytest.mark.skip
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
    ret = merge_rows_namecfg(indata=deepcopy(data), cfg=cfg2)
    out, err = capsys.readouterr()
    assert ret == res
    assert '' == out
    assert '' == err
