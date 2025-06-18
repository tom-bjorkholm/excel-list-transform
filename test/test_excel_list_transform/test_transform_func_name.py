#! /usr/local/bin/python3
"""Test the excel_list_transform functions functionality."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from copy import deepcopy
from collections import namedtuple
from tempfile import TemporaryDirectory
import pytest
from excel_list_transform.transform_func_named import \
    rename_columns_name, \
    insert_columns_name, fix_indata_empty_rows_name, \
    check_indata_name, transform_data_name
from excel_list_transform.transform_func import transform_named_files
from excel_list_transform.config_excel_list_transform import \
    SplitWhere, FileType
from excel_list_transform.handle_excel import read_excel_num, write_excel_named
from excel_list_transform.handle_csv import read_csv_num, write_csv_named
from excel_list_transform.config_xls_list_refmt_name import \
    ConfigXlsListRefmtName
from excel_list_transform.commontypes import NameData
from excel_list_transform.config_enums import RewriteKind, CaseSensitivity
from excel_list_transform.transform_func_common import col_must_exist_name, \
    store_col_split_name, split_columns, merge_columns, rewrite_columns


@pytest.mark.parametrize('col, row, par',
                         [('a', {'x': 'b', 'a': 'c', 'y': 'e'}, 't1'),
                          ('y', {'x': 'b', 'a': 'c', 'y': 'e'}, 't2'),
                          ('x', {'x': 'b', 'a': 'c', 'y': 'e'}, 't3')])
def test_col_must_exist_name_ok(capsys, col, row, par):
    """Test OK cases of col_must_exist_name."""
    col_must_exist_name(col=col, row=row, param=par)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('col, row, par, msg',
                         [('b', {'x': 'b', 'a': 'c', 'y': 'e'}, 't1',
                           't1: no column named "b" in data row.'),
                          ('y', {}, 't2',
                           't2: no column named "y" in data row.'),
                          ('e', {'x': 'b', 'a': 'c', 'y': 'e'}, 't3',
                           't3: no column named "e" in data row.')])
def test_col_must_exist_name_nok(capsys, col, row, par, msg):
    """Test not OK cases of col_must_exist_name."""
    with pytest.raises(SystemExit):
        col_must_exist_name(col=col, row=row, param=par)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('row, col, val, sr, nrow',
                         [({'a': '1', 'b': '2'}, 'c', ['4', '5'],
                           {'column': 'c', 'right_name': 'd'},
                           {'a': '1', 'b': '2', 'c': '4', 'd': '5'}),
                          ({'a': '1', 'b': '2'}, 'c', ['4'],
                           {'column': 'c', 'right_name': 'd'},
                           {'a': '1', 'b': '2', 'c': '4', 'd': None}),
                          ({'a': '1', 'b': '2'}, 'c', [],
                           {'column': 'c', 'right_name': 'd'},
                           {'a': '1', 'b': '2', 'c': None, 'd': None}),
                          ({'a': '1', 'b': '2'}, 'a', ['4', '5'],
                           {'column': 'a', 'right_name': 'b'},
                           {'a': '4', 'b': '5'}),
                          ])
def test_store_col_split_name(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments  # noqa: E501
                              row, col, val, sr, nrow):
    """Test ok cases for store_col_split_name."""
    inrow = deepcopy(row)
    store_col_split_name(row=inrow, colref=col, val=val, singlerule=sr)
    out, err = capsys.readouterr()
    assert dict(sorted(inrow.items())) == dict(sorted(nrow.items()))
    assert '' == out
    assert '' == err


def assert_data_is_equal(left: NameData, right: NameData) -> None:
    """Assert that left and right data is equal."""
    assert len(left) == len(right)
    for leftrow, rightrow in zip(left, right):
        assert dict(sorted(leftrow.items())) == dict(sorted(rightrow.items()))


@pytest.mark.parametrize('ind, split, exp',
                         [([{'abc': 'a b', 'd e': 'c d o',
                             'e': 'e+f', 'g h': 'g'},
                            {'abc': 'h i', 'd e': 'j',
                             'e': 'k+l+p', 'g h': 'm n'}],
                           [{'column': 'd e', 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'right_name': 'z y'},
                            {'column': 'e', 'separator': '+',
                             'where': SplitWhere.LEFTMOST,
                             'right_name': 'xx'}],
                           [{'abc': 'a b', 'd e': 'c d',
                             'z y': 'o', 'e': 'e', 'xx': 'f',
                             'g h': 'g'},
                            {'abc': 'h i', 'd e': 'j',
                             'z y': None, 'e': 'k', 'xx': 'l+p',
                             'g h': 'm n'}]),
                          ([{'abc': 'a b', 'd e': 'c d o',
                             'e': 'e+f', 'g h': 'g'},
                            {'abc': 'h i', 'd e': 'j',
                             'e': 'k+l+p', 'g h': 'm n'}],
                           [],
                           [{'abc': 'a b', 'd e': 'c d o',
                             'e': 'e+f', 'g h': 'g'},
                            {'abc': 'h i', 'd e': 'j',
                             'e': 'k+l+p', 'g h': 'm n'}]),
                          ([{'abc': None, 'd e': 'c d o',
                             'e': 'e+f', 'g h': 'g'},
                            {'abc': 'h i', 'd e': 'j',
                             'e': 'k+l+p', 'g h': 'm n'}],
                           [{'column': 'abc', 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'right_name': 'added'}],
                           [{'abc': None, 'added': None, 'd e': 'c d o',
                             'e': 'e+f', 'g h': 'g'},
                            {'abc': 'h', 'added': 'i', 'd e': 'j',
                             'e': 'k+l+p', 'g h': 'm n'}])])
def test_split_columns_name(capsys, ind, split, exp):
    """Test splitting of columns (column names)."""
    cfg = ConfigXlsListRefmtName()
    cfg.s03_split_columns = split
    ret = split_columns(indata=ind, cfg=cfg, tinfo='a')
    out, err = capsys.readouterr()
    assert_data_is_equal(ret, exp)
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, split, msg',
                         [([{'x': 'a b', 'y': 'c d o', 'z': 'e+f', 'q': 'g'},
                           {'x': 'h i', 'y': 'j', 'z': 'k+l+p', 'q': 'm n'}],
                           [{'column': 'foobar', 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'right_name': 'something'}],
                           'no column named "foobar" in data row'),
                          ([{'x': 'a b', 'y': 'c d o', 'z': 'e+f', 'q': 'g'},
                           {'x': 'h i', 'y': 'j', 'z': 8, 'q': 'm n'}],
                           [{'column': 'z', 'separator': '+',
                             'where': SplitWhere.RIGHTMOST,
                             'right_name': 'something'}],
                           'Column "z" has value of type int')])
def test_split_columns_nok_name(capsys, ind, split, msg):
    """Test not OK splitting of columns (column numbers)."""
    cfg = ConfigXlsListRefmtName()
    cfg.s03_split_columns = split
    with pytest.raises(SystemExit):
        _ = split_columns(indata=ind, cfg=cfg, tinfo='a')
    out, err = capsys.readouterr()
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, merg, exp',
                         [([{'a': 'a', 'b': 'b', 'c': 'c'},
                            {'a': 'd', 'b': 'e', 'c': 'f'}],
                           [],
                           [{'a': 'a', 'b': 'b', 'c': 'c'},
                            {'a': 'd', 'b': 'e', 'c': 'f'}]),
                          ([{'q': 'a', 'x': 'b', 'y': 'c', 'z': 'd'},
                            {'q': 'd', 'x': 'e', 'y': None, 'z': 'f'},
                            {'q': 'g', 'x': None, 'y': 'h', 'z': 'j'},
                            {'q': 'k', 'x': None, 'y': None, 'z': 'm'}],
                           [{'columns': ['x', 'y'], 'separator': 'ww'}],
                           [{'q': 'a', 'x': 'bwwc', 'z': 'd'},
                            {'q': 'd', 'x': 'e', 'z': 'f'},
                            {'q': 'g', 'x': 'h', 'z': 'j'},
                            {'q': 'k', 'x': None, 'z': 'm'}]),
                          ([{'q': 'a', 'x': 'b', 'y': 'c', 'z': 'd'},
                            {'q': 'd', 'x': 'e', 'y': None, 'z': 'f'},
                            {'q': 'g', 'x': None, 'y': 'h', 'z': 'j'},
                            {'q': 'k', 'x': None, 'y': None, 'z': 'm'}],
                           [{'columns': ['x', 'y', 'z'], 'separator': 'ww'}],
                           [{'q': 'a', 'x': 'bwwcwwd'},
                            {'q': 'd', 'x': 'ewwf'},
                            {'q': 'g', 'x': 'hwwj'},
                            {'q': 'k', 'x': 'm'}]),
                          ([{'q': 'a', 'x': 'b', 'y': 'c', 'z': 'd'},
                            {'q': 'd', 'x': 'e', 'y': None, 'z': 'f'},
                            {'q': 'g', 'x': None, 'y': 'h', 'z': 'j'},
                            {'q': 'k', 'x': None, 'y': None, 'z': 'm'}],
                           [{'columns': ['x', 'z', 'y'], 'separator': 'ww'}],
                           [{'q': 'a', 'x': 'bwwdwwc'},
                            {'q': 'd', 'x': 'ewwf'},
                            {'q': 'g', 'x': 'jwwh'},
                            {'q': 'k', 'x': 'm'}]),
                          ([{'x': 'b', 'y': 'c', 'z': 'd'},
                            {'x': 'e', 'y': None, 'z': 'f'},
                            {'x': None, 'y': 'h', 'z': 'j'},
                            {'x': None, 'y': None, 'z': 'm'}],
                           [{'columns': ['x', 'y'], 'separator': 'ww'}],
                           [{'x': 'bwwc', 'z': 'd'}, {'x': 'e', 'z': 'f'},
                            {'x': 'h', 'z': 'j'}, {'x': None, 'z': 'm'}]),
                          ([{'x': 'b', 'y': 'c', 'z': 'd'},
                            {'x': 'e', 'y': None, 'z': 'f'},
                            {'x': None, 'y': 'h', 'z': 'j'},
                            {'x': None, 'y': None, 'z': 'm'}],
                           [{'columns': ['y'], 'separator': 'ww'}],
                           [{'x': 'b', 'y': 'c', 'z': 'd'},
                            {'x': 'e', 'y': None, 'z': 'f'},
                            {'x': None, 'y': 'h', 'z': 'j'},
                            {'x': None, 'y': None, 'z': 'm'}])])
def test_merge_columns_ok_name(capsys, ind, merg, exp):
    """Test merging of columns with name ref."""
    cfg = ConfigXlsListRefmtName()
    cfg.s05_merge_columns = merg
    ret = merge_columns(indata=ind, cfg=cfg, tinfo='a')
    out, err = capsys.readouterr()
    assert_data_is_equal(ret, exp)
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, merg, msg',
                         [([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [{'columns': ['z', 'q'], 'separator': '--'}],
                           's05_merge_columns: no column named "q" in data')])
def test_merge_columns_nok_name(capsys, ind, merg, msg):
    """Test not OK merging of columns with number ref."""
    cfg = ConfigXlsListRefmtName()
    cfg.s05_merge_columns = merg
    with pytest.raises(SystemExit):
        _ = merge_columns(indata=ind, cfg=cfg, tinfo='a')
    out, err = capsys.readouterr()
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, nam, exp',
                         [([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [],
                           [{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}]),
                          ([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [{'column': 'y', 'name': 'One'},
                            {'column': 'z', 'name': 'Zwei'}],
                           [{'x': 'a', 'One': 'b', 'Zwei': 'c'},
                            {'x': 'd', 'One': 'e', 'Zwei': 'f'}])])
def test_rename_columns_ok_name(capsys, ind, nam, exp):
    """Test ok renaming of columns."""
    cfg = ConfigXlsListRefmtName()
    cfg.s07_rename_columns = nam
    ret = rename_columns_name(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert ret == exp
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('nam, msg',
                         [([{'column': 'y', 'name': 'One'},
                            {'column': 'q', 'name': 'Zwei'}],
                           's07_rename_columns: no column named "q" in data')])
@pytest.mark.parametrize('ind',
                         [[{'x': 'a', 'y': 'b', 'z': 'c'},
                           {'x': 'd', 'y': 'e', 'z': 'f'}],
                          [{'x': 'a', 'y': 'b', 'z': 'c'},
                           {'x': 'd', 'y': 'e', 'z': 'f'}]])
def test_rename_columns_nok_name(capsys, ind, nam, msg):
    """Test nok renaming of columns."""
    cfg = ConfigXlsListRefmtName()
    cfg.s07_rename_columns = nam
    with pytest.raises(SystemExit):
        _ = rename_columns_name(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, ins, exp',
                         [([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [],
                           [{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}]),
                          ([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [{'column': 'p', 'value': None},
                            {'column': 'q', 'value': 'text'}],
                           [{'p': None, 'q': 'text',
                             'x': 'a', 'y': 'b', 'z': 'c'},
                            {'p': None, 'q': 'text',
                             'x': 'd', 'y': 'e', 'z': 'f'}])])
def test_insert_columns_ok_name(capsys, ind, ins, exp):
    """Test ok insertion of columns (name refs)."""
    cfg = ConfigXlsListRefmtName()
    cfg.s08_insert_columns = ins
    ret = insert_columns_name(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert_data_is_equal(exp, ret)
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ins, msg',
                         [([{'column': 'p', 'name': 'One', 'value': None},
                            {'column': 'x', 'name': 'Zwei', 'value': 'text'}],
                           's08_insert_columns: column "x" already exists')])
@pytest.mark.parametrize('ind',
                         [[{'x': 'a', 'y': 'b', 'z': 'c'},
                           {'x': 'd', 'y': 'e', 'z': 'f'}],
                          [{'x': 'a', 'y': 'b', 'z': 'c'},
                           {'x': 'd', 'y': 'e', 'z': 'f'}]])
def test_insert_columns_nok_name(capsys, ind, ins, msg):
    """Test nok inserting of columns (name refs)."""
    cfg = ConfigXlsListRefmtName()
    cfg.s08_insert_columns = ins
    with pytest.raises(SystemExit):
        _ = insert_columns_name(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, spec, exp',
                         [([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [],
                           [{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}]),
                          ([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [{'column': 'x', 'from': 'd', 'to': 'aba',
                             'kind': RewriteKind.STR_SUBSTITUTE,
                             'case': CaseSensitivity.MATCH_CASE},
                            {'column': 'x', 'from': '^a', 'to': 'x',
                             'kind': RewriteKind.REGEX_SUBSTITUTE,
                             'case': CaseSensitivity.IGNORE_CASE}],
                           [{'x': 'x', 'y': 'b', 'z': 'c'},
                            {'x': 'xba', 'y': 'e', 'z': 'f'}])])
def test_rewrite_columns_ok_name(capsys, ind, spec, exp):
    """Test ok insertion of columns (name refs)."""
    cfg = ConfigXlsListRefmtName()
    cfg.s09_rewrite_columns = spec
    ret = rewrite_columns(indata=ind, cfg=cfg, tinfo='a')
    out, err = capsys.readouterr()
    assert_data_is_equal(exp, ret)
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, spec',
                         [([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [{'column': 'x', 'from': 'd', 'to': 'aba',
                             'kind': RewriteKind.STR_SUBSTITUTE,
                             'case': CaseSensitivity.MATCH_CASE},
                            {'column': 'q', 'from': '^a', 'to': 'x',
                             'kind': RewriteKind.REGEX_SUBSTITUTE,
                             'case': CaseSensitivity.IGNORE_CASE}])])
def test_rewrite_columns_nok_name(capsys, ind, spec):
    """Test not ok insertion of columns (name refs)."""
    cfg = ConfigXlsListRefmtName()
    cfg.s09_rewrite_columns = spec
    with pytest.raises(SystemExit):
        _ = rewrite_columns(indata=ind, cfg=cfg, tinfo='a')
    out, err = capsys.readouterr()
    assert 's09_rewrite_columns: no column named "q" in data row' in err
    assert '' == out


@pytest.mark.parametrize('ind, outd',
                         [([{'x': 'a', 'y': 'b', 'z': 'c'},
                           {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [{'x': 'a', 'y': 'b', 'z': 'c'},
                           {'x': 'd', 'y': 'e', 'z': 'f'}]),
                          ([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {}],
                           [{'x': 'a', 'y': 'b', 'z': 'c'}]),
                          ([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {}, {}],
                           [{'x': 'a', 'y': 'b', 'z': 'c'}]),
                          ([{'x': 3, 'y': 7, 'z': 9},
                            {}, {}],
                           [{'x': 3, 'y': 7, 'z': 9}]),
                          ([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': None, 'y': '', 'z': ''}],
                           [{'x': 'a', 'y': 'b', 'z': 'c'}]),
                          ([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': None, 'y': '', 'z': ''},
                            {'x': None, 'y': '', 'z': ''}, {}],
                           [{'x': 'a', 'y': 'b', 'z': 'c'}])])
def test_fix_empty_rows_name_ok(capsys, ind, outd):
    """Test OK cases of fix_indata_empty_rows_num."""
    fix_indata_empty_rows_name(indata=ind)
    out, err = capsys.readouterr()  # pylint: disable=duplicate-code  # noqa: E501
    assert ind == outd
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, msg',   # pylint: disable=duplicate-code  # noqa: E501
                         [([{'x': 'd', 'y': 'e', 'z': 'f'},
                            ['a', 'b', 'c']],
                           'Expected dict of columns but got list'),
                          ([{'x': 'd', 'y': 'e', 'z': 'f'},
                            'e'],
                           'Expected dict of columns but got str'),
                          ([{'x': 'd', 'y': 'e', 'z': 'f'},
                            42],
                           'Expected dict of columns but got int')])
def test_fix_empty_rows_name_nok(capsys, ind, msg):
    """Test OK cases of fix_indata_empty_rows_num."""
    with pytest.raises(TypeError) as exc:
        fix_indata_empty_rows_name(indata=ind)
    out, err = capsys.readouterr()
    assert msg in str(exc)
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, res',
                         [([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}],
                           [{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'z': 'f'}]),
                          ([{'x': 'a', 'y': 'b'},
                            {'x': 'c', 'y': 'd'},
                            {'x': 'e', 'y': 'f'}],
                           [{'x': 'a', 'y': 'b'},
                            {'x': 'c', 'y': 'd'},
                            {'x': 'e', 'y': 'f'}]),
                          ([{'x': 'a', 'y': 'b'}, {}],
                           [{'x': 'a', 'y': 'b'}]),
                          ([{'x': 'a', 'y': 'b'}, {'x': None}],
                           [{'x': 'a', 'y': 'b'}]),
                          ([{'x': 'a', 'y': 'b'}, {'x': ''}],
                           [{'x': 'a', 'y': 'b'}])])
def test_check_indata_ok_name(capsys, ind, res):
    """Test check_indata for OK case with named refs."""
    indata = deepcopy(ind)
    check_indata_name(indata=indata)
    out, err = capsys.readouterr()
    assert indata == res
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, exc, msg, excmsg',
                         [([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'e', 'y': 'f'}], SystemExit,
                           'Number of columns different between lines', None),
                          ('data', AssertionError, None,
                           'Expected list of rows but got'),
                          ([], SystemExit, 'No rows in input data', None),
                          (['data'], TypeError,
                           'Expected dict of columns but got', None),
                          ([{'x': 'a', 'y': 'b'},
                            {'x': 'e'}], SystemExit,
                           'Number of columns different between lines', None),
                          ([{}, {}], SystemExit,
                           'No columns in input data', None),
                          ([{'x': 'a', 'y': 'b', 'z': 'c'},
                            {'x': 'd', 'y': 'e', 'q': 'f'}],
                           RuntimeError, 'Columns names different between ' +
                           "lines. Found ['x', 'y', 'z'] and " +
                           "['q', 'x', 'y']. Aborting.", None)])
def test_check_indata_nok_name(capsys, ind, exc, msg, excmsg):
    """Test check_indata for nok case with named refs."""
    indata = deepcopy(ind)
    with pytest.raises(exc) as exc_inst:
        check_indata_name(indata=indata)
    out, err = capsys.readouterr()
    if msg is not None:
        assert msg in err
    if excmsg is not None:
        assert excmsg in str(exc_inst)
    assert '' == out


DataToUseName = namedtuple('DataToUseName',
                           ['indata', 'split_cols',
                            'merge_cols', 'first_cols', 'rename_cols',
                            'insert_cols', 'column_order', 'result',
                            'rewrite_cols', 'in_col_order'])


def get_test_data_name(written_result: bool) -> DataToUseName:
    """Get a test data set (name refs)."""
    result = []
    if written_result:
        result = [['q', 'the new', 'Second'],
                  ['h', 'x', 'då'],
                  [None, 'x', 'gÅÄÖåäö']]
    else:
        result = [{'q': 'h', 'the new': 'x', 'Second': 'då', 'x': 'a b'},
                  {'q': None, 'the new': 'x', 'Second': 'gÅÄÖåäö', 'x': 'e f'}]
    ret = DataToUseName(indata=[{'x': 'a', 'y': 'b c', 'z': 'då'},
                                {'x': 'e', 'y': 'f', 'z': 'gÅÄÖåäö'}],
                        split_cols=[{'column': 'y', 'separator': ' ',
                                     'where': SplitWhere.RIGHTMOST,
                                     'right_name': 'q'}],
                        merge_cols=[{'columns': ['x', 'y'], 'separator': ' '}],
                        first_cols=None,
                        rename_cols=[{'column': 'z', 'name': 'Second'}],
                        insert_cols=[{'column': 'the new',
                                      'value': 'x'}],
                        column_order=['q', 'the new', 'Second'],
                        rewrite_cols=[{'kind': RewriteKind.STR_SUBSTITUTE,
                                       'column': 'q',
                                       'from': 'c', 'to': 'h',
                                       'case': CaseSensitivity.MATCH_CASE}],
                        result=result,
                        in_col_order=['x', 'y', 'z'])
    return ret


@pytest.mark.smoke
def test_transform_data_ok_name(capsys):
    """Test transform_data with OK input (num refs)."""
    cfg = ConfigXlsListRefmtName()
    test_data = get_test_data_name(written_result=False)
    cfg.s03_split_columns = test_data.split_cols
    cfg.s05_merge_columns = test_data.merge_cols
    cfg.s07_rename_columns = test_data.rename_cols
    cfg.s08_insert_columns = test_data.insert_cols
    cfg.s09_rewrite_columns = test_data.rewrite_cols
    res = transform_data_name(indata=test_data.indata, cfg=cfg)
    out, err = capsys.readouterr()
    assert '' == err
    assert '' == out
    assert_data_is_equal(res, test_data.result)


@pytest.mark.parametrize('enc', ['utf-8', 'iso8859-1'])
def test_rfmt_nmd_files_xl2cs_name(capsys, enc):
    """Test transform_name_files from xlsx to csv (named refs)."""
    cfg = ConfigXlsListRefmtName()
    cfg.out_csv_encoding = enc
    test_data = get_test_data_name(written_result=True)
    cfg.s03_split_columns = test_data.split_cols
    cfg.s05_merge_columns = test_data.merge_cols
    cfg.s07_rename_columns = test_data.rename_cols
    cfg.s08_insert_columns = test_data.insert_cols
    cfg.s09_rewrite_columns = test_data.rewrite_cols
    cfg.s10_column_order = test_data.column_order
    cfg.in_type = FileType.EXCEL  # pylint: disable=duplicate-code  # noqa: E501
    cfg.out_type = FileType.CSV
    with TemporaryDirectory() as dirname:
        cfgname = dirname + '/test.cfg'  # pylint: disable=duplicate-code  # noqa: E501
        cfg.write(cfgname)
        infilename = dirname + '/in'
        outfilename = dirname + '/out'
        write_excel_named(data=test_data.indata,
                          filename=infilename + '.xlsx',
                          column_order=test_data.in_col_order)
        transform_named_files(infilename=infilename, outfilename=outfilename,  # pylint: disable=duplicate-code  # noqa: E501
                              cfgfilename=cfgname)
        res = read_csv_num(filename=outfilename + '.csv',
                           dialect=cfg.get_out_csv_dialect(),
                           encoding=cfg.out_csv_encoding,
                           max_column_read=20)
        out, err = capsys.readouterr()
        assert '' == err
        assert f'Wrote {outfilename}.csv' == out.strip()
        assert res == test_data.result


@pytest.mark.parametrize('enc', ['utf-8', 'iso8859-1'])
def test_rfmt_nmd_files_cs2xl_name(capsys, enc):
    """Test transform_name_files from csv to xlsx (name refs)."""
    cfg = ConfigXlsListRefmtName()
    cfg.in_csv_encoding = enc
    test_data = get_test_data_name(written_result=True)
    cfg.s03_split_columns = test_data.split_cols
    cfg.s05_merge_columns = test_data.merge_cols
    cfg.s07_rename_columns = test_data.rename_cols
    cfg.s08_insert_columns = test_data.insert_cols
    cfg.s09_rewrite_columns = test_data.rewrite_cols
    cfg.s10_column_order = test_data.column_order
    cfg.out_type = FileType.EXCEL  # pylint: disable=duplicate-code  # noqa: E501
    cfg.in_type = FileType.CSV
    with TemporaryDirectory() as dirname:
        cfgname = dirname + '/test.cfg'  # pylint: disable=duplicate-code  # noqa: E501
        cfg.write(cfgname)
        infilename = dirname + '/in'
        outfilename = dirname + '/out'
        write_csv_named(data=test_data.indata, filename=infilename + '.csv',
                        dialect=cfg.get_in_csv_dialect(),
                        encoding=cfg.in_csv_encoding,
                        column_order=test_data.in_col_order)
        transform_named_files(infilename=infilename, outfilename=outfilename,  # pylint: disable=duplicate-code  # noqa: E501
                              cfgfilename=cfgname)
        res = read_excel_num(filename=outfilename + '.xlsx',
                             max_column_read=20, strip_col_names=False,
                             strip_values=False)
        out, err = capsys.readouterr()
        assert '' == err
        assert f'Wrote {outfilename}.xlsx' == out.strip()
        assert res == test_data.result
