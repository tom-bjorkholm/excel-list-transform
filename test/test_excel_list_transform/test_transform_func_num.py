#! /usr/local/bin/python3
"""Test the excel_list_transform functions functionality."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from typing import NamedTuple, Optional, cast
from copy import deepcopy
from tempfile import TemporaryDirectory
import pytest
from pytest import CaptureFixture
from test_excel_list_transform.tableio_helpers import \
    configure_input_csv, configure_input_excel, configure_output_csv, \
    configure_output_excel, read_excel_num, write_csv_num, write_excel_num
from excel_list_transform.config_enums import RewriteKind, CaseSensitivity, \
    SplitWhere
from excel_list_transform.transform_func_num import \
    remove_columns_num, place_columns_first, \
    rename_columns_num, insert_columns_num, \
    check_indata_num, transform_data_num, fix_indata_empty_rows_num
from excel_list_transform.transform_func import transform_named_files
from excel_list_transform.handle_tableio import read_table_num
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.config_excel_list_transform import Rule, RuleMerge, \
    RulePlace, RuleRemove, RuleRewrite, RuleSplit, SingleRuleSplit
from excel_list_transform.commontypes import NumData, NumRow
from excel_list_transform.transform_func_common import col_must_exist_num, \
    store_col_split_num, split_columns, merge_columns, rewrite_columns


def read_csv_table(filename: str, cfg: ConfigXlsListTransfNum) -> NumData:
    """Read TableIO CSV output as numbered rows."""
    read_cfg = ConfigXlsListTransfNum()
    read_cfg.input_table.format_name = 'CSV'
    read_cfg.input_table.character_encoding = \
        cfg.output_table.character_encoding
    read_cfg.input_table.csv = deepcopy(cfg.output_table.csv)
    read_cfg.max_column_read = 20
    return read_table_num(filename=filename, cfg=read_cfg)


@pytest.mark.parametrize('col, row, par',
                         [(0, ['a', 'b', 'c'], 't1'),
                          (0, ['a', 'b', 'c'], 't2'),
                          (0, ['a', 'b', 'c'], 't3')])
def test_col_must_exst_num_ok(capsys: CaptureFixture[str], col: int,
                              row: NumRow, par: str) -> None:
    """Test OK cases for col_must_exist_num."""
    col_must_exist_num(col=col, row=row, param=par)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('col, row, par, msg',
                         [(-1, ['a', 'b', 'c'], 't1',
                           't1: column index -1 out of range [0, 2].'),
                          (3, ['a', 'b', 'c'], 't2',
                           't2: column index 3 out of range [0, 2].'),
                          (17, ['a', 'b', 'c', 'd'], 't3',
                           't3: column index 17 out of range [0, 3].')])
def test_col_must_exst_num_no(capsys: CaptureFixture[str], col: int,
                              row: NumRow, par: str, msg: str) -> None:
    """Test not OK cases for col_must_exist_num."""
    with pytest.raises(SystemExit):
        col_must_exist_num(col=col, row=row, param=par)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('row, col, val, sr, res',
                         [(['a', 'b', 'c'], 1, ['x', 'z'],
                           {'column': 1,
                            'store_single': SplitWhere.RIGHTMOST},
                           ['a', 'x', 'z', 'b', 'c']),
                          (['a', 'b', 'c'], 1, ['x'],
                           {'column': 1,
                            'store_single': SplitWhere.RIGHTMOST},
                           ['a', None, 'x', 'b', 'c']),
                          (['a', 'b', 'c'], 1, ['x'],
                           {'column': 1,
                            'store_single': SplitWhere.LEFTMOST},
                           ['a', 'x', None, 'b', 'c']),
                          (['a', 'b', 'c'], 1, [],
                           {'column': 1,
                            'store_single': SplitWhere.RIGHTMOST},
                           ['a', None, None, 'b', 'c']),
                          (['a', 'b', 'c'], 0, ['x', 'z'],
                           {'column': 1,
                            'store_single': SplitWhere.RIGHTMOST},
                           ['x', 'z', 'a', 'b', 'c']),
                          (['a', 'b', 'c'], 3, ['x', 'z'],
                           {'column': 1,
                            'store_single': SplitWhere.RIGHTMOST},
                           ['a', 'b', 'c', 'x', 'z'])])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_store_col_split_num(capsys: CaptureFixture[str], row: NumRow,
                             col: int, val: list[str],
                             sr: SingleRuleSplit[int], res: NumRow) -> None:
    """Test OK cases of store_col_split_num."""
    inrow = deepcopy(row)
    store_col_split_num(row=inrow, colref=col, val=val, singlerule=sr)
    out, err = capsys.readouterr()
    assert inrow == res
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('ind, split, exp',
                         [([['a b', 'c d o', 'e+f', 'g']],
                           [{'column': 2, 'separator': '+',
                             'where': SplitWhere.LEFTMOST,
                             'store_single': SplitWhere.RIGHTMOST}],
                           [['a b', 'c d o', 'e', 'f', 'g']]),
                          ([['h i', 'j', 'k+l+p', 'm n']],
                           [{'column': 1, 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'store_single': SplitWhere.LEFTMOST}],
                           [['h i', 'j', None, 'k+l+p', 'm n']]),
                          ([['a b', 'c d o', 'e+f', 'g'],
                            ['h i', 'j', 'k+l+p', 'm n']],
                           [{'column': 1, 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'store_single': SplitWhere.LEFTMOST},
                            {'column': 2, 'separator': '+',
                             'where': SplitWhere.LEFTMOST,
                             'store_single': SplitWhere.RIGHTMOST}],
                           [['a b', 'c d', 'o', 'e', 'f', 'g'],
                            ['h i', 'j', None, 'k', 'l+p', 'm n']]),
                          ([['a b', 'c d o', 'e f', 'g'],
                            ['h i', 'j', 'k l p', 'm n']],
                           [{'column': 1, 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'store_single': SplitWhere.LEFTMOST}],
                           [['a b', 'c d', 'o', 'e f', 'g'],
                            ['h i', 'j', None, 'k l p', 'm n']]),
                          ([['a b', '', 'e f', 'g'],
                            ['h i', None, 'k l p', 'm n']],
                           [{'column': 1, 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'store_single': SplitWhere.LEFTMOST}],
                           [['a b', '', None, 'e f', 'g'],
                            ['h i', None, None, 'k l p', 'm n']]),
                          ([['a b', '', 'e f', 'g'],
                            ['h i', None, 'k l p', 'm n']],
                           [],
                           [['a b', '', 'e f', 'g'],
                            ['h i', None, 'k l p', 'm n']]),
                          ([['a b', 'c d o', 'e f', 'g'],
                            ['h i', 'j', 'k l p', 'm n']],
                           [{'column': 1, 'separator': ' ',
                             'where': SplitWhere.LEFTMOST,
                             'store_single': SplitWhere.RIGHTMOST}],
                           [['a b', 'c', 'd o', 'e f', 'g'],
                            ['h i', None, 'j', 'k l p', 'm n']])])
def test_split_columns_num(capsys: CaptureFixture[str], ind: NumData,
                           split: RuleSplit[int], exp: NumData) -> None:
    """Test OK splitting of columns (column numbers)."""
    cfg = ConfigXlsListTransfNum()
    cfg.s03_split_columns = split
    ret = split_columns(indata=ind, cfg=cfg, tinfo=2)
    out, err = capsys.readouterr()
    assert ret == exp
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, split, msg',
                         [([['a b', 'c d o', 'e+f', 'g'],
                            ['h i', 'j', 'k+l+p', 'm n']],
                           [{'column': -1, 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'store_single': SplitWhere.LEFTMOST}],
                           'column index -1 out of range [0, 3]'),
                          ([['a b', 'c d o', 'e+f'],
                            ['h i', 'j', 'k+l+p']],
                           [{'column': 10, 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'store_single': SplitWhere.LEFTMOST}],
                           'column index 10 out of range [0, 2]'),
                          ([['a b', 'c d o', 'e+f', 'g'],
                            ['h i', 'j', 'k+l+p', 'm n']],
                           [{'column': 4, 'separator': ' ',
                             'where': SplitWhere.RIGHTMOST,
                             'store_single': SplitWhere.LEFTMOST}],
                           'column index 4 out of range [0, 3]')])
def test_spl_cols_nok_num(capsys: CaptureFixture[str], ind: NumData,
                          split: RuleSplit[int], msg: str) -> None:
    """Test not OK splitting of columns (column numbers)."""
    cfg = ConfigXlsListTransfNum()
    cfg.s03_split_columns = split
    with pytest.raises(SystemExit):
        _ = split_columns(indata=ind, cfg=cfg, tinfo=2)
    out, err = capsys.readouterr()
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, rem, exp',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [],
                           [['a', 'b', 'c'], ['d', 'e', 'f']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [0],
                           [['b', 'c'], ['e', 'f']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [1],
                           [['a', 'c'], ['d', 'f']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [2],
                           [['a', 'b'], ['d', 'e']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [2],
                           [['a', 'b'], ['d', 'e']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [0, 2],
                           [['b'], ['e']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [0, 1],
                           [['c'], ['f']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [1, 2],
                           [['a'], ['d']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [0, 1, 2],
                           [[], []])])
def test_remove_columns_ok(capsys: CaptureFixture[str], ind: NumData,
                           rem: RuleRemove, exp: NumData) -> None:
    """Test OK cases of removal of columns."""
    cfg = ConfigXlsListTransfNum()
    cfg.s04_remove_columns = rem
    ret = remove_columns_num(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert ret == exp
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, rem, msg',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [3], 'column index 3 out of range [0, 2].'),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [-2], 'column index -2 out of range [0, 2].')])
def test_remove_columns_nok(capsys: CaptureFixture[str], ind: NumData,
                            rem: RuleRemove, msg: str) -> None:
    """Test not OK cases of removal of columns."""
    cfg = ConfigXlsListTransfNum()
    cfg.s04_remove_columns = rem
    with pytest.raises(SystemExit):
        _ = remove_columns_num(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, merg, exp',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [],
                           [['a', 'b', 'c'], ['d', 'e', 'f']]),
                          ([['a', 'b', 'c', 'd'],
                            ['d', 'e', None, 'f'],
                            ['g', None, 'h', 'j'],
                            ['k', None, None, 'm']],
                           [{'columns': [1, 2], 'separator': 'ww'}],
                           [['a', 'bwwc', 'd'], ['d', 'e', 'f'],
                            ['g', 'h', 'j'], ['k', None, 'm']]),
                          ([['b', 'c', 'd'],
                            ['e', None, 'f'],
                            [None, 'h', 'j'],
                            [None, None, 'm']],
                           [{'columns': [0, 1], 'separator': 'ww'}],
                           [['bwwc', 'd'], ['e', 'f'],
                            ['h', 'j'], [None, 'm']]),
                          ([['b', 'c', 'd'],
                            ['e', None, 'f'],
                            [None, 'h', 'j'],
                            [None, None, 'm']],
                           [{'columns': [1], 'separator': 'ww'}],
                           [['b', 'c', 'd'], ['e', None, 'f'],
                            [None, 'h', 'j'], [None, None, 'm']])])
def test_merge_columns_ok_num(capsys: CaptureFixture[str], ind: NumData,
                              merg: RuleMerge[int], exp: NumData) -> None:
    """Test merging of columns with number ref."""
    cfg = ConfigXlsListTransfNum()
    cfg.s05_merge_columns = merg
    ret = merge_columns(indata=ind, cfg=cfg, tinfo=2)
    out, err = capsys.readouterr()
    assert ret == exp
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, merg, msg',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [{'columns': [2, 3], 'separator': '--'}],
                           'column index 3 out of range [0, 2]'),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [{'columns': [-1, 0], 'separator': '--'}],
                           'column index -1 out of range [0, 2]')])
def test_mrg_cols_nok_num(capsys: CaptureFixture[str], ind: NumData,
                          merg: RuleMerge[int], msg: str) -> None:
    """Test not OK merging of columns with number ref."""
    cfg = ConfigXlsListTransfNum()
    cfg.s05_merge_columns = merg
    with pytest.raises(SystemExit):
        _ = merge_columns(indata=ind, cfg=cfg, tinfo=2)
    out, err = capsys.readouterr()
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, pla, exp',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [],
                           [['a', 'b', 'c'], ['d', 'e', 'f']]),
                          ([['a', 'b', 'c', 'x'], ['d', 'e', 'f', 'y']],
                           [2],
                           [['c', 'a', 'b', 'x'], ['f', 'd', 'e', 'y']]),
                          ([['a', 'b', 'c', 'x'], ['d', 'e', 'f', 'y']],
                           [2, 1],
                           [['c', 'b', 'a', 'x'], ['f', 'e', 'd', 'y']]),
                          ([['a', 'b', 'c', 'x'], ['d', 'e', 'f', 'y']],
                           [1, 2],
                           [['b', 'c', 'a', 'x'], ['e', 'f', 'd', 'y']]),
                          ([['a', 'b', 'c', 'x'], ['d', 'e', 'f', 'y']],
                           [2, 1, 0],
                           [['c', 'b', 'a', 'x'], ['f', 'e', 'd', 'y']]),
                          ([['a', 'b', 'c', 'x'], ['d', 'e', 'f', 'y']],
                           [1, 2, 0],
                           [['b', 'c', 'a', 'x'], ['e', 'f', 'd', 'y']])
                          ])
def test_place_cols_first_ok(capsys: CaptureFixture[str], ind: NumData,
                             pla: RulePlace, exp: NumData) -> None:
    """Test placind of columns first."""
    cfg = ConfigXlsListTransfNum()
    cfg.s06_place_columns_first = pla
    ret = place_columns_first(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert ret == exp
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('pla', [[-1], [4], [2, 4, 1]])
@pytest.mark.parametrize('ind', [([['a', 'b', 'c'], ['d', 'e', 'f']])])
def test_place_cols_first_nok(capsys: CaptureFixture[str], ind: NumData,
                              pla: RulePlace) -> None:
    """Test not OK placing of columns firs."""
    cfg = ConfigXlsListTransfNum()
    cfg.s06_place_columns_first = pla
    with pytest.raises(SystemExit):
        _ = place_columns_first(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert 'out of range' in err
    assert '' == out


@pytest.mark.parametrize('ind, nam, exp',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [],
                           [['a', 'b', 'c'], ['d', 'e', 'f']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [{'column': 1, 'name': 'One'},
                            {'column': 2, 'name': 'Zwei'}],
                           [['a', 'One', 'Zwei'], ['d', 'e', 'f']])])
def test_ren_cols_ok_num(capsys: CaptureFixture[str], ind: NumData,
                         nam: Rule[int], exp: NumData) -> None:
    """Test ok renaming of columns."""
    cfg = ConfigXlsListTransfNum()
    cfg.s07_rename_columns = nam
    ret = rename_columns_num(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert ret == exp
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('nam, msg',
                         [([{'column': 1, 'name': 'One'},
                            {'column': 20, 'name': 'Zwei'}],
                           's07_rename_columns: column index 20 ' +
                           'out of range [0, 2]'),
                          ([{'column': 1, 'name': 'One'},
                            {'column': -1, 'name': 'Zwei'}],
                           's07_rename_columns: column index -1 out ' +
                           'of range')])
@pytest.mark.parametrize('ind',
                         [[['a', 'b', 'c'], ['d', 'e', 'f']],
                          [['a', 'b', 'c'], ['d', 'e', 'f']]])
def test_ren_cols_nok_num(capsys: CaptureFixture[str], ind: NumData,
                          nam: Rule[int], msg: str) -> None:
    """Test nok renaming of columns."""
    cfg = ConfigXlsListTransfNum()
    cfg.s07_rename_columns = nam
    with pytest.raises(SystemExit):
        _ = rename_columns_num(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, ins, exp',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [],
                           [['a', 'b', 'c'], ['d', 'e', 'f']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [{'column': 1, 'name': 'One', 'value': None},
                            {'column': 4, 'name': 'Zwei', 'value': 'text'}],
                           [['a', 'One', 'b', 'c', 'Zwei'],
                            ['d', None, 'e', 'f', 'text']])])
def test_ins_cols_ok_num(capsys: CaptureFixture[str], ind: NumData,
                         ins: Rule[int], exp: NumData) -> None:
    """Test ok insertion of columns (number refs)."""
    cfg = ConfigXlsListTransfNum()
    cfg.s08_insert_columns = ins
    ret = insert_columns_num(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert ret == exp
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ins, msg',
                         [([{'column': 1, 'name': 'One', 'value': None},
                            {'column': 20, 'name': 'Zwei', 'value': 'text'}],
                           's08_insert_columns: column index 20 ' +
                           'out of range [0, 4].'),
                          ([{'column': 1, 'name': 'One', 'value': 'text'},
                            {'column': -1, 'name': 'Zwei', 'value': None}],
                           's08_insert_columns: column index -1 ' +
                           'out of range [0, 4].')])
@pytest.mark.parametrize('ind',
                         [[['a', 'b', 'c'], ['d', 'e', 'f']],
                          [['a', 'b', 'c'], ['d', 'e', 'f']]])
def test_ins_cols_nok_num(capsys: CaptureFixture[str], ind: NumData,
                          ins: Rule[int], msg: str) -> None:
    """Test nok inserting of columns (number refs)."""
    cfg = ConfigXlsListTransfNum()
    cfg.s08_insert_columns = ins
    with pytest.raises(SystemExit):
        _ = insert_columns_num(indata=ind, cfg=cfg)
    out, err = capsys.readouterr()
    assert msg in err
    assert '' == out


@pytest.mark.parametrize('ind, spec, exp',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [],
                           [['a', 'b', 'c'], ['d', 'e', 'f']]),
                          ([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [{'column': 0, 'from': 'd', 'to': 'aba',
                             'kind': RewriteKind.STR_SUBSTITUTE,
                             'case': CaseSensitivity.MATCH_CASE},
                            {'column': 0, 'from': '^a', 'to': 'x',
                             'kind': RewriteKind.REGEX_SUBSTITUTE,
                             'case': CaseSensitivity.IGNORE_CASE}],
                           [['x', 'b', 'c'], ['xba', 'e', 'f']])])
def test_rew_cols_ok_num(capsys: CaptureFixture[str], ind: NumData,
                         spec: RuleRewrite[int], exp: NumData) -> None:
    """Test ok insertion of columns (name refs)."""
    cfg = ConfigXlsListTransfNum()
    cfg.s09_rewrite_columns = spec
    ret = rewrite_columns(indata=ind, cfg=cfg, tinfo=2)
    out, err = capsys.readouterr()
    assert exp == ret
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, spec',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [{'column': 0, 'from': 'd', 'to': 'aba',
                             'kind': RewriteKind.STR_SUBSTITUTE,
                             'case': CaseSensitivity.MATCH_CASE},
                            {'column': 5, 'from': '^a', 'to': 'x',
                             'kind': RewriteKind.REGEX_SUBSTITUTE,
                             'case': CaseSensitivity.IGNORE_CASE}])])
def test_rew_cols_nok_num(capsys: CaptureFixture[str], ind: NumData,
                          spec: RuleRewrite[int]) -> None:
    """Test not ok insertion of columns (name refs)."""
    cfg = ConfigXlsListTransfNum()
    cfg.s09_rewrite_columns = spec
    with pytest.raises(SystemExit):
        _ = rewrite_columns(indata=ind, cfg=cfg, tinfo=2)
    out, err = capsys.readouterr()
    assert 's09_rewrite_columns: column index 5 out of range' in err
    assert '' == out


@pytest.mark.parametrize('ind, outd',
                         [([['a', 'b', 'c'],
                            ['d', 'e', 'f']],
                           [['a', 'b', 'c'],
                            ['d', 'e', 'f']]),
                          ([['a', 'b', 'c'],
                            []],
                           [['a', 'b', 'c']]),
                          ([['a', 'b', 'c'],
                            [], []],
                           [['a', 'b', 'c']]),
                          ([['a', 'b', 'c'],
                            ['']],
                           [['a', 'b', 'c']]),
                          ([['a', 'b', 'c'],
                            [''], [''], []],
                           [['a', 'b', 'c']])])
# pylint: disable-next=duplicate-code
def test_fix_empty_rows_num_o(capsys: CaptureFixture[str], ind: NumData,
                              outd: NumData) -> None:
    """Test OK cases of fix_indata_empty_rows_num."""
    # pylint: disable-next=duplicate-code
    fix_indata_empty_rows_num(indata=ind)
    out, err = capsys.readouterr()
    assert ind == outd
    assert '' == err
    assert '' == out


# pylint: disable-next=duplicate-code
@pytest.mark.parametrize('ind, msg',
                         [([['a', 'b', 'c'],
                            {'q': 'd', 'x': 'e'}],
                           'Expected list of columns but got dict'),
                          ([['a', 'b', 'c'],
                            'e'],
                           'Expected list of columns but got str'),
                          ([['a', 'b', 'c'],
                            42],
                           'Expected list of columns but got int')])
def test_fix_empty_rows_num_n(capsys: CaptureFixture[str], ind: object,
                              msg: str) -> None:
    """Test OK cases of fix_indata_empty_rows_num."""
    with pytest.raises(AssertionError) as exc:
        fix_indata_empty_rows_num(indata=cast(NumData, ind))
    out, _ = capsys.readouterr()
    assert msg in str(exc)
    assert '' == out


@pytest.mark.parametrize('ind, res',
                         [([['a', 'b', 'c'], ['d', 'e', 'f']],
                           [['a', 'b', 'c'], ['d', 'e', 'f']]),
                          ([['a', 'b'], ['d', 'e']],
                           [['a', 'b'], ['d', 'e']]),
                          ([['a', 'b'], []], [['a', 'b']]),
                          ([['a', 'b'], [None]], [['a', 'b']]),
                          ([['a', 'b'], ['']], [['a', 'b']])])
def test_check_indata_ok_num(capsys: CaptureFixture[str], ind: NumData,
                             res: NumData) -> None:
    """Test check_indata for OK case with num refs."""
    indata = deepcopy(ind)
    check_indata_num(indata=indata)
    out, err = capsys.readouterr()
    assert indata == res
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, exc, msg, excmsg',
                         [([['a', 'b', 'c'], ['d', 'e']], SystemExit,
                           'Number of columns different between lines', None),
                          ('data', AssertionError, None,
                           'Expected list of rows but got'),
                          ([], SystemExit, 'No rows in input data', None),
                          (['data'], AssertionError, None,
                           'Expected list of columns but got'),
                          ([['a', 'b'], ['b']], SystemExit,
                           'Number of columns different between lines', None),
                          ([[], []], SystemExit,
                           'No columns in input data', None)])
def test_check_indata_nok_num(capsys: CaptureFixture[str], ind: object,
                              exc: type[BaseException], msg: Optional[str],
                              excmsg: Optional[str]) -> None:
    """Test check_indata for nok case with num refs."""
    indata = deepcopy(ind)
    with pytest.raises(exc) as exc_inst:
        check_indata_num(indata=cast(NumData, indata))
    out, err = capsys.readouterr()
    if msg is not None:
        assert msg in err
    if excmsg is not None:
        assert excmsg in str(exc_inst)
    assert '' == out


class DataToUseNum(NamedTuple):
    """Data and transformation rules for one number-based transform test."""

    indata: NumData
    split_cols: RuleSplit[int]
    rem_cols: RuleRemove
    merge_cols: RuleMerge[int]
    first_cols: RulePlace
    rename_cols: Rule[int]
    insert_cols: Rule[int]
    rewrite_cols: RuleRewrite[int]
    result: NumData


def get_test_data_num() -> DataToUseNum:
    """Get a test data set (num refs)."""
    ret = DataToUseNum(indata=[['a', 'b c', 'd', 'gå'],
                               ['e', 'f', 'g', 'ÅÄÖåäö']],
                       split_cols=[{'column': 1, 'separator': ' ',
                                    'where': SplitWhere.RIGHTMOST,
                                    'store_single': SplitWhere.LEFTMOST}],
                       rem_cols=[3],
                       merge_cols=[{'columns': [0, 1], 'separator': ' '}],
                       first_cols=[1],
                       rename_cols=[{'column': 1, 'name': 'Second'}],
                       insert_cols=[{'column': 1, 'name': 'the new',
                                    'value': 'x'}],
                       rewrite_cols=[{'column': 2, 'from': 'e', 'to': 'y',
                                      'kind': RewriteKind.STR_SUBSTITUTE,
                                      'case': CaseSensitivity.MATCH_CASE}],
                       result=[['c', 'the new', 'Sycond', 'gå'],
                               [None, 'x', 'y f', 'ÅÄÖåäö']])
    return ret


@pytest.mark.smoke
def test_tr_data_ok_num(capsys: CaptureFixture[str]) -> None:
    """Test transform_data with OK input (num refs)."""
    cfg = ConfigXlsListTransfNum()
    test_data = get_test_data_num()
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = test_data.split_cols
    cfg.s04_remove_columns = test_data.rem_cols
    cfg.s05_merge_columns = test_data.merge_cols
    cfg.s06_place_columns_first = test_data.first_cols
    cfg.s07_rename_columns = test_data.rename_cols
    cfg.s08_insert_columns = test_data.insert_cols
    cfg.s09_rewrite_columns = test_data.rewrite_cols
    res = transform_data_num(indata=test_data.indata, cfg=cfg)
    out, err = capsys.readouterr()
    assert '' == err
    assert '' == out
    assert res == test_data.result


@pytest.mark.parametrize('enc', ['utf-8', 'iso8859-1'])
# pylint: disable-next=duplicate-code
def test_rfmt_nmd_fls_xl2csv(capsys: CaptureFixture[str], enc: str) -> None:
    """Test transform_name_files from xlsx to csv (num refs)."""
    cfg = ConfigXlsListTransfNum()
    configure_input_excel(cfg)
    configure_output_csv(cfg, encoding=enc)
    test_data = get_test_data_num()
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = test_data.split_cols
    cfg.s04_remove_columns = test_data.rem_cols
    cfg.s05_merge_columns = test_data.merge_cols
    cfg.s06_place_columns_first = test_data.first_cols
    cfg.s07_rename_columns = test_data.rename_cols
    cfg.s08_insert_columns = test_data.insert_cols
    cfg.s09_rewrite_columns = test_data.rewrite_cols
    with TemporaryDirectory() as dirname:
        # pylint: disable-next=duplicate-code
        cfgname = dirname + '/test.cfg'
        cfg.write(cfgname)
        infilename = dirname + '/in'
        outfilename = dirname + '/out'
        write_excel_num(data=test_data.indata, filename=infilename + '.xlsx')
        # pylint: disable-next=duplicate-code
        transform_named_files(infilename=infilename, outfilename=outfilename,
                              cfgfilename=cfgname)
        res = read_csv_table(filename=outfilename + '.csv', cfg=cfg)
        out, err = capsys.readouterr()
        assert '' == err
        assert f'Wrote {outfilename}.csv' == out.strip()
        assert res == test_data.result


@pytest.mark.parametrize('enc', ['utf-8', 'iso8859-1'])
# pylint: disable-next=duplicate-code
def test_rfmt_nmd_fls_csv2xl(capsys: CaptureFixture[str], enc: str) -> None:
    """Test transform_name_files from csv to xlsx (num refs)."""
    cfg = ConfigXlsListTransfNum()
    configure_input_csv(cfg, encoding=enc)
    configure_output_excel(cfg)
    test_data = get_test_data_num()
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = test_data.split_cols
    cfg.s04_remove_columns = test_data.rem_cols
    cfg.s05_merge_columns = test_data.merge_cols
    cfg.s06_place_columns_first = test_data.first_cols
    cfg.s07_rename_columns = test_data.rename_cols
    cfg.s08_insert_columns = test_data.insert_cols
    cfg.s09_rewrite_columns = test_data.rewrite_cols
    with TemporaryDirectory() as dirname:
        # pylint: disable-next=duplicate-code
        cfgname = dirname + '/test.cfg'
        cfg.write(cfgname)
        infilename = dirname + '/in'
        outfilename = dirname + '/out'
        write_csv_num(data=test_data.indata, filename=infilename + '.csv',
                      encoding=enc)
        # pylint: disable-next=duplicate-code
        transform_named_files(infilename=infilename, outfilename=outfilename,
                              cfgfilename=cfgname)
        res = read_excel_num(filename=outfilename + '.xlsx', max_col_read=20,
                             strip_col_names=False, strip_values=False)
        out, err = capsys.readouterr()
        assert '' == err
        assert f'Wrote {outfilename}.xlsx' == out.strip()
        assert res == test_data.result
