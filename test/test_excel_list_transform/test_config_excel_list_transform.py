#! /usr/local/bin/python3
"""Test the ConfigExcelListTransform class."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from typing import Any
from collections import namedtuple
from copy import deepcopy
import sys
import pytest
from pytest import CaptureFixture
from config_as_json import InvalidConfiguration
from excel_list_transform.config_enums import ColumnRef, FileType
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform, ConfigReadOld, ColInfo
from excel_list_transform.assert_dict_equal import assert_dict_equal


@pytest.mark.smoke
def test_cfg_xlt_lst_def_def(capsys: CaptureFixture[str]) -> None:
    """Test default values of ConfigExcelListTransform."""
    col_to_use = ['street', 'street number', 'name', 'last name',
                  'Phone', 'Phone', 'Phone', 'Phone', 'Phone',
                  'Last Name']
    col_to_use_row = ['Club Name', 'name', 'last name']
    colinfo = ColInfo('right_name', None, [], [], col_to_use, col_to_use_row,
                      'a')
    cfg = ConfigExcelListTransform(col_ref=ColumnRef.BY_NAME, colinfo=colinfo,
                                   tinfo='a')
    assert cfg.in_type == FileType.EXCEL
    assert cfg.out_type == FileType.EXCEL
    assert cfg.column_ref == ColumnRef.BY_NAME
    assert cfg.max_column_read == 20
    str_cfg = cfg.as_json_string()
    assert len(str_cfg) > 1
    assert 'input_table' in str_cfg
    assert 'output_table' in str_cfg
    assert 'in_type' not in str_cfg
    assert 'out_type' not in str_cfg
    zcfg = ConfigExcelListTransform(col_ref=ColumnRef.BY_NAME, colinfo=colinfo,
                                    tinfo='a')
    assert_dict_equal(cfg.__dict__, zcfg.__dict__, ['_hook_cfg_autochange'])
    ycfg = ConfigExcelListTransform(col_ref=ColumnRef.BY_NAME, colinfo=colinfo,
                                    tinfo='a', from_json_text=str_cfg)
    assert_dict_equal(cfg.__dict__, ycfg.__dict__, ['_hook_cfg_autochange'])
    assert cfg.get_out_csv_dialect().lineterminator == \
        ycfg.get_out_csv_dialect().lineterminator
    assert ycfg.get_out_csv_dialect().lineterminator == '\r\n'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t',
                         ['{"out_type_" : "CSV"}',
                          '{"outfilen" : "out.dat"}'])
def test_xlt_rd_rd_inc4(capsys: CaptureFixture[str], t: Any) -> None:
    """Test ConfigExcelListTransform read incomplete 4."""
    col_to_use = [15, 16, 1, 2, 5, 5, 5, 5, 5, 6]
    col_to_use_row = [7, 1, 2]
    colinfo = ColInfo('store_single', 'name', [], [], col_to_use,
                      col_to_use_row, 2)
    cfg = ConfigExcelListTransform(col_ref=ColumnRef.BY_NUMBER,
                                   colinfo=colinfo, tinfo=2)
    with pytest.raises(KeyError) as exc_info:
        cfg.parse_json(t, ok_to_use_defaults=True, stderr_file=sys.stderr)
    assert exc_info.type is KeyError
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Unexpected' in err


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'column': 2},
                            {'a': 1, 'column': 3},
                            {'a': 3, 'column': 4}], 'test')])
def test_check_no_dupl_key_ok(capsys: CaptureFixture[str], data: Any,
                              par: Any) -> None:
    """Test check:_no_cuplicates for OK cases."""
    # pylint: disable-next=protected-access
    ConfigExcelListTransform._check_no_duplicate_single(data, par, 2)
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'columns': [1, 2]},
                            {'a': 1, 'columns': [3, 7]},
                            {'a': 3, 'columns': [4, 8]}], 'test')])
def test_check_no_dupl_mul_ok(capsys: CaptureFixture[str], data: Any,
                              par: Any) -> None:
    """Test check:_no_cuplicates for OK cases."""
    # pylint: disable-next=protected-access
    ConfigExcelListTransform._check_no_duplicate_multi(data, par, 2)
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'column': 2},
                            {'a': 1, 'column': 3},
                            {'a': 3, 'column': 2}], 'test')])
def test_chk_no_dupl_keyd_nok(capsys: CaptureFixture[str], data: Any,
                              par: Any) -> None:
    """Test check:_no_cuplicates for OK cases."""
    with pytest.raises(SystemExit):
        # pylint: disable-next=protected-access
        ConfigExcelListTransform._check_no_duplicate_single(data, par, 2)
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Duplicates not allowed in' in err


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'columns': [1, 2]},
                            {'a': 1, 'columns': [3, 4]},
                            {'a': 3, 'columns': [2, 5]}], 'test')])
def test_chk_no_dupl_num_nok(capsys: CaptureFixture[str], data: Any,
                             par: Any) -> None:
    """Test check:_no_cuplicates for OK cases."""
    with pytest.raises(SystemExit):
        # pylint: disable-next=protected-access
        ConfigExcelListTransform._check_no_duplicate_multi(data, par, 2)
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Duplicates not allowed in' in err


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'columns': [2, 4]},
                            {'a': 1, 'columns': [5, 6]},
                            {'a': 3, 'columns': [7, 8]}], 'test')])
def test_check_increasing_ok(capsys: CaptureFixture[str], data: Any,
                             par: Any) -> None:
    """Test check:_no_cuplicates for OK cases."""
    # pylint: disable-next=protected-access
    ConfigExcelListTransform._check_increasing_multi(data, par, 2)
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'columns': [2, 4]},
                            {'a': 1, 'columns': [3, 6]},
                            {'a': 3, 'columns': [7, 8]}], 'test')])
def test_check_increasing_nok(capsys: CaptureFixture[str], data: Any,
                              par: Any) -> None:
    """Test check:_no_cuplicates for OK cases."""
    with pytest.raises(KeyError) as exc:
        # pylint: disable-next=protected-access
        ConfigExcelListTransform._check_increasing_multi(data, par, 2)
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Increasing order needed in' in err
    assert 'Increasing order needed in' in str(exc)


MockInitArgs = namedtuple('MockInitArgs', ['colref', 'colinfo', 'tinfo'])


def get_mock_init_args(colref: ColumnRef) -> Any:
    """Get arguments for  ConfigExcelListTransform init."""
    if colref == ColumnRef.BY_NUMBER:
        col_to_us1 = [2, 3, 0, 1, 4, 4, 4, 4, 4, 1]
        col_to_use_r1 = [7, 1, 2]
        colinf1 = ColInfo('right_name', None, [], [], col_to_us1,
                          col_to_use_r1, 2)
        return MockInitArgs(colref=colref, colinfo=colinf1, tinfo=2)
    col_to_use = ['street', 'street number', 'name', 'last name',
                  'Phone', 'Phone', 'Phone', 'Phone', 'Phone',
                  'Last Name']
    col_to_use_r2 = ['Club Name', 'name', 'last name']
    colinf2 = ColInfo('right_name', None, [], [], col_to_use, col_to_use_r2,
                      'a')
    return MockInitArgs(colref=colref, colinfo=colinf2, tinfo='a')


@pytest.mark.parametrize('cref', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
def test_cfg_transf_def_vals(capsys: CaptureFixture[str], cref: Any) -> None:
    """Test defaults supplied for old configuration files."""
    _ = cref
    data = ConfigReadOld._old_defaults()  # pylint: disable=protected-access
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert len(data) == 10
    assert 'in_csv_encoding' in data
    assert 'out_csv_encoding' in data
    assert 'in_excel_col_name_strip' in data
    assert 'in_excel_values_strip' in data
    assert 's01_split_rows' in data
    assert 's02_merge_rows' in data
    assert 'output_borders' in data
    assert 'output_filtered_table' in data
    split_rows = data['s01_split_rows']
    merge_rows = data['s02_merge_rows']
    assert isinstance(split_rows, list)
    assert isinstance(merge_rows, list)
    assert len(split_rows) == 0
    assert len(merge_rows) == 0
    assert 'utf_8_sig' == data['in_csv_encoding']
    assert 'utf-8' == data['out_csv_encoding']


@pytest.mark.parametrize('cref', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
def test_cfg_tr_enc_def(capsys: CaptureFixture[str], cref: Any) -> None:
    """Test encoding in default constructed ConfigExcelListTransform."""
    args = get_mock_init_args(colref=cref)
    cfg = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo)
    assert 'utf_8_sig' == cfg.in_csv_encoding
    assert 'utf-8' == cfg.out_csv_encoding
    txt = cfg.as_json_string()
    assert 'character_encoding' not in txt
    assert 'in_csv_encoding' not in txt
    assert 'out_csv_encoding' not in txt
    cf2 = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo, from_json_text=txt,
                                   from_json_filename=None)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert 'utf_8_sig' == cf2.in_csv_encoding
    assert 'utf-8' == cf2.out_csv_encoding


@pytest.mark.parametrize('cref', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
@pytest.mark.parametrize('in_enc, out_enc',
                         [('utf-8', 'iso8859-1'),
                          ('iso8859-2', 'ascii')])
@pytest.mark.parametrize('scol', [False, True])
@pytest.mark.parametrize('sval', [False, True])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_cfg_transf_enc_1_ok(capsys: CaptureFixture[str], in_enc: Any,
                             out_enc: Any, cref: Any, scol: Any,
                             sval: Any) -> None:
    """Test configured encoding for ConfigExcelListTransform."""
    args = get_mock_init_args(colref=cref)
    cfg = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo)
    assert 'utf_8_sig' == cfg.in_csv_encoding
    assert 'utf-8' == cfg.out_csv_encoding
    cfg.in_csv_encoding = in_enc
    cfg.out_csv_encoding = out_enc
    cfg.in_excel_col_name_strip = scol
    cfg.in_excel_values_strip = sval
    assert cfg.in_csv_encoding == in_enc
    assert cfg.out_csv_encoding == out_enc
    txt = cfg.as_json_string()
    assert in_enc in txt
    assert out_enc in txt
    cf2 = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo, from_json_text=txt,
                                   from_json_filename=None)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert cf2.in_csv_encoding == in_enc
    assert cf2.out_csv_encoding == out_enc
    assert cf2.in_excel_col_name_strip == scol
    assert cf2.in_excel_values_strip == sval


@pytest.mark.parametrize('cref', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
@pytest.mark.parametrize('in_enc, out_enc',
                         [('utf-88', 'iso8859-1'),
                          ('ixo8859-2', 'ascii')])
def test_cfg_transf_enc_1_nok(capsys: CaptureFixture[str], in_enc: Any,
                              out_enc: Any, cref: Any) -> None:
    """Test not OK configured encoding for ConfigExcelListTransform."""
    with pytest.raises(InvalidConfiguration):
        arg = get_mock_init_args(colref=cref)
        cfg = ConfigExcelListTransform(col_ref=arg.colref, colinfo=arg.colinfo,
                                       tinfo=arg.tinfo)
        cfg.in_csv_encoding = in_enc
        cfg.out_csv_encoding = out_enc
        _ = cfg.as_json_string(stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert '' == out
    assert 'is not a recognized character encoding' in err


@pytest.mark.parametrize('cref', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
@pytest.mark.parametrize('in_enc, out_enc',
                         [('utf-8', 'iso8859-1'),
                          ('iso8859-2', 'ascii')])
def test_cfg_transf_enc_2_ok(capsys: CaptureFixture[str], in_enc: Any,
                             out_enc: Any, cref: Any) -> None:
    """Test default encoding for ConfigExcelListTransform."""
    args = get_mock_init_args(colref=cref)
    cfg = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo)
    assert 'utf_8_sig' == cfg.in_csv_encoding
    assert 'utf-8' == cfg.out_csv_encoding
    cfg.in_csv_encoding = in_enc
    cfg.out_csv_encoding = out_enc
    assert cfg.in_csv_encoding == in_enc
    assert cfg.out_csv_encoding == out_enc
    txt = cfg.as_json_string()
    lines = txt.splitlines()
    filtered_lines = [row for row in lines if 'encoding' not in row]
    filtered_txt = '\n'.join(filtered_lines)
    cf2 = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo,
                                   from_json_text=filtered_txt,
                                   from_json_filename=None)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert cf2.in_csv_encoding == 'utf_8_sig'
    assert cf2.out_csv_encoding == 'utf-8'


@pytest.mark.parametrize('sep, nosep',
                         [([';'], ['\\;']),
                          ([';', '+'], ['\\;', '++', ' + '])])
def test_check_sep_not_sep_ok(capsys: CaptureFixture[str], sep: Any,
                              nosep: Any) -> None:
    """Test OK cases of check_sep_not_sep."""
    ConfigExcelListTransform.check_sep_not_sep(separators=sep,
                                               not_separators=nosep)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('sep, nosep, msgs',
                         [([';'], ['\\;', ';'],
                           ['Cannot have same string as both separator and ',
                            'separator and not separator: ;']),
                          ([';', '+'], ['\\;', '++', ' '],
                           ['Not separator " " does not affect any separator'
                            ])])
def test_chk_sep_not_sep_nok(capsys: CaptureFixture[str], sep: Any, nosep: Any,
                             msgs: Any) -> None:
    """Test OK cases of check_sep_not_sep."""
    with pytest.raises(SystemExit):
        ConfigExcelListTransform.check_sep_not_sep(separators=sep,
                                                   not_separators=nosep)
    out, err = capsys.readouterr()
    assert '' == out
    assert 'Error in s01_split_rows:\n' in err
    for msg in msgs:
        assert msg in err


@pytest.mark.parametrize('cref', [(ColumnRef.BY_NAME, 'a'),
                                  (ColumnRef.BY_NUMBER, 2)])
@pytest.mark.parametrize('rsplit',
                         [[], [{'column': None, 'separators': [';'],
                                'not_separators': ['\\;', ' ; ']}],
                          [{'column': None, 'separators': [';', '+'],
                            'not_separators': ['\\;', ' + ']}]])
def test_chk_spl_row_cfg_ok(capsys: CaptureFixture[str], cref: Any,
                            rsplit: Any) -> None:
    """Test OK cases for check_split_row_cfg."""
    args = get_mock_init_args(colref=cref[0])
    cfg = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo)
    rsplit2 = deepcopy(rsplit)
    for i in rsplit2:
        i['column'] = cref[1]
    cfg.s01_split_rows = deepcopy(rsplit2)
    cfg.check_split_row_cfg()
    out, err = capsys.readouterr()
    assert rsplit2 == cfg.s01_split_rows
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('cref', [(ColumnRef.BY_NAME, 'a'),
                                  (ColumnRef.BY_NUMBER, 2)])
@pytest.mark.parametrize('rsplit, msgs',
                         [([[]],
                           ['Expected dict in list but found list']),
                          ([{'column': None, 'separators': [2, 3],
                             'not_separators': ['\\;', '\\2']}],
                           ['But element in list is int',
                            'key separators expected to be list of str']),
                          ([{'separators': ['2', '3'],
                             'not_separators': ['\\;',]}],
                           ['Expected key column not in dict in list']),
                          ([{'column': None, 'separators': [';', '+'],
                            'not_separators': '\\;'}],
                           ['Value for key not_separators expected to ',
                            'be of type list but is of type str'])])
def test_chk_spl_row_cfg_nok(capsys: CaptureFixture[str], cref: Any,
                             rsplit: Any, msgs: Any) -> None:
    """Test not OK cases for check_split_row_cfg."""
    args = get_mock_init_args(colref=cref[0])
    cfg = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo)
    rsplit2 = deepcopy(rsplit)
    for i in rsplit2:
        if 'column' in i:
            i['column'] = cref[1]
    cfg.s01_split_rows = deepcopy(rsplit2)
    with pytest.raises(SystemExit):
        cfg.check_split_row_cfg()
    out, err = capsys.readouterr()
    assert '' == out
    assert 'Error in parameter s01_split_rows.' in err
    for msg in msgs:
        assert msg in err


@pytest.mark.parametrize('cref', [(ColumnRef.BY_NAME, ['a', 'b']),
                                  (ColumnRef.BY_NUMBER, [2, 4])])
@pytest.mark.parametrize('rmerg',
                         [[], [{'columns': None, 'separator': ';'}],
                          [{'columns': None, 'separator': '+'}],
                          [{'columns': None, 'separator': ';'},
                           {'columns': None, 'separator': ' '}]])
def test_chk_mrg_row_cfg_ok(capsys: CaptureFixture[str], cref: Any,
                            rmerg: Any) -> None:
    """Test OK cases for check_split_row_cfg."""
    args = get_mock_init_args(colref=cref[0])
    cfg = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo)
    rsplit2 = deepcopy(rmerg)
    for i in rsplit2:
        i['columns'] = cref[1]
    cfg.s02_merge_rows = deepcopy(rsplit2)
    cfg.check_merge_row_cfg()
    out, err = capsys.readouterr()
    assert rsplit2 == cfg.s02_merge_rows
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('cref', [(ColumnRef.BY_NAME, ['a', 'b']),
                                  (ColumnRef.BY_NUMBER, [2, 4])])
@pytest.mark.parametrize('rmerg, msgs',
                         [([{'columns': None}],
                           ['Expected key separator not in dict in list',
                            'Error in parameter s02_merge_rows.']),
                          ([{'separator': ';'}],
                           ['Expected key columns not in dict in list',
                            'Error in parameter s02_merge_rows.']),
                          ([{'columns': None, 'separator': ['a', 'b']}],
                           ['Value for key separator expected to be of ',
                            ' to be of type str but is of type list',
                            'Error in parameter s02_merge_rows.']),
                          ([{'columns': None, 'separator': '+',
                             'not_separator': '\\+'}],
                           ['Found non-allowed key "not_separator" ',
                            'in config of s02_merge_rows'])])
def test_chk_mrg_row_cfg_nok(capsys: CaptureFixture[str], cref: Any,
                             rmerg: Any, msgs: Any) -> None:
    """Test not OK cases for check_split_row_cfg."""
    args = get_mock_init_args(colref=cref[0])
    cfg = ConfigExcelListTransform(col_ref=args.colref, colinfo=args.colinfo,
                                   tinfo=args.tinfo)
    rsplit2 = deepcopy(rmerg)
    for i in rsplit2:
        if 'columns' in i:
            i['columns'] = cref[1]
    cfg.s02_merge_rows = deepcopy(rsplit2)
    with pytest.raises(SystemExit):
        cfg.check_merge_row_cfg()
    out, err = capsys.readouterr()
    assert rsplit2 == cfg.s02_merge_rows
    assert '' == out
    for msg in msgs:
        assert msg in err
