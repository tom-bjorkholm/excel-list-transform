#! /usr/local/bin/python3
"""Test the ConfigXlsListRefmtNum class."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from typing import Any
from copy import deepcopy
import sys
import pytest
from pytest import CaptureFixture
from tableio import CsvDialect
from tableio_cfg_json import TioJsonCsvConfig
from test_excel_list_transform.tableio_helpers import \
    configure_output_csv
from excel_list_transform.config_xls_list_transf_num \
    import ConfigXlsListTransfNum
from excel_list_transform.config_enums import SplitWhere
from excel_list_transform.assert_dict_equal import assert_dict_equal
from excel_list_transform.migrate_cfg_warn_hook import MigrateCfgWarnHook


@pytest.mark.smoke
def test_cfg_xls_lst_rfmt_def(capsys: CaptureFixture[str]) -> None:
    """Test default values of ConfigXlsListRefmtNum."""
    cfg = ConfigXlsListTransfNum()
    assert isinstance(cfg.s03_split_columns, list)
    assert isinstance(cfg.s04_remove_columns, list)
    assert isinstance(cfg.s05_merge_columns, list)
    assert isinstance(cfg.s06_place_columns_first, list)
    assert len(cfg.s05_merge_columns) > 0
    assert 'columns' in cfg.s05_merge_columns[0]
    assert cfg.s05_merge_columns[0]['columns'] == [15, 16]
    assert cfg.s04_remove_columns == [1, 2, 3]
    assert cfg.s06_place_columns_first == [7, 3, 6]
    assert len(cfg.s03_split_columns) > 0
    assert 'column' in cfg.s03_split_columns[0]
    assert cfg.s03_split_columns[0]['where'] == SplitWhere.RIGHTMOST
    assert cfg.s03_split_columns[0]['column'] == 15
    assert cfg.input_table.format_name == 'Excel'
    assert cfg.output_table.format_name == 'Excel'
    str_cfg = cfg.as_json_string(stderr_file=sys.stderr)
    assert len(str_cfg) > 1
    assert 'input_table' in str_cfg
    assert 'output_table' in str_cfg
    assert 'in_type' not in str_cfg
    zcfg = ConfigXlsListTransfNum()
    assert_dict_equal(cfg.__dict__, zcfg.__dict__, ['_hook_cfg_autochange'])
    ycfg = ConfigXlsListTransfNum(from_json_data_text=str_cfg)
    assert_dict_equal(ycfg.__dict__, cfg.__dict__, ['_hook_cfg_autochange'])
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t, val',
                         [('{"s04_remove_columns" : [2]}', [2]),
                          ('{"s04_remove_columns" : [21, 14]}', [14, 21])])
def test_xls_rfmt_rd_inc3(capsys: CaptureFixture[str], t: str,
                          val: list[int]) -> None:
    """Test ConfigXlsListRefmtNum read incomplete 3."""
    ycfg = ConfigXlsListTransfNum()
    ycfg.parse_json(t, ok_to_use_defaults=True, stderr_file=sys.stderr)
    assert ycfg.s04_remove_columns == val
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t',
                         ['{"out_type_" : "CSV"}',
                          '{"outfilen" : "out.dat"}'])
def test_xls_rfmt_rd_inc4(capsys: CaptureFixture[str], t: str) -> None:
    """Test ConfigXlsListRefmtNum read incomplete 4."""
    cfg = ConfigXlsListTransfNum()
    with pytest.raises(KeyError) as exc_info:
        cfg.parse_json(t, ok_to_use_defaults=True, stderr_file=sys.stderr)
    assert exc_info.type is KeyError
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Unexpected' in err


@pytest.mark.parametrize('t, errtxt',
                         [('{'
                           '"input_table": {'
                           '"format_name": "CSV", "csv": "B"}}',
                           'Nested Config member csv must be a JSON object'),
                          ('{"s04_remove_columns": {"delimiter" : ";"}}',
                           'Unexpected dictionary for')])
def test_cfg_xls_lst_rfmt_rd(capsys: CaptureFixture[str], t: str,
                             errtxt: str) -> None:
    """Test ConfigXlsListRefmtNum read dict mismatch."""
    cfg = ConfigXlsListTransfNum()
    with pytest.raises(KeyError) as exc_info:
        cfg.parse_json(t, ok_to_use_defaults=True, stderr_file=sys.stderr)
    assert exc_info.type is KeyError
    out, err = capsys.readouterr()
    assert out == ''
    assert errtxt in err


def test_bak_cmpt_0_7_13_num(capsys: CaptureFixture[str]) -> None:
    """Test backward compatible reading om 0.7.13 config file."""
    refcfg = ConfigXlsListTransfNum()
    configure_output_csv(refcfg)
    refcfg.input_table.character_encoding = 'utf_8_sig'
    refcfg.output_table.character_encoding = 'utf-8'
    refcfg.input_table.csv = TioJsonCsvConfig(dialect=CsvDialect.UNIX,
                                              delimiter=',', quotechar='"')
    refcfg.output_table.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                               delimiter=',', quotechar='"')
    refcfg.s01_split_rows = []
    refcfg.s02_merge_rows = []
    refcfg.s07_rename_columns[1]['name'] = 'Family Name'
    refcfg.s08_insert_columns[1]['name'] = 'Something else'
    filename = 'test/test_excel_list_transform/bak_compat_0_7_13_number.cfg'
    cfg = ConfigXlsListTransfNum(from_json_filename=filename,
                                 stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert_dict_equal(refcfg.__dict__, cfg.__dict__, ['_hook_cfg_autochange'])
    assert cfg.s07_rename_columns[1]['name'] == 'Family Name'
    assert cfg.s08_insert_columns[1]['name'] == 'Something else'
    assert '' == out
    assert MigrateCfgWarnHook.migrate_warn_msg() == err


@pytest.mark.parametrize('splitr',
                         [[],
                          [{'column': 2, 'separators': [';'],
                            'not_separators': ['\\;']}],
                          [{'column': 3, 'separators': ['+'],
                            'not_separators': [' + ']},
                           {'column': 4, 'separators': ['+', '-'],
                            'not_separators': [' + ', '--']}]])
@pytest.mark.parametrize('merger',
                         [[],
                          [{'columns': [1, 2], 'separator': ' '}],
                          [{'columns': [3, 4], 'separator': ' '},
                           {'columns': [5, 6], 'separator': ';'}]])
def test_row_spl_mrg_nu_ok(capsys: CaptureFixture[str], splitr: Any,
                           merger: Any) -> None:
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListTransfNum()
    cfg1.s01_split_rows = deepcopy(splitr)
    cfg1.s02_merge_rows = deepcopy(merger)
    txt = cfg1.as_json_string(stderr_file=sys.stderr)
    cfg2 = ConfigXlsListTransfNum(from_json_data_text=txt)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert_dict_equal(cfg1.__dict__, cfg2.__dict__, ['_hook_cfg_autochange'])
    assert splitr == cfg2.s01_split_rows
    assert merger == cfg2.s02_merge_rows


@pytest.mark.parametrize('splitr, msgs',
                         [([{'column': 2, 'separators': [],
                            'not_separators': []}],
                           ['Error in parameter s01_split_rows.',
                            'List for key separators shall be ' +
                            'minimum 1 elements',
                            'But it is 0 elements']),
                          ([{'column': 3, 'separators': [],
                            'not_separators': [], 'start': 2}],
                           ['Found non-allowed key "start" in ',
                            ' config of s01_split_rows']),
                          ([{'column': 'foo', 'separators': [';'],
                            'not_separators': [';;']}],
                           ['Error in parameter s01_split_rows.',
                            'Value for key column expected to be of ' +
                            'type int but is of type str']),
                          ([{'separators': ['+'],
                             'not_separators': [' + ']}],
                           ['Error in parameter s01_split_rows.',
                            'Expected key column not in dict in list']),
                          ([{'column': 4, 'separators': ['+', '-'],
                             'not_separators': [' + ', '--', '*']}],
                           ['Error in s01_split_rows:',
                            'Not separator "*" does not affect ' +
                            'any separator.'])])
def test_row_split_cfg_nu_nok(capsys: CaptureFixture[str], splitr: Any,
                              msgs: list[str]) -> None:
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListTransfNum()
    cfg1.s01_split_rows = deepcopy(splitr)
    with pytest.raises(SystemExit):
        _ = cfg1.as_json_string(stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err


@pytest.mark.parametrize('merger,msgs',
                         [([{'columns': [], 'separator': ' '}],
                           ['Error in parameter s02_merge_rows.',
                            'List for key columns shall be minimum 1 eleme',
                            'But it is 0 elements only.']),
                          ([{'columns': [3, 4], 'separator': [' ']}],
                           ['Error in parameter s02_merge_rows.',
                            'Value for key separator expected to ' +
                            'be of type str but is of type list']),
                          ([{'columns': [5, 6], 'separator': ' ',
                             'split': 2}],
                           ['Found non-allowed key "split" ',
                            'in config of s02_merge_rows']),
                          ([{'columns': ['foo', 'bar'], 'separator': ' '}],
                           ['Error in parameter s02_merge_rows.',
                            'Value for key columns expected to be list of int',
                            'But element in list is str']),
                          ([{'columns': [7, 8], 'separator': 3}],
                           ['Error in parameter s02_merge_rows.',
                            'Value for key separator expected to be ' +
                            'of type str but is of type int']),
                          ([{'columns': 9, 'separator': ' '}],
                           ['Error in parameter s02_merge_rows.',
                            'Value for key columns expected to be ' +
                            'of type list but is of type int'])])
def test_row_merge_cfg_na_nok(capsys: CaptureFixture[str], merger: Any,
                              msgs: list[str]) -> None:
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListTransfNum()
    cfg1.s02_merge_rows = deepcopy(merger)
    with pytest.raises(SystemExit):
        _ = cfg1.as_json_string(stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err
