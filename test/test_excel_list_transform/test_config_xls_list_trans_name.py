#! /usr/local/bin/python3
"""Test the ConfigXlsListRefmtName class."""

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
from excel_list_transform.config_xls_list_transf_name \
    import ConfigXlsListTransfName
from excel_list_transform.config_enums import FileType, SplitWhere
from excel_list_transform.assert_dict_equal import assert_dict_equal
from excel_list_transform.migrate_cfg_warn_hook import MigrateCfgWarnHook


@pytest.mark.smoke
def test_cfg_xls_lst_rfmt_def(capsys: CaptureFixture[str]) -> None:
    """Test default values of ConfigXlsListRefmtName."""
    cfg = ConfigXlsListTransfName()
    assert isinstance(cfg.s03_split_columns, list)
    assert isinstance(cfg.s10_column_order, list)
    assert isinstance(cfg.s05_merge_columns, list)
    assert len(cfg.s05_merge_columns) > 0
    assert 'columns' in cfg.s05_merge_columns[0]
    assert cfg.s05_merge_columns[0]['columns'] == ['street', 'street number']
    assert cfg.s10_column_order == ['Class', 'Division', 'Nationality',
                                    'Sail Number', 'Boat Name', 'First Name',
                                    'Last Name', 'Club Name', 'Email',
                                    'Phone', 'WhatsApp']
    assert len(cfg.s03_split_columns) > 0
    assert 'column' in cfg.s03_split_columns[0]
    assert cfg.s03_split_columns[0]['where'] == SplitWhere.RIGHTMOST
    assert cfg.s03_split_columns[0]['column'] == 'name'
    assert cfg.in_type == FileType.EXCEL
    assert cfg.out_type == FileType.EXCEL
    str_cfg = cfg.as_json_string()
    assert len(str_cfg) > 1
    assert 'input_table' in str_cfg
    assert 'output_table' in str_cfg
    assert 'in_type' not in str_cfg
    zcfg = ConfigXlsListTransfName()
    assert_dict_equal(cfg.__dict__, zcfg.__dict__, ['_hook_cfg_autochange'])
    ycfg = ConfigXlsListTransfName(from_json_text=str_cfg)
    assert_dict_equal(ycfg.__dict__, cfg.__dict__, ['_hook_cfg_autochange'])
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t, val',
                         [('{"s10_column_order" : ["Name"]}', ['Name']),
                          ('{"s10_column_order" : ["ab", "1a"]}',
                           ['ab', '1a'])])
def test_xls_rfmt_rd_inc3(capsys: CaptureFixture[str], t: str,
                          val: list[str]) -> None:
    """Test ConfigXlsListRefmtName read incomplete 3."""
    ycfg = ConfigXlsListTransfName()
    ycfg.parse_json(t, ok_to_use_defaults=True, stderr_file=sys.stderr)
    assert ycfg.s10_column_order == val
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t',
                         ['{"out_type_" : "CSV"}',
                          '{"outfilen" : "out.dat"}'])
def test_xls_rfmt_rd_inc4(capsys: CaptureFixture[str], t: str) -> None:
    """Test ConfigXlsListRefmtName read incomplete 4."""
    cfg = ConfigXlsListTransfName()
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
                          ('{"s10_column_order": {"delimiter" : ";"}}',
                           'Unexpected dictionary for')])
def test_cfg_xls_lst_rfmt_rd(capsys: CaptureFixture[str], t: str,
                             errtxt: str) -> None:
    """Test ConfigXlsListRefmtName read dict mismatch."""
    cfg = ConfigXlsListTransfName()
    with pytest.raises(KeyError) as exc_info:
        cfg.parse_json(t, ok_to_use_defaults=True, stderr_file=sys.stderr)
    assert exc_info.type is KeyError
    out, err = capsys.readouterr()
    assert out == ''
    assert errtxt in err


def test_bak_cmpt_0_7_13_nam(capsys: CaptureFixture[str]) -> None:
    """Test backward compatible reading om 0.7.13 config file."""
    refcfg = ConfigXlsListTransfName()
    refcfg.out_type = FileType.CSV
    refcfg.input_table.character_encoding = 'utf_8_sig'
    refcfg.output_table.character_encoding = 'utf-8'
    refcfg.input_table.csv = TioJsonCsvConfig(dialect=CsvDialect.UNIX,
                                              delimiter=',', quotechar='"')
    refcfg.output_table.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                               delimiter=',', quotechar='"')
    refcfg.s01_split_rows = []
    refcfg.s02_merge_rows = []
    refcfg.s03_split_columns[0]['right_name'] = 'Family Name'
    refcfg.s08_insert_columns[1]['column'] = 'Something Else'
    filename = 'test/test_excel_list_transform/bak_compat_0_7_13_name.cfg'
    cfg = ConfigXlsListTransfName(from_json_filename=filename)
    out, err = capsys.readouterr()
    assert_dict_equal(refcfg.__dict__, cfg.__dict__, ['_hook_cfg_autochange'])
    assert cfg.s03_split_columns[0]['right_name'] == 'Family Name'
    assert cfg.s08_insert_columns[1]['column'] == 'Something Else'
    assert '' == out
    assert MigrateCfgWarnHook.migrate_warn_msg() == err


@pytest.mark.parametrize('splitr',
                         [[],
                          [{'column': 'foo', 'separators': [';'],
                            'not_separators': ['\\;']}],
                          [{'column': 'foo', 'separators': ['+'],
                            'not_separators': [' + ']},
                           {'column': 'bar', 'separators': ['+', '-'],
                            'not_separators': [' + ', '--']}]])
@pytest.mark.parametrize('merger',
                         [[],
                          [{'columns': ['foo', 'bar'], 'separator': ' '}],
                          [{'columns': ['foo', 'bar'], 'separator': ' '},
                           {'columns': ['col1', 'col2'], 'separator': ';'}]])
def test_row_spl_mrg_na_ok(capsys: CaptureFixture[str], splitr: Any,
                           merger: Any) -> None:
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListTransfName()
    cfg1.s01_split_rows = deepcopy(splitr)
    cfg1.s02_merge_rows = deepcopy(merger)
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListTransfName(from_json_text=txt)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert_dict_equal(cfg1.__dict__, cfg2.__dict__, ['_hook_cfg_autochange'])
    assert splitr == cfg2.s01_split_rows
    assert merger == cfg2.s02_merge_rows


@pytest.mark.parametrize('splitr, msgs',
                         [([{'column': 'foo', 'separators': [],
                            'not_separators': []}],
                           ['Error in parameter s01_split_rows.',
                            'List for key separators shall be ' +
                            'minimum 1 elements',
                            'But it is 0 elements']),
                          ([{'column': 'foo', 'separators': [],
                            'not_separators': [], 'start': 2}],
                           ['Found non-allowed key "start" in ',
                            ' config of s01_split_rows']),
                          ([{'column': 2, 'separators': [';'],
                            'not_separators': [';;']}],
                           ['Error in parameter s01_split_rows.',
                            'Value for key column expected to be of ' +
                            'type str but is of type int']),
                          ([{'separators': ['+'],
                             'not_separators': [' + ']}],
                           ['Error in parameter s01_split_rows.',
                            'Expected key column not in dict in list']),
                          ([{'column': 'bar', 'separators': ['+', '-'],
                             'not_separators': [' + ', '--', '*']}],
                           ['Error in s01_split_rows:',
                            'Not separator "*" does not affect ' +
                            'any separator.'])])
def test_row_split_cfg_na_nok(capsys: CaptureFixture[str], splitr: Any,
                              msgs: list[str]) -> None:
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListTransfName()
    cfg1.s01_split_rows = deepcopy(splitr)
    with pytest.raises(SystemExit):
        _ = cfg1.as_json_string()
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err


@pytest.mark.parametrize('merger,msgs',
                         [([{'columns': [], 'separator': ' '}],
                           ['Error in parameter s02_merge_rows.',
                            'List for key columns shall be minimum 1 eleme',
                            'But it is 0 elements only.']),
                          ([{'columns': ['foo', 'bar'], 'separator': [' ']}],
                           ['Error in parameter s02_merge_rows.',
                            'Value for key separator expected to ' +
                            'be of type str but is of type list']),
                          ([{'columns': ['foo', 'bar'], 'separator': ' ',
                             'split': 2}],
                           ['Found non-allowed key "split" ',
                            'in config of s02_merge_rows']),
                          ([{'columns': [1, 2], 'separator': ' '}],
                           ['Error in parameter s02_merge_rows.',
                            'Value for key columns expected to be list of str',
                            'But element in list is int']),
                          ([{'columns': ['foo', 'bar'], 'separator': 3}],
                           ['Error in parameter s02_merge_rows.',
                            'Value for key separator expected to be ' +
                            'of type str but is of type int']),
                          ([{'columns': 'foo', 'separator': ' '}],
                           ['Error in parameter s02_merge_rows.',
                            'Value for key columns expected to be ' +
                            'of type list but is of type str'])])
def test_row_merge_cfg_na_nok(capsys: CaptureFixture[str], merger: Any,
                              msgs: list[str]) -> None:
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListTransfName()
    cfg1.s02_merge_rows = deepcopy(merger)
    with pytest.raises(SystemExit):
        _ = cfg1.as_json_string()
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err
