#! /usr/local/bin/python3
"""Test the ConfigXlsListRefmtName class."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from typing import cast
from copy import deepcopy
import sys
import pytest
from pytest import CaptureFixture
from config_as_json import InvalidConfiguration, assert_dict_equal
from tableio import CsvDialect
from tableio_cfg_json import TioJsonCsvConfig
from test_excel_list_transform.tableio_helpers import \
    configure_output_csv
from excel_list_transform.config_xls_list_transf_name \
    import ConfigXlsListTransfName
from excel_list_transform.config_enums import SplitWhere
from excel_list_transform.config_excel_list_transform import RuleMerge, \
    RuleRowSplit
from excel_list_transform.migrate_cfg_warn_hook import EltMigrateCfgWarnHook

LEGACY_RRS_ORDER = ['Class', 'Division', 'Nationality', 'Sail Number',
                    'Boat Name', 'First Name', 'Last Name', 'Club Name',
                    'Email', 'Phone', 'WhatsApp']


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
                                    'MNA No.', 'Sail Number', 'Boat Name',
                                    'First Name', 'Last Name', 'Club Name',
                                    'Email', 'Phone', 'Whats App Number']
    assert len(cfg.s03_split_columns) > 0
    assert 'column' in cfg.s03_split_columns[0]
    assert cfg.s03_split_columns[0]['where'] == SplitWhere.RIGHTMOST
    assert cfg.s03_split_columns[0]['column'] == 'name'
    assert cfg.input_table.format_name == 'Excel'
    assert cfg.output_table.format_name == 'Excel'
    str_cfg = cfg.as_json_string(stderr_file=sys.stderr)
    assert len(str_cfg) > 1
    assert 'input_table' in str_cfg
    assert 'output_table' in str_cfg
    assert 'in_type' not in str_cfg
    zcfg = ConfigXlsListTransfName()
    assert_dict_equal(cfg.__dict__, zcfg.__dict__, ['_hook_cfg_autochange'])
    ycfg = ConfigXlsListTransfName(from_json_data_text=str_cfg)
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
    assert EltMigrateCfgWarnHook.migrate_warn_msg() == err


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
    configure_output_csv(refcfg)
    refcfg.input_table.character_encoding = 'utf_8_sig'
    refcfg.output_table.character_encoding = 'utf-8'
    refcfg.output_borders = False
    refcfg.output_filtered_table = False
    refcfg.input_table.csv = TioJsonCsvConfig(dialect=CsvDialect.UNIX,
                                              delimiter=',', quotechar='"')
    refcfg.output_table.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                               delimiter=',', quotechar='"')
    refcfg.s01_split_rows = []
    refcfg.s02_merge_rows = []
    refcfg.s03_split_columns[0]['right_name'] = 'Family Name'
    refcfg.s08_insert_columns[1]['column'] = 'Something Else'
    refcfg.s10_column_order = deepcopy(LEGACY_RRS_ORDER)
    filename = 'test/test_excel_list_transform/bak_compat_0_7_13_name.cfg'
    cfg = ConfigXlsListTransfName(from_json_filename=filename,
                                  stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert_dict_equal(refcfg.__dict__, cfg.__dict__, ['_hook_cfg_autochange'])
    assert cfg.s03_split_columns[0]['right_name'] == 'Family Name'
    assert cfg.s08_insert_columns[1]['column'] == 'Something Else'
    assert '' == out
    assert EltMigrateCfgWarnHook.migrate_warn_msg() == err


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
def test_row_spl_mrg_na_ok(capsys: CaptureFixture[str],
                           splitr: RuleRowSplit[str],
                           merger: RuleMerge[str]) -> None:
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListTransfName()
    cfg1.s01_split_rows = deepcopy(splitr)
    cfg1.s02_merge_rows = deepcopy(merger)
    txt = cfg1.as_json_string(stderr_file=sys.stderr)
    cfg2 = ConfigXlsListTransfName(from_json_data_text=txt)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert_dict_equal(cfg1.__dict__, cfg2.__dict__, ['_hook_cfg_autochange'])
    assert splitr == cfg2.s01_split_rows
    assert merger == cfg2.s02_merge_rows


@pytest.mark.parametrize('splitr, msgs',
                         [([{'column': 'foo', 'separators': [],
                            'not_separators': []}],
                           ['s01_split_rows[0][separators]',
                            'less than minimum 1']),
                          ([{'column': 'foo', 'separators': [],
                            'not_separators': [], 'start': 2}],
                           ["Unknown key 'start' in s01_split_rows[0]"]),
                          ([{'column': 2, 'separators': [';'],
                            'not_separators': [';;']}],
                           ['s01_split_rows[0][column]',
                            'not of type str']),
                          ([{'separators': ['+'],
                             'not_separators': [' + ']}],
                           ["Mandatory key 'column' is missing",
                            's01_split_rows[0]']),
                          ([{'column': 'foo', 'separators': ['+'],
                             'not_separators': ['+']}],
                           ['s01_split_rows[0]',
                            'same string as a separator',
                            "not-separator: '+'."]),
                          ([{'column': 'bar', 'separators': ['+', '-'],
                             'not_separators': [' + ', '--', '*']}],
                           ['s01_split_rows[0]',
                            "not-separator '*'",
                            'does not affect any separator'])])
def test_row_split_cfg_na_nok(capsys: CaptureFixture[str], splitr: object,
                              msgs: list[str]) -> None:
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListTransfName()
    cfg1.s01_split_rows = cast(RuleRowSplit[str], deepcopy(splitr))
    with pytest.raises(InvalidConfiguration):
        _ = cfg1.as_json_string(stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert '' == out
    assert 'Invalid configuration' in err
    for msg in msgs:
        assert msg in err


@pytest.mark.parametrize('merger,msgs',
                         [([{'columns': [], 'separator': ' '}],
                           ['s02_merge_rows[0][columns]',
                            'less than minimum 1']),
                          ([{'columns': ['foo', 'bar'], 'separator': [' ']}],
                           ['s02_merge_rows[0][separator]',
                            'not of type str']),
                          ([{'columns': ['foo', 'bar'], 'separator': ' ',
                             'split': 2}],
                           ["Unknown key 'split' in s02_merge_rows[0]"]),
                          ([{'columns': [1, 2], 'separator': ' '}],
                           ['s02_merge_rows[0][columns][0]',
                            'not of type str']),
                          ([{'columns': ['foo', 'bar'], 'separator': 3}],
                           ['s02_merge_rows[0][separator]',
                            'not of type str']),
                          ([{'columns': 'foo', 'separator': ' '}],
                           ['s02_merge_rows[0][columns]',
                            'is not a list'])])
def test_row_merge_cfg_na_nok(capsys: CaptureFixture[str], merger: object,
                              msgs: list[str]) -> None:
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListTransfName()
    cfg1.s02_merge_rows = cast(RuleMerge[str], deepcopy(merger))
    with pytest.raises(InvalidConfiguration):
        _ = cfg1.as_json_string(stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert '' == out
    assert 'Invalid configuration' in err
    for msg in msgs:
        assert msg in err
