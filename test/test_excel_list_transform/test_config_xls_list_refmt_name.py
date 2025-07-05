#! /usr/local/bin/python3
"""Test the ConfigXlsListRefmtName class."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from copy import deepcopy
import pytest
from excel_list_transform.config_xls_list_refmt_name \
    import ConfigXlsListRefmtName
from excel_list_transform.config_excel_list_transform \
    import FileType, SplitWhere


@pytest.mark.smoke
def test_config_xls_list_refmt_def(capsys):
    """Test default values of ConfigXlsListRefmtName."""
    cfg = ConfigXlsListRefmtName()
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
    assert 'in_type' in str_cfg
    zcfg = ConfigXlsListRefmtName()
    assert cfg.__dict__ == zcfg.__dict__
    ycfg = ConfigXlsListRefmtName(from_json_text=str_cfg)
    assert ycfg.__dict__ == cfg.__dict__
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t, val',
                         [('{"s10_column_order" : ["Name"]}', ['Name']),
                          ('{"s10_column_order" : ["ab", "1a"]}',
                           ['ab', '1a'])])
def test_config_xls_list_refmt_read_incomplete3(capsys, t, val):
    """Test ConfigXlsListRefmtName read incomplete 3."""
    ycfg = ConfigXlsListRefmtName()
    ycfg.parse_json(t, ok_to_use_defaults=True)
    assert ycfg.s10_column_order == val
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t',
                         ['{"out_type_" : "CSV"}',
                          '{"outfilen" : "out.dat"}'])
def test_config_xls_list_refmt_read_incomplete4(capsys, t):
    """Test ConfigXlsListRefmtName read incomplete 4."""
    cfg = ConfigXlsListRefmtName()
    with pytest.raises(KeyError) as exc_info:
        cfg.parse_json(t, ok_to_use_defaults=True)
    assert exc_info.type is KeyError
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Unexpected' in err


@pytest.mark.parametrize('t, errtxt',
                         [('{"in_csv_dialect": "B"}', 'Not dictionary for'),
                          ('{"s10_column_order": {"delimiter" : ";"}}',
                           'Unexpected dictionary for')])
def test_config_xls_list_refmt_read_dict_mismatch(capsys, t, errtxt):
    """Test ConfigXlsListRefmtName read dict mismatch."""
    cfg = ConfigXlsListRefmtName()
    with pytest.raises(KeyError) as exc_info:
        cfg.parse_json(t, ok_to_use_defaults=True)
    assert exc_info.type is KeyError
    out, err = capsys.readouterr()
    assert out == ''
    assert errtxt in err


def test_bak_compat_0_7_13_name(capsys):
    """Test backward compatible reading om 0.7.13 config file."""
    refcfg = ConfigXlsListRefmtName()
    refcfg.out_type = FileType.CSV
    refcfg.in_csv_dialect['name'] = 'csv.unix_dialect'
    refcfg.s01_split_rows = []
    refcfg.s02_merge_rows = []
    refcfg.s03_split_columns[0]['right_name'] = 'Family Name'
    refcfg.s08_insert_columns[1]['column'] = 'Something Else'
    filename = 'test/test_excel_list_transform/bak_compat_0_7_13_name.cfg'
    cfg = ConfigXlsListRefmtName(from_json_filename=filename)
    out, err = capsys.readouterr()
    assert refcfg.__dict__ == cfg.__dict__
    assert cfg.s03_split_columns[0]['right_name'] == 'Family Name'
    assert cfg.s08_insert_columns[1]['column'] == 'Something Else'
    assert '' == out
    assert '' == err


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
def test_row_split_merge_cfg_na_ok(capsys, splitr, merger):
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListRefmtName()
    cfg1.s01_split_rows = deepcopy(splitr)
    cfg1.s02_merge_rows = deepcopy(merger)
    txt = cfg1.as_json_string()
    cfg2 = ConfigXlsListRefmtName(from_json_text=txt)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert cfg1.__dict__ == cfg2.__dict__
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
def test_row_split_cfg_na_nok(capsys, splitr, msgs):
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListRefmtName()
    cfg1.s01_split_rows = deepcopy(splitr)
    txt = cfg1.as_json_string()
    with pytest.raises(SystemExit):
        _ = ConfigXlsListRefmtName(from_json_text=txt)
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
def test_row_merge_cfg_na_nok(capsys, merger, msgs):
    """Test OK cases of row split and merge config."""
    cfg1 = ConfigXlsListRefmtName()
    cfg1.s02_merge_rows = deepcopy(merger)
    txt = cfg1.as_json_string()
    with pytest.raises(SystemExit):
        _ = ConfigXlsListRefmtName(from_json_text=txt)
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err
