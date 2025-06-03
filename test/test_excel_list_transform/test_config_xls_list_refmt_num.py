#! /usr/local/bin/python3
"""Test the ConfigXlsListRefmtNum class."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

import pytest
from excel_list_transform.config_xls_list_refmt_num \
    import ConfigXlsListRefmtNum
from excel_list_transform.config_excel_list_transform \
    import FileType, SplitWhere


@pytest.mark.smoke
def test_config_xls_list_refmt_def(capsys):
    """Test default values of ConfigXlsListRefmtNum."""
    cfg = ConfigXlsListRefmtNum()
    assert isinstance(cfg.s1_split_columns, list)
    assert isinstance(cfg.s2_remove_columns, list)
    assert isinstance(cfg.s3_merge_columns, list)
    assert isinstance(cfg.s4_place_columns_first, list)
    assert len(cfg.s3_merge_columns) > 0
    assert 'columns' in cfg.s3_merge_columns[0]
    assert cfg.s3_merge_columns[0]['columns'] == [15, 16]
    assert cfg.s2_remove_columns == [1, 2, 3]
    assert cfg.s4_place_columns_first == [7, 3, 6]
    assert len(cfg.s1_split_columns) > 0
    assert 'column' in cfg.s1_split_columns[0]
    assert cfg.s1_split_columns[0]['where'] == SplitWhere.RIGHTMOST
    assert cfg.s1_split_columns[0]['column'] == 15
    assert cfg.in_type == FileType.EXCEL
    assert cfg.out_type == FileType.EXCEL
    str_cfg = cfg.as_json_string()
    assert len(str_cfg) > 1
    assert 'in_type' in str_cfg
    zcfg = ConfigXlsListRefmtNum()
    assert cfg.__dict__ == zcfg.__dict__
    ycfg = ConfigXlsListRefmtNum(from_json_text=str_cfg)
    assert ycfg.__dict__ == cfg.__dict__
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t, val',
                         [('{"s2_remove_columns" : [2]}', [2]),
                          ('{"s2_remove_columns" : [21, 14]}', [21, 14])])
def test_config_xls_list_refmt_read_incomplete3(capsys, t, val):
    """Test ConfigXlsListRefmtNum read incomplete 3."""
    ycfg = ConfigXlsListRefmtNum()
    ycfg.parse_json(t, ok_to_use_defaults=True)
    assert ycfg.s2_remove_columns == val
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t',
                         ['{"out_type_" : "CSV"}',
                          '{"outfilen" : "out.dat"}'])
def test_config_xls_list_refmt_read_incomplete4(capsys, t):
    """Test ConfigXlsListRefmtNum read incomplete 4."""
    cfg = ConfigXlsListRefmtNum()
    with pytest.raises(KeyError) as exc_info:
        cfg.parse_json(t, ok_to_use_defaults=True)
    assert exc_info.type is KeyError
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Unexpected' in err


@pytest.mark.parametrize('t, errtxt',
                         [('{"in_csv_dialect": "B"}', 'Not dictionary for'),
                          ('{"s2_remove_columns": {"delimiter" : ";"}}',
                           'Unexpected dictionary for')])
def test_config_xls_list_refmt_read_dict_mismatch(capsys, t, errtxt):
    """Test ConfigXlsListRefmtNum read dict mismatch."""
    cfg = ConfigXlsListRefmtNum()
    with pytest.raises(KeyError) as exc_info:
        cfg.parse_json(t, ok_to_use_defaults=True)
    assert exc_info.type is KeyError
    out, err = capsys.readouterr()
    assert out == ''
    assert errtxt in err
