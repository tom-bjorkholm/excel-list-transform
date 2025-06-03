#! /usr/local/bin/python3
"""Test the ConfigXlsListRefmtName class."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

import pytest
from excel_list_transform.config_xls_list_refmt_name \
    import ConfigXlsListRefmtName
from excel_list_transform.config_excel_list_transform \
    import FileType, SplitWhere


@pytest.mark.smoke
def test_config_xls_list_refmt_def(capsys):
    """Test default values of ConfigXlsListRefmtName."""
    cfg = ConfigXlsListRefmtName()
    assert isinstance(cfg.s1_split_columns, list)
    assert isinstance(cfg.s8_column_order, list)
    assert isinstance(cfg.s3_merge_columns, list)
    assert len(cfg.s3_merge_columns) > 0
    assert 'columns' in cfg.s3_merge_columns[0]
    assert cfg.s3_merge_columns[0]['columns'] == ['street', 'street number']
    assert cfg.s8_column_order == ['Class', 'Division', 'Nationality',
                                   'Sail Number', 'Boat Name', 'First Name',
                                   'Last Name', 'Club Name', 'Email',
                                   'Phone', 'WhatsApp']
    assert len(cfg.s1_split_columns) > 0
    assert 'column' in cfg.s1_split_columns[0]
    assert cfg.s1_split_columns[0]['where'] == SplitWhere.RIGHTMOST
    assert cfg.s1_split_columns[0]['column'] == 'name'
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
                         [('{"s8_column_order" : ["Name"]}', ['Name']),
                          ('{"s8_column_order" : ["ab", "1a"]}',
                           ['ab', '1a'])])
def test_config_xls_list_refmt_read_incomplete3(capsys, t, val):
    """Test ConfigXlsListRefmtName read incomplete 3."""
    ycfg = ConfigXlsListRefmtName()
    ycfg.parse_json(t, ok_to_use_defaults=True)
    assert ycfg.s8_column_order == val
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
                          ('{"s8_column_order": {"delimiter" : ";"}}',
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
