#! /usr/local/bin/python3
"""Test current excel-list-transform configuration behavior."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from io import StringIO
import sys
import pytest
from pytest import CaptureFixture
from config_as_json import InvalidConfiguration
from test_excel_list_transform.tableio_helpers import \
    configure_input_csv, configure_output_csv
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.config_read_old import ConfigReadOld
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform, ColInfo
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.migrate_cfg_warn_hook import EltMigrateCfgWarnHook


def _make_base_config(colref: ColumnRef) -> ConfigExcelListTransform[int] | \
        ConfigExcelListTransform[str]:
    """Return a base transform config for the requested column reference."""
    if colref == ColumnRef.BY_NUMBER:
        num_colinfo = ColInfo[int]('right_name', None, [], [],
                                   [2, 3, 0, 1, 4, 4, 4, 4, 4, 1], [7, 1, 2],
                                   2)
        return ConfigExcelListTransform[int](col_ref=colref,
                                             colinfo=num_colinfo, tinfo=2)
    name_colinfo = ColInfo[str](
        'right_name', None, [], [],
        ['street', 'street number', 'name', 'last name',
         'Phone', 'Phone', 'Phone', 'Phone', 'Phone', 'Last Name'],
        ['Club Name', 'name', 'last name'], 'a')
    return ConfigExcelListTransform[str](col_ref=colref, colinfo=name_colinfo,
                                         tinfo='a')


@pytest.mark.parametrize('colref', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
def test_base_config_defaults(capsys: CaptureFixture[str],
                              colref: ColumnRef) -> None:
    """Test default current configuration values."""
    cfg = _make_base_config(colref)
    assert cfg.column_ref == colref
    assert cfg.input_table.format_name == 'Excel'
    assert cfg.output_table.format_name == 'Excel'
    txt = cfg.as_json_string(stderr_file=sys.stderr)
    assert 'input_table' in txt
    assert 'output_table' in txt
    assert 'in_type' not in txt
    assert 'out_type' not in txt
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


def test_old_default_values() -> None:
    """Test current-shape defaults inserted after old migration."""
    data = ConfigReadOld().process_json(json_data={'in_type': 'EXCEL'},
                                        auto_ch_hook=EltMigrateCfgWarnHook(),
                                        stderr_file=StringIO())
    input_table = data['input_table']
    output_table = data['output_table']
    assert isinstance(input_table, dict)
    assert isinstance(output_table, dict)
    assert input_table['format_name'] == 'Excel'
    assert output_table['format_name'] == 'Excel'
    assert input_table['character_encoding'] == 'utf_8_sig'
    assert output_table['character_encoding'] == 'utf-8'
    assert not data['s01_split_rows']
    assert not data['s02_merge_rows']


def test_missing_defaults() -> None:
    """Test safe defaults inserted for missing current-shape values."""
    data = ConfigReadOld().process_json(json_data={
        'input_table': {'format_name': 'CSV'}, 'output_table': {}},
        auto_ch_hook=EltMigrateCfgWarnHook(), stderr_file=StringIO())
    input_table = data['input_table']
    output_table = data['output_table']
    assert isinstance(input_table, dict)
    assert isinstance(output_table, dict)
    assert input_table['format_name'] == 'CSV'
    assert output_table['format_name'] == 'Excel'
    assert input_table['character_encoding'] == 'utf_8_sig'
    assert output_table['character_encoding'] == 'utf-8'
    assert not data['s01_split_rows']
    assert not data['s02_merge_rows']


@pytest.mark.parametrize('config_class, filename', [
    (ConfigXlsListTransfName,
     'test/test_excel_list_transform/bak_compat_0_7_13_name.cfg'),
    (ConfigXlsListTransfNum,
     'test/test_excel_list_transform/bak_compat_0_7_13_number.cfg')])
def test_old_config_warns(config_class: type[ConfigXlsListTransfName] |
                          type[ConfigXlsListTransfNum],
                          filename: str) -> None:
    """Test old configuration files read into the current shape."""
    err_file = StringIO()
    hook = EltMigrateCfgWarnHook()
    cfg = config_class(from_json_filename=filename, auto_ch_hook=hook,
                       stderr_file=err_file)
    assert cfg.input_table.format_name == 'Excel'
    assert cfg.output_table.format_name == 'CSV'
    assert cfg.input_table.character_encoding == 'utf_8_sig'
    assert cfg.output_table.character_encoding == 'utf-8'
    assert 'Backward compatibility' in err_file.getvalue()


@pytest.mark.parametrize('encoding', ['utf-8', 'iso8859-1'])
def test_table_encoding(encoding: str) -> None:
    """Test TableIO encoding values are stored in nested config."""
    cfg = ConfigXlsListTransfName()
    configure_input_csv(cfg, encoding=encoding)
    configure_output_csv(cfg, encoding=encoding)
    txt = cfg.as_json_string(stderr_file=sys.stderr)
    assert encoding in txt
    reread = ConfigXlsListTransfName(from_json_data_text=txt)
    assert reread.input_table.character_encoding == encoding
    assert reread.output_table.character_encoding == encoding


def test_bad_table_encoding() -> None:
    """Test invalid TableIO encoding is rejected during validation."""
    cfg = ConfigXlsListTransfName()
    configure_input_csv(cfg, encoding='not-a-real-encoding')
    with pytest.raises(InvalidConfiguration):
        cfg.as_json_string(stderr_file=sys.stderr)
