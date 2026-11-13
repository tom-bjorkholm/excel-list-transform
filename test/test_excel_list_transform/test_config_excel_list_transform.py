#! /usr/local/bin/python3
"""Test current excel-list-transform configuration behavior."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from copy import deepcopy
from io import StringIO
import json
import sys
import pytest
from pytest import CaptureFixture
from config_as_json import ConfigBadJson, InvalidConfiguration
from tableio import TableBorderStyle
from test_excel_list_transform.tableio_helpers import \
    configure_input_csv, configure_output_csv
import excel_list_transform.config_excel_list_transform as config_mod
from excel_list_transform.config_enums import ColumnRef, EmptyParts, \
    EndSeparators
from excel_list_transform.config_read_old import ConfigReadOld
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform, ColInfo
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.migrate_cfg_warn_hook import EltMigrateCfgWarnHook

OUTER_FIRST_STYLE = TableBorderStyle.OUTER_FIRST_ROW_THICK_INNER_THIN


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


@pytest.mark.parametrize('output_borders, border_style',
                         [(True, OUTER_FIRST_STYLE),
                          (False, TableBorderStyle.NONE)])
def test_table_border_style(output_borders: bool,
                            border_style: TableBorderStyle) -> None:
    """Test mapping from output border flag to TableIO border style."""
    cfg = ConfigXlsListTransfName()
    cfg.output_borders = output_borders
    assert cfg.table_border_style() == border_style


def _project_rule_columns(member_value: object,
                          stderr_file: StringIO) -> object:
    """Project single-column rules through the private implementation."""
    cfg = ConfigXlsListTransfName()
    # pylint: disable-next=protected-access
    return config_mod._project_rule_columns(cfg, 's03_split_columns',
                                            member_value, stderr_file)


def _project_merge_columns(member_value: object,
                           stderr_file: StringIO) -> object:
    """Project merge-column rules through the private implementation."""
    cfg = ConfigXlsListTransfName()
    # pylint: disable-next=protected-access
    return config_mod._project_merge_columns(cfg, 's05_merge_columns',
                                             member_value, stderr_file)


@pytest.mark.parametrize('member_value, msg',
                         [('bad', 'must be a list.'),
                          ([2], 'non-object rule at index 0'),
                          ([{'name': 'First'}],
                           'misses key column at index 0')])
def test_proj_rule_cols_bad(member_value: object, msg: str) -> None:
    """Test invalid single-column projection values."""
    stderr_file = StringIO()
    with pytest.raises(InvalidConfiguration) as exc_info:
        _project_rule_columns(member_value, stderr_file)
    assert msg in str(exc_info.value)
    assert msg in stderr_file.getvalue()


@pytest.mark.parametrize('member_value, msg',
                         [('bad', 'must be a list.'),
                          ([2], 'non-object rule at index 0'),
                          ([{'column': 'Name'}],
                           'misses key columns at index 0'),
                          ([{'columns': 'Name'}],
                           'has non-list columns at index 0')])
def test_proj_merge_cols_bad(member_value: object, msg: str) -> None:
    """Test invalid merge-column projection values."""
    stderr_file = StringIO()
    with pytest.raises(InvalidConfiguration) as exc_info:
        _project_merge_columns(member_value, stderr_file)
    assert msg in str(exc_info.value)
    assert msg in stderr_file.getvalue()


@pytest.mark.parametrize('member_value, msg',
                         [(['bad'], 'must be a dict.'),
                          ({'separators': '+', 'not_separators': []},
                           'must contain separator lists.')])
def test_split_row_val_bad(member_value: object, msg: str) -> None:
    """Test split-row separator validation for prevalidated bad shapes."""
    cfg = ConfigXlsListTransfName()
    stderr_file = StringIO()
    validator = config_mod.SplitRowSepValidator()
    with pytest.raises(InvalidConfiguration) as exc_info:
        validator.validate_member(cfg, 's01_split_rows[0]', member_value,
                                  stderr_file)
    assert msg in str(exc_info.value)
    assert msg in stderr_file.getvalue()


def _empty_separator(rules: object, member: str) -> None:
    """Set an empty separator in the first rule of member rules."""
    assert isinstance(rules, list) and isinstance(rules[0], dict)
    if member == 's01_split_rows':
        rules[0]['separators'] = [',', '']
    else:
        rules[0]['separator'] = ''


@pytest.mark.parametrize('cls', [ConfigXlsListTransfName,
                                 ConfigXlsListTransfNum])
@pytest.mark.parametrize('member, path',
                         [('s01_split_rows',
                           's01_split_rows[0][separators][1]'),
                          ('s03_split_columns',
                           's03_split_columns[0][separator]')])
def test_empty_separator_bad(capsys: CaptureFixture[str],
                             cls: type[ConfigXlsListTransfName] |
                             type[ConfigXlsListTransfNum],
                             member: str, path: str) -> None:
    """Test that empty separators are rejected when writing and reading."""
    cfg = cls()
    data = json.loads(cfg.as_json_string(stderr_file=sys.stderr))
    _empty_separator(data[member], member)
    _empty_separator(getattr(cfg, member), member)
    with pytest.raises(InvalidConfiguration):
        _ = cfg.as_json_string(stderr_file=sys.stderr)
    with pytest.raises(InvalidConfiguration):
        _ = cls(from_json_data_text=json.dumps(data), stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert '' == out
    assert err.count(f'{path} has length 0') == 2


SPLIT_CFG_CLASSES = [ConfigXlsListTransfName, ConfigXlsListTransfNum]
SPLIT_MODE_KEYS = [('s01_split_rows', 'empty_parts'),
                   ('s03_split_columns', 'end_separators')]


def _split_mode_json(cls: type[ConfigXlsListTransfName] |
                     type[ConfigXlsListTransfNum], member: str, key: str,
                     value: object) -> str:
    """Return default config JSON with key of first member rule replaced."""
    data = json.loads(cls().as_json_string(stderr_file=sys.stderr))
    data[member][0][key] = value
    return json.dumps(data)


@pytest.mark.parametrize('cls', SPLIT_CFG_CLASSES)
def test_split_mode_defaults(cls: type[ConfigXlsListTransfName] |
                             type[ConfigXlsListTransfNum]) -> None:
    """Test the default values of empty_parts and end_separators."""
    cfg = cls()
    assert cfg.s01_split_rows[0]['empty_parts'] == EmptyParts.DROP_TRAILING
    ends = cfg.s03_split_columns[0]['end_separators']
    assert ends == EndSeparators.STRIP_TRAILING


@pytest.mark.parametrize('cls', SPLIT_CFG_CLASSES)
@pytest.mark.parametrize('member, key, text, expected',
                         [('s01_split_rows', 'empty_parts', 'KEEP_ALL',
                           EmptyParts.KEEP_ALL),
                          ('s01_split_rows', 'empty_parts', 'drop_all',
                           EmptyParts.DROP_ALL),
                          ('s01_split_rows', 'empty_parts', 'DROP_T',
                           EmptyParts.DROP_TRAILING),
                          ('s01_split_rows', 'empty_parts', 'KEEP',
                           EmptyParts.KEEP_ALL),
                          ('s03_split_columns', 'end_separators', 'KEEP',
                           EndSeparators.KEEP),
                          ('s03_split_columns', 'end_separators',
                           'strip_leading', EndSeparators.STRIP_LEADING),
                          ('s03_split_columns', 'end_separators', 'STRIP_T',
                           EndSeparators.STRIP_TRAILING),
                          ('s03_split_columns', 'end_separators', 'Strip_B',
                           EndSeparators.STRIP_BOTH_ENDS)])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_split_mode_json_ok(capsys: CaptureFixture[str],
                            cls: type[ConfigXlsListTransfName] |
                            type[ConfigXlsListTransfNum],
                            member: str, key: str, text: str,
                            expected: EmptyParts | EndSeparators) -> None:
    """Test reading empty_parts and end_separators (unique prefixes)."""
    txt = _split_mode_json(cls, member, key, text)
    cfg = cls(from_json_data_text=txt, stderr_file=sys.stderr)
    assert getattr(cfg, member)[0][key] == expected
    written = json.loads(cfg.as_json_string(stderr_file=sys.stderr))
    assert written[member][0][key] == expected.name
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('cls', SPLIT_CFG_CLASSES)
@pytest.mark.parametrize('member, key, value, msg',
                         [('s01_split_rows', 'empty_parts', 'DROP',
                           'DROP is not one of: KEEP_ALL, DROP_ALL, '
                           'DROP_TRAILING'),
                          ('s01_split_rows', 'empty_parts', 'KEEP_SOME',
                           'KEEP_SOME is not one of'),
                          ('s01_split_rows', 'empty_parts', 2,
                           'called with int not str'),
                          ('s01_split_rows', 'empty_parts', None,
                           'called with NoneType not str'),
                          ('s03_split_columns', 'end_separators', 'STRIP',
                           'STRIP is not one of: KEEP, STRIP_LEADING, '
                           'STRIP_TRAILING, STRIP_BOTH_ENDS'),
                          ('s03_split_columns', 'end_separators', '',
                           ' is not one of'),
                          ('s03_split_columns', 'end_separators', ['KEEP'],
                           'called with list not str')])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_split_mode_json_bad(capsys: CaptureFixture[str],
                             cls: type[ConfigXlsListTransfName] |
                             type[ConfigXlsListTransfNum],
                             member: str, key: str, value: object,
                             msg: str) -> None:
    """Test that invalid empty_parts and end_separators are rejected."""
    txt = _split_mode_json(cls, member, key, value)
    with pytest.raises(ConfigBadJson):
        _ = cls(from_json_data_text=txt, stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert '' == out
    assert msg in err


@pytest.mark.parametrize('cls', SPLIT_CFG_CLASSES)
@pytest.mark.parametrize('member, key', SPLIT_MODE_KEYS)
@pytest.mark.parametrize('value', ['DROP_ALL', 'KEEP', None, 1])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_split_mode_type_bad(capsys: CaptureFixture[str],
                             cls: type[ConfigXlsListTransfName] |
                             type[ConfigXlsListTransfNum],
                             member: str, key: str, value: object) -> None:
    """Test that a value that is not of the enum type is not written."""
    cfg = cls()
    getattr(cfg, member)[0][key] = value
    with pytest.raises(InvalidConfiguration):
        _ = cfg.as_json_string(stderr_file=sys.stderr)
    out, err = capsys.readouterr()
    assert '' == out
    assert f'Value for {member}[0][{key}] is not of type' in err


@pytest.mark.parametrize('cls', SPLIT_CFG_CLASSES)
def test_split_mode_missing(capsys: CaptureFixture[str],
                            cls: type[ConfigXlsListTransfName] |
                            type[ConfigXlsListTransfNum]) -> None:
    """Test that missing keys get the old behavior and a migrate warning."""
    data = json.loads(cls().as_json_string(stderr_file=sys.stderr))
    for member, key in SPLIT_MODE_KEYS:
        for rule in data[member]:
            del rule[key]
    data['s01_split_rows'].append(deepcopy(data['s01_split_rows'][0]))
    cfg = cls(from_json_data_text=json.dumps(data), stderr_file=sys.stderr)
    assert [rule['empty_parts'] for rule in cfg.s01_split_rows] == \
        [EmptyParts.KEEP_ALL, EmptyParts.KEEP_ALL]
    ends = cfg.s03_split_columns[0]['end_separators']
    assert ends == EndSeparators.KEEP
    out, err = capsys.readouterr()
    assert '' == out
    assert EltMigrateCfgWarnHook.migrate_warn_msg() == err


def test_rocf_split_defaults() -> None:
    """Test that old configurations get the old split behavior."""
    missing = ConfigReadOld().get_missing_path_values()
    empty_parts = missing[('s01_split_rows', '[', 'empty_parts')]
    assert empty_parts == EmptyParts.KEEP_ALL
    ends = missing[('s03_split_columns', '[', 'end_separators')]
    assert ends == EndSeparators.KEEP
