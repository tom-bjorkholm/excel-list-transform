#! /usr/local/bin/python3
"""Test reading backward compatible 0.8.5 configuration files."""

# Copyright (c) 2026 Tom Bjorkholm
# MIT License

from collections.abc import Sequence
from io import StringIO
from pathlib import Path
from typing import NamedTuple, Optional
import pytest
from config_as_json import config_factory_from_json
from tableio import CsvDialect
from tableio_cfg_json import TioJsonCsvConfig
from excel_list_transform.config_enums import ColumnRef, RewriteKind, \
    SplitWhere
from excel_list_transform.config_match import MATCH_CONFIGS
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.migrate_cfg_warn_hook import EltMigrateCfgWarnHook


LEGACY_DIR = Path('test/data/backward/0.8.5')
RRS_ORDER = ('Class', 'Division', 'Nationality', 'Sail Number',
             'Boat Name', 'First Name', 'Last Name', 'Club Name',
             'Email', 'Phone', 'WhatsApp')
SW_ORDER = ('Class', 'Division', 'Nat', 'SailNo', 'Boat', 'HelmName',
            'Club', 'HelmEmail', 'HelmPhone')
FORM_REMOVE_COLS = [0, 1, 2, 3, 4, 5, 14, 15, 16, 17, 18, 19]


class ReadResult(NamedTuple):
    """Configuration and warning captured while reading one legacy file."""

    cfg: ConfigXlsListTransfName | ConfigXlsListTransfNum
    warning: str


class IoExpectation(NamedTuple):
    """Expected migrated TableIO values for one legacy file."""

    in_format: str
    out_format: str
    in_encoding: str
    out_encoding: str
    in_delimiter: str
    out_delimiter: str
    in_dialect: CsvDialect
    out_dialect: CsvDialect


class RuleExpectation(NamedTuple):
    """Expected distinctive transform values for one legacy file."""

    row_split_column: Optional[str | int] = None
    row_merge_columns: Sequence[str | int] = ()
    split_column: Optional[str | int] = None
    split_right_name: Optional[str] = None
    split_store_single: Optional[SplitWhere] = None
    merge_columns: Sequence[str | int] = ()
    remove_columns: Sequence[int] = ()
    place_first: Sequence[int] = ()
    rename_rule: Optional[tuple[str | int, str]] = None
    insert_rule: Optional[tuple[str | int, Optional[str], Optional[str]]] = \
        None
    rewrite_rule: Optional[tuple[str | int, RewriteKind]] = None
    rewrite_replace: Optional[tuple[str | int, str, str]] = None
    column_order: Sequence[str] = ()


class ConfigExpectation(NamedTuple):
    """Expected values for one backward compatible 0.8.5 config file."""

    filename: str
    config_class: type[object]
    column_ref: ColumnRef
    io: IoExpectation
    rules: RuleExpectation


EXCEL_TO_CSV_IO = IoExpectation('Excel', 'CSV', 'utf_8_sig', 'utf-8', ',', ',',
                                CsvDialect.EXCEL, CsvDialect.EXCEL)
EXCEL_TO_EXCEL_IO = IoExpectation('Excel', 'Excel', 'utf_8_sig', 'utf-8', ',',
                                  ',', CsvDialect.EXCEL, CsvDialect.EXCEL)
SA_TO_RRS_IO = IoExpectation('CSV', 'Excel', 'cp1252', 'utf-8', ';', ',',
                             CsvDialect.EXCEL, CsvDialect.EXCEL)
EXAMPLE_IO = IoExpectation('Excel', 'CSV', 'utf_8_sig', 'utf-8', ',', ',',
                           CsvDialect.UNIX, CsvDialect.EXCEL)


def _name_config(filename: str, io: IoExpectation,
                 rules: RuleExpectation) -> ConfigExpectation:
    """Return expectation data for a name-based legacy config."""
    return ConfigExpectation(filename, ConfigXlsListTransfName,
                             ColumnRef.BY_NAME, io, rules)


def _number_config(filename: str, io: IoExpectation,
                   rules: RuleExpectation) -> ConfigExpectation:
    """Return expectation data for a number-based legacy config."""
    return ConfigExpectation(filename, ConfigXlsListTransfNum,
                             ColumnRef.BY_NUMBER, io, rules)


EXPECTED_CONFIGS = [
    _name_config('example_X_by_name.cfg', EXAMPLE_IO,
                 RuleExpectation(row_split_column='Club Name',
                                 row_merge_columns=['name', 'last name'],
                                 split_column='name',
                                 split_right_name='last name',
                                 merge_columns=['street', 'street number'],
                                 insert_rule=('Other', None, 'some text'),
                                 rewrite_replace=('Last Name', 'donald',
                                                  'duck'),
                                 column_order=RRS_ORDER)),
    _number_config('example_X_by_number.cfg', EXAMPLE_IO,
                   RuleExpectation(row_split_column=7,
                                   row_merge_columns=[1, 2], split_column=15,
                                   split_store_single=SplitWhere.LEFTMOST,
                                   merge_columns=[15, 16],
                                   remove_columns=[1, 2, 3],
                                   place_first=[7, 3, 6],
                                   insert_rule=(7, 'Other', 'some text'),
                                   rewrite_replace=(6, 'donald', 'duck'))),
    _name_config('forms_to_rrs_X_by_name.cfg', EXCEL_TO_EXCEL_IO,
                 RuleExpectation(rename_rule=('Mobiltelefonnummer', 'Phone'),
                                 insert_rule=('WhatsApp', None, None),
                                 rewrite_replace=('Phone', '^4607', '+467'),
                                 column_order=RRS_ORDER)),
    _number_config('forms_to_rrs_X_by_number.cfg', EXCEL_TO_EXCEL_IO,
                   RuleExpectation(remove_columns=FORM_REMOVE_COLS,
                                   place_first=[5, 6, 7, 0, 1, 4, 2, 3],
                                   rename_rule=(7, 'Phone'),
                                   insert_rule=(10, 'WhatsApp', None),
                                   rewrite_replace=(9, '^4607', '+467'))),
    _name_config('forms_to_sw_X_by_name.cfg', EXCEL_TO_CSV_IO,
                 RuleExpectation(merge_columns=['Förnamn', 'Efternamn'],
                                 rename_rule=('Mobiltelefonnummer',
                                              'HelmPhone'),
                                 insert_rule=('Boat', None, None),
                                 rewrite_replace=('HelmPhone', '^4607',
                                                  '+467'),
                                 column_order=SW_ORDER)),
    _number_config('forms_to_sw_X_by_number.cfg', EXCEL_TO_CSV_IO,
                   RuleExpectation(merge_columns=[0, 1],
                                   remove_columns=FORM_REMOVE_COLS,
                                   place_first=[4, 5, 6, 0, 3, 1, 2],
                                   rename_rule=(6, 'HelmPhone'),
                                   insert_rule=(4, 'Boat', None),
                                   rewrite_replace=(8, '^4607', '+467'))),
    _name_config('row_split_merge_X_by_name.cfg', EXCEL_TO_EXCEL_IO,
                 RuleExpectation(row_split_column='To',
                                 row_merge_columns=['To'],
                                 column_order=['To', 'What', 'From'])),
    _number_config('row_split_merge_X_by_number.cfg', EXCEL_TO_EXCEL_IO,
                   RuleExpectation(row_split_column=2, row_merge_columns=[2],
                                   place_first=[2, 1])),
    _name_config('rrs_to_sw_X_by_name.cfg', EXCEL_TO_CSV_IO,
                 RuleExpectation(merge_columns=['First Name', 'Last Name'],
                                 rename_rule=('Boat Name', 'Boat'),
                                 rewrite_replace=('HelmPhone', '^\\+\\+',
                                                  '+'),
                                 column_order=SW_ORDER)),
    _number_config('rrs_to_sw_X_by_number.cfg', EXCEL_TO_CSV_IO,
                   RuleExpectation(merge_columns=[5, 6], remove_columns=[10],
                                   rename_rule=(8, 'HelmPhone'),
                                   rewrite_replace=(8, '^\\+H', 'H'))),
    _name_config('sa_to_rrs_X_by_name.cfg', SA_TO_RRS_IO,
                 RuleExpectation(rewrite_replace=('Phone', '^4607', '+467'),
                                 column_order=RRS_ORDER)),
    _number_config('sa_to_rrs_X_by_number.cfg', SA_TO_RRS_IO,
                   RuleExpectation(rewrite_replace=(9, '^4607', '+467'))),
    _name_config('sw_to_rrs_X_by_name.cfg', EXCEL_TO_EXCEL_IO,
                 RuleExpectation(split_column='Name',
                                 split_right_name='Last Name',
                                 rename_rule=('Name', 'First Name'),
                                 insert_rule=('WhatsApp', None, None),
                                 rewrite_replace=('Phone', '^4607', '+467'),
                                 column_order=RRS_ORDER)),
    _number_config('sw_to_rrs_X_by_number.cfg', EXCEL_TO_EXCEL_IO,
                   RuleExpectation(split_column=5,
                                   split_store_single=SplitWhere.RIGHTMOST,
                                   rename_rule=(6, 'Last Name'),
                                   insert_rule=(10, 'WhatsApp', None),
                                   rewrite_replace=(9, '^4607', '+467')))]


def _read_config(filename: str) -> ReadResult:
    """Read one 0.8.5 config and return the captured warning text."""
    err_file = StringIO()
    cfg = config_factory_from_json(match_configs=MATCH_CONFIGS,
                                   auto_ch_hook=EltMigrateCfgWarnHook(),
                                   from_json_filename=LEGACY_DIR / filename,
                                   stderr_file=err_file)
    assert isinstance(cfg, (ConfigXlsListTransfName, ConfigXlsListTransfNum))
    return ReadResult(cfg=cfg, warning=err_file.getvalue())


def _assert_moved_io(cfg: ConfigXlsListTransfName |
                     ConfigXlsListTransfNum,
                     expected: IoExpectation) -> None:
    """Assert old flat I/O keys were migrated into nested TableIO config."""
    input_csv = cfg.input_table.csv
    output_csv = cfg.output_table.csv
    assert isinstance(input_csv, TioJsonCsvConfig)
    assert isinstance(output_csv, TioJsonCsvConfig)
    assert cfg.input_table.format_name == expected.in_format
    assert cfg.output_table.format_name == expected.out_format
    assert cfg.input_table.character_encoding == expected.in_encoding
    assert cfg.output_table.character_encoding == expected.out_encoding
    assert cfg.input_table.implementation is None
    assert cfg.output_table.implementation is None
    assert input_csv.delimiter == expected.in_delimiter
    assert output_csv.delimiter == expected.out_delimiter
    assert input_csv.dialect == expected.in_dialect
    assert output_csv.dialect == expected.out_dialect


def _assert_rule_values(cfg: ConfigXlsListTransfName |
                        ConfigXlsListTransfNum,
                        expected: RuleExpectation) -> None:
    """Assert distinctive migrated transform values for one config."""
    if expected.row_split_column is not None:
        assert cfg.s01_split_rows[0]['column'] == expected.row_split_column
    if expected.row_merge_columns:
        assert cfg.s02_merge_rows[0]['columns'] == \
            list(expected.row_merge_columns)
    _assert_split_rule(cfg, expected)
    if expected.merge_columns:
        assert cfg.s05_merge_columns[0]['columns'] == \
            list(expected.merge_columns)
    if expected.remove_columns:
        assert isinstance(cfg, ConfigXlsListTransfNum)
        assert cfg.s04_remove_columns == list(expected.remove_columns)
    if expected.place_first:
        assert isinstance(cfg, ConfigXlsListTransfNum)
        assert cfg.s06_place_columns_first == list(expected.place_first)
    _assert_named_rules(cfg, expected)


def _assert_split_rule(cfg: ConfigXlsListTransfName |
                       ConfigXlsListTransfNum,
                       expected: RuleExpectation) -> None:
    """Assert one expected split-column rule if the config has one."""
    if expected.split_column is None:
        return
    split_rule = cfg.s03_split_columns[0]
    assert split_rule['column'] == expected.split_column
    assert split_rule['where'] == SplitWhere.RIGHTMOST
    if expected.split_right_name is not None:
        assert split_rule['right_name'] == expected.split_right_name
    if expected.split_store_single is not None:
        assert split_rule['store_single'] == expected.split_store_single


def _assert_named_rules(cfg: ConfigXlsListTransfName |
                        ConfigXlsListTransfNum,
                        expected: RuleExpectation) -> None:
    """Assert rename, insert, rewrite, and output-order expectations."""
    if expected.rename_rule is not None:
        column, name = expected.rename_rule
        assert any(rule['column'] == column and rule['name'] == name
                   for rule in cfg.s07_rename_columns)
    if expected.insert_rule is not None:
        _assert_insert_rule(cfg, expected.insert_rule)
    if expected.rewrite_rule is not None:
        column, kind = expected.rewrite_rule
        assert any(rule['column'] == column and rule['kind'] == kind
                   for rule in cfg.s09_rewrite_columns)
    if expected.rewrite_replace is not None:
        column, from_text, to_text = expected.rewrite_replace
        assert any(rule['column'] == column and rule.get('from') == from_text
                   and rule.get('to') == to_text
                   for rule in cfg.s09_rewrite_columns)
    if expected.column_order:
        assert isinstance(cfg, ConfigXlsListTransfName)
        assert cfg.s10_column_order == list(expected.column_order)


def _assert_insert_rule(cfg: ConfigXlsListTransfName |
                        ConfigXlsListTransfNum,
                        expected: tuple[str | int, Optional[str],
                                        Optional[str]]) -> None:
    """Assert one expected insert-column rule exists."""
    column, name, value = expected
    matches = [rule for rule in cfg.s08_insert_columns
               if rule['column'] == column]
    assert matches
    rule = matches[0]
    if name is None:
        assert 'name' not in rule
    else:
        assert rule['name'] == name
    assert rule['value'] == value


@pytest.mark.parametrize('expected', EXPECTED_CONFIGS,
                         ids=[item.filename for item in EXPECTED_CONFIGS])
def test_read_0_8_5_config(expected: ConfigExpectation) -> None:
    """Test that old 0.8.5 configs read and warn about migration."""
    result = _read_config(expected.filename)
    assert isinstance(result.cfg, expected.config_class)
    assert result.cfg.column_ref == expected.column_ref
    assert result.warning == EltMigrateCfgWarnHook.migrate_warn_msg()
    _assert_moved_io(result.cfg, expected.io)
    _assert_rule_values(result.cfg, expected.rules)
