#! /usr/local/bin/python3
"""Read old excel-list-transform configuration files."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from enum import Enum
from typing import Optional, override
from config_as_json import ConfigPath, ReadOldConfiguration, RocfKeyMove, \
    RocfKeyRename
from tableio import CsvDialect


def _old_value_name(value: object) -> str:
    """Return a normalized old configuration enum/string name."""
    if isinstance(value, Enum):
        return value.name
    assert isinstance(value, str)
    return value


def old_file_type_to_format(value: object) -> str:
    """Convert old input/output file type to a TableIO format name."""
    name = _old_value_name(value).lower()
    if name == 'csv':
        return 'CSV'
    if name == 'excel':
        return 'Excel'
    raise KeyError(f'Unknown old file type: {value}')


def _old_quoting_to_tableio(value: object) -> Optional[str]:
    """Convert old csv.QUOTE_* config text to TableIO text."""
    if value is None:
        return None
    assert isinstance(value, str)
    values = {'csv.quote_all': 'all',
              'csv.quote_minimal': 'minimal',
              'csv.quote_none': 'none',
              'csv.quote_nonnumeric': 'nonnumeric'}
    lower = value.lower()
    if lower not in values:
        raise KeyError(f'Unknown old CSV quoting: {value}')
    return values[lower]


def old_csv_spec_to_tableio(value: object) -> dict[str, object]:
    """Convert old CSV dialect config to tableio-cfg-json shape."""
    assert isinstance(value, dict)
    name_obj = value.get('name')
    name = '' if name_obj is None else str(name_obj).lower()
    dialect = CsvDialect.EXCEL.name
    if name == 'csv.unix_dialect':
        dialect = CsvDialect.UNIX.name
    ret: dict[str, object] = {'dialect': dialect}
    delimiter = value.get('delimiter')
    if delimiter is None and name == 'csv.excel_tab':
        delimiter = '\t'
    if delimiter is not None:
        ret['delimiter'] = delimiter
    quoting = _old_quoting_to_tableio(value.get('quoting'))
    if quoting is not None:
        ret['quoting'] = quoting
    for old_key in ['quotechar', 'lineterminator', 'escapechar']:
        item = value.get(old_key)
        if item is not None:
            ret[old_key] = item
    return ret


class ConfigReadOld(ReadOldConfiguration):
    """Normalize old excel-list-transform configuration files."""

    @override
    def get_missing_path_values(self) -> dict[ConfigPath, object]:
        """Return safe current-schema defaults for missing config paths."""
        return {('input_table', 'format_name'): 'Excel',
                ('output_table', 'format_name'): 'Excel',
                ('input_table', 'character_encoding'): 'utf_8_sig',
                ('output_table', 'character_encoding'): 'utf-8',
                ('in_excel_col_name_strip',): False,
                ('in_excel_values_strip',): False,
                ('s01_split_rows',): [],
                ('s02_merge_rows',): [],
                ('output_borders',): False,
                ('output_filtered_table',): False}

    @override
    def get_keys_to_prune(self) -> list[str]:
        """Return old keys accepted and discarded during migration."""
        return ['in_excel_library', 'out_excel_library']

    @override
    def get_json_key_renames(self) -> list[RocfKeyRename]:
        """Return old transform rule key renames."""
        return [
            RocfKeyRename(old='s1_split_columns', new='s03_split_columns'),
            RocfKeyRename(old='s2_remove_columns', new='s04_remove_columns'),
            RocfKeyRename(old='s3_merge_columns', new='s05_merge_columns'),
            RocfKeyRename(old='s4_place_columns_first',
                          new='s06_place_columns_first'),
            RocfKeyRename(old='s5_rename_columns', new='s07_rename_columns'),
            RocfKeyRename(old='s6_insert_columns', new='s08_insert_columns'),
            RocfKeyRename(old='s7_rewrite_columns', new='s09_rewrite_columns'),
            RocfKeyRename(old='s8_column_order', new='s10_column_order')
        ]

    @override
    def get_json_key_moves(self) -> list[RocfKeyMove]:
        """Return old I/O settings moved into TableIO config sections."""
        return [
            RocfKeyMove(old_path=('in_type',),
                        new_path=('input_table', 'format_name'),
                        transform_value=old_file_type_to_format),
            RocfKeyMove(old_path=('out_type',),
                        new_path=('output_table', 'format_name'),
                        transform_value=old_file_type_to_format),
            RocfKeyMove(old_path=('in_csv_encoding',),
                        new_path=('input_table', 'character_encoding')),
            RocfKeyMove(old_path=('out_csv_encoding',),
                        new_path=('output_table', 'character_encoding')),
            RocfKeyMove(old_path=('in_csv_dialect',),
                        new_path=('input_table', 'csv'),
                        transform_value=old_csv_spec_to_tableio),
            RocfKeyMove(old_path=('out_csv_dialect',),
                        new_path=('output_table', 'csv'),
                        transform_value=old_csv_spec_to_tableio)
        ]
