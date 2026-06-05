#! /usr/local/bin/python3
"""Read and write table files using TableIO."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from typing import cast
from tableio import FileAccess, tio_config_create
from excel_list_transform.commontypes import NameData, NameDataMap, \
    NumData, NumDataSeq
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform, input_capabilities, output_capabilities
from excel_list_transform.handle_empty_column import handle_empty_column_in_lst
from excel_list_transform.num_named_conversion import named_cols_from_num_cols


def _allow_overwrite(filename: str) -> None:
    """Ask the user whether TableIO may overwrite an existing file."""
    print(f'Output file {filename} already exists.')
    answer = input('Overwrite it? [y/N] ')
    if answer.strip().lower() not in ['y', 'yes']:
        raise FileExistsError(f'Output file already exists: {filename}')


def _is_excel_input(cfg: ConfigExcelListTransform[int] |
                    ConfigExcelListTransform[str]) -> bool:
    """Return whether the input TableIO format is Excel."""
    return cfg.input_table.format_name.lower() == 'excel'


def _trim_table(data: NumDataSeq, max_column_read: int) -> NumData:
    """Return table data limited to the configured number of columns."""
    ret: NumData = []
    for row in data:
        ret.append(list(row)[:max_column_read])
    return ret


def _strip_table(data: NumData, strip_col_names: bool,
                 strip_values: bool) -> NumData:
    """Strip whitespace from title row and values when configured."""
    if not strip_col_names and not strip_values:
        return data
    for index, value in enumerate(data[0]):
        if strip_col_names and isinstance(value, str):
            data[0][index] = value.strip()
    for row in data[1:]:
        for index, value in enumerate(row):
            if strip_values and isinstance(value, str):
                row[index] = value.strip()
    return data


def read_table_num(filename: str,
                   cfg: ConfigExcelListTransform[int] |
                   ConfigExcelListTransform[str]) -> NumData:
    """Read table file as numbered-column data."""
    with tio_config_create(config=cfg.input_table, file_name=filename,
                           file_access=FileAccess.READ,
                           capabilities=input_capabilities()) as reader:
        read_result = reader.read_table_listdata()
    data = _trim_table(cast(NumDataSeq, read_result.data),
                       max_column_read=cfg.max_column_read)
    data = handle_empty_column_in_lst(data)
    if _is_excel_input(cfg):
        data = _strip_table(data, cfg.in_excel_col_name_strip,
                            cfg.in_excel_values_strip)
    return data


def read_table_named(filename: str,
                     cfg: ConfigExcelListTransform[int] |
                     ConfigExcelListTransform[str]) -> NameData:
    """Read table file as named-column data."""
    data = read_table_num(filename=filename, cfg=cfg)
    return named_cols_from_num_cols(data=data, filename=filename)


def write_table_num(data: NumData, filename: str,
                    cfg: ConfigExcelListTransform[int] |
                    ConfigExcelListTransform[str]) -> str:
    """Write numbered-column data to a table file."""
    written_file = ''
    with tio_config_create(config=cfg.output_table, file_name=filename,
                           file_access=FileAccess.CREATE,
                           capabilities=output_capabilities(),
                           file_exists_callback=_allow_overwrite) as writer:
        writer.write_table_listdata(
            data=data, filtered_data_range=cfg.output_filtered_table,
            border_style=cfg.table_border_style())
        written_file = writer.file_name
    return written_file


def write_table_named(data: NameDataMap, filename: str,
                      cfg: ConfigExcelListTransform[int] |
                      ConfigExcelListTransform[str],
                      column_order: list[str]) -> str:
    """Write named-column data to a table file."""
    written_file = ''
    with tio_config_create(config=cfg.output_table, file_name=filename,
                           file_access=FileAccess.CREATE,
                           capabilities=output_capabilities(),
                           file_exists_callback=_allow_overwrite) as writer:
        writer.write_table_dictdata(
            data=data, column_order=column_order, extra_ok=True,
            filtered_data_range=cfg.output_filtered_table,
            border_style=cfg.table_border_style())
        written_file = writer.file_name
    return written_file
