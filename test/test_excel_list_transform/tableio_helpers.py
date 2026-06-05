#! /usr/local/bin/python3
"""Test helpers for creating and reading table files with TableIO."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from tableio import CsvDialect, FileAccess, tio_config_create
from tableio_cfg_json import TioJsonConfig, TioJsonCsvConfig
from excel_list_transform.commontypes import NameDataMap, NumData, NumDataSeq
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform, output_capabilities, output_table_factory
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.handle_tableio import read_table_num
from excel_list_transform.num_named_conversion import num_cols_from_named_cols


# python-layout disable-next=more-fits
def configure_input_csv(cfg: ConfigExcelListTransform[int] |
                        ConfigExcelListTransform[str],
                        encoding: str = 'utf_8_sig',
                        delimiter: str = ',') -> None:
    """Configure an application config to read CSV input."""
    cfg.input_table.format_name = 'CSV'
    cfg.input_table.character_encoding = encoding
    cfg.input_table.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                           delimiter=delimiter, quotechar='"')


# python-layout disable-next=more-fits
def configure_output_csv(cfg: ConfigExcelListTransform[int] |
                         ConfigExcelListTransform[str],
                         encoding: str = 'utf-8',
                         delimiter: str = ',') -> None:
    """Configure an application config to write CSV output."""
    cfg.output_table.format_name = 'CSV'
    cfg.output_table.character_encoding = encoding
    cfg.output_table.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                            delimiter=delimiter, quotechar='"')


def configure_input_excel(cfg: ConfigExcelListTransform[int] |
                          ConfigExcelListTransform[str]) -> None:
    """Configure an application config to read Excel input."""
    cfg.input_table.format_name = 'Excel'


def configure_output_excel(cfg: ConfigExcelListTransform[int] |
                           ConfigExcelListTransform[str]) -> None:
    """Configure an application config to write Excel output."""
    cfg.output_table.format_name = 'Excel'


def write_csv_num(data: NumData, filename: str,
                  encoding: str = 'utf-8') -> None:
    """Write numbered rows to a CSV file through TableIO."""
    cfg = output_table_factory()
    cfg.format_name = 'CSV'
    cfg.character_encoding = encoding
    cfg.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL, delimiter=',',
                               quotechar='"')
    _write_num(data=data, filename=filename, cfg=cfg)


def write_csv_named(data: NameDataMap, filename: str, column_order: list[str],
                    encoding: str = 'utf-8') -> None:
    """Write named rows to a CSV file through TableIO."""
    write_csv_num(data=num_cols_from_named_cols(data=data,
                                                column_order=column_order),
                  filename=filename, encoding=encoding)


def write_excel_num(data: NumData | NumDataSeq, filename: str) -> None:
    """Write numbered rows to an Excel file through TableIO."""
    cfg = output_table_factory()
    cfg.format_name = 'Excel'
    _write_num(data=list(data), filename=filename, cfg=cfg)


def write_excel_named(data: NameDataMap, filename: str,
                      column_order: list[str]) -> None:
    """Write named rows to an Excel file through TableIO."""
    write_excel_num(data=num_cols_from_named_cols(data=data,
                                                  column_order=column_order),
                    filename=filename)


def read_excel_num(filename: str, max_col_read: int, strip_col_names: bool,
                   strip_values: bool) -> NumData:
    """Read numbered rows from an Excel file through the app reader."""
    cfg = ConfigXlsListTransfNum()
    configure_input_excel(cfg)
    cfg.max_column_read = max_col_read
    cfg.in_excel_col_name_strip = strip_col_names
    cfg.in_excel_values_strip = strip_values
    return read_table_num(filename=filename, cfg=cfg)


def read_csv_num(filename: str, max_col_read: int,
                 encoding: str = 'utf-8') -> NumData:
    """Read numbered rows from a CSV file through the app reader."""
    cfg = ConfigXlsListTransfNum()
    configure_input_csv(cfg, encoding=encoding)
    cfg.max_column_read = max_col_read
    return read_table_num(filename=filename, cfg=cfg)


def _write_num(data: NumData | NumDataSeq, filename: str,
               cfg: TioJsonConfig) -> None:
    """Write numbered rows using one TableIO output config."""
    with tio_config_create(config=cfg, file_name=filename,
                           file_access=FileAccess.CREATE,
                           capabilities=output_capabilities(),
                           file_exists_callback=_allow_overwrite) as writer:
        writer.write_table_listdata(data=data)


def _allow_overwrite(filename: str) -> None:
    """Allow overwriting temporary test files."""
    _ = filename
