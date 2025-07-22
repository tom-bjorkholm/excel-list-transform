#! /usr/local/bin/python3
"""Functions for doing transform of lists excel files."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from copy import deepcopy
from excel_list_transform.handle_csv import read_csv_named, write_csv_named
from excel_list_transform.handle_excel import \
    read_excel_named, write_excel_named
from excel_list_transform.config_enums import \
    FileType
from excel_list_transform.config import Config
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.check_indata_common import check_indata_common
from excel_list_transform.commontypes import NameData, get_checked_type
from excel_list_transform.transform_func_common import \
    cols_must_exist_dict, split_columns, merge_columns, rewrite_columns
from excel_list_transform.row_split_merge import split_rows_cfg, \
    merge_rows_cfg


def rename_columns_name(indata: NameData,
                        cfg: ConfigXlsListTransfName) -> NameData:
    """Rename columns in the list in indata with column name refs."""
    if len(cfg.s07_rename_columns) == 0:
        return indata
    ret = deepcopy(indata)
    for row in ret:
        cols_must_exist_dict(rule=cfg.s07_rename_columns, row=row,
                             param='s07_rename_columns', tinfo='a')
        for i in cfg.s07_rename_columns:
            from_name: str = get_checked_type(value=i['column'], istype=str)
            to_name: str = get_checked_type(value=i['name'], istype=str)
            val = row.pop(from_name)
            row[to_name] = val
    return ret


def insert_columns_name(indata: NameData,
                        cfg: ConfigXlsListTransfName) -> NameData:
    """Insert columns in the list in indata with column name refs."""
    if len(cfg.s08_insert_columns) == 0:
        return indata
    ret = deepcopy(indata)
    for row in ret:
        for i in cfg.s08_insert_columns:
            col: str = get_checked_type(value=i['column'], istype=str)
            val = i['value']
            if col in row:
                msg = f's08_insert_columns: column "{col}" '
                msg += 'already exists.'
                print(msg, file=sys.stderr)
                sys.exit(1)
            row[col] = val
    return ret


def fix_indata_empty_rows_name(indata: NameData) -> None:  # pylint: disable=duplicate-code # noqa: E501
    """Check rows and remove empty rows with column name refs."""
    for i, row in reversed(list(enumerate(indata))):
        if not isinstance(row, dict):
            msg = 'Internal error. Expected dict of columns but got '
            msg += type(row).__name__  # pylint: disable=duplicate-code # noqa: E501
            print(msg, file=sys.stderr)
            raise TypeError(msg)
        if len(row) == 0:
            del indata[i]
            continue
        for val in row.values():
            if val is not None and not isinstance(val, str):
                return  # last row not empty
            if val is not None and isinstance(val, str):
                if len(val) > 0:
                    return  # last rown not empty
        # last row is empty, remove it
        keys = deepcopy(list(row.keys()))
        for key in keys:
            del row[key]
        del indata[i]
    return


def check_indata_name(indata: NameData) -> None:
    """Check that the indata is well formed with column name refs."""
    check_indata_common(indata=indata,
                        fix_indata_empty_rows=fix_indata_empty_rows_name)
    cols = None
    for row in indata:
        rowcols = sorted(row.keys())
        if cols is None:
            cols = rowcols
        else:
            if cols != rowcols:
                msg = 'Columns names different between lines. '
                msg += f'Found {cols} and {rowcols}. Aborting.'
                print(msg, file=sys.stderr)
                raise RuntimeError(msg)


def transform_data_name(indata: NameData,
                        cfg: ConfigXlsListTransfName) -> NameData:
    """Transform list in the data with column number refs."""
    check_indata_name(indata=indata)
    ret = split_rows_cfg(indata=indata, cfg=cfg, tinfo='a')
    ret = merge_rows_cfg(indata=ret, cfg=cfg, tinfo='a')
    ret = split_columns(indata=ret, cfg=cfg, tinfo='a')
    ret = merge_columns(indata=ret, cfg=cfg, tinfo='a')
    ret = rename_columns_name(indata=ret, cfg=cfg)
    ret = insert_columns_name(indata=ret, cfg=cfg)
    ret = rewrite_columns(indata=ret, cfg=cfg, tinfo='a')
    return ret


def transform_named_files_name(infilename: str, outfilename: str,
                               cfg: Config) -> None:
    """Transform list in the named excel file to named file."""
    cfgn: ConfigXlsListTransfName = \
        get_checked_type(value=cfg, istype=ConfigXlsListTransfName)
    indata = None
    if cfgn.in_type == FileType.CSV:
        indata = read_csv_named(infilename, cfgn.get_in_csv_dialect(),
                                encoding=cfgn.in_csv_encoding,
                                max_column_read=cfgn.max_column_read)
    else:
        indata = read_excel_named(infilename,
                                  max_column_read=cfgn.max_column_read,
                                  strip_col_names=cfgn.in_excel_col_name_strip,
                                  strip_values=cfgn.in_excel_values_strip,
                                  excel_lib=cfgn.in_excel_library)
    outdata = transform_data_name(indata=indata, cfg=cfgn)
    if cfgn.out_type == FileType.CSV:
        write_csv_named(data=outdata, filename=outfilename,
                        dialect=cfgn.get_out_csv_dialect(),
                        encoding=cfgn.out_csv_encoding,
                        column_order=cfgn.s10_column_order)
    else:
        write_excel_named(data=outdata, filename=outfilename,
                          column_order=cfgn.s10_column_order,
                          excel_lib=cfgn.out_excel_library)
    print(f'Wrote {outfilename}')
