#! /usr/local/bin/python3
"""Functions for transforming list data with named columns."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


import sys
from copy import deepcopy
from config_as_json import Config
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.check_indata_common import check_indata_common
from excel_list_transform.commontypes import NameData, NameRow, \
    get_checked_type
from excel_list_transform.transform_func_common import \
    cols_must_exist_lst, col_must_exist_name, split_columns, merge_columns, \
    rewrite_columns
from excel_list_transform.row_split_merge import split_rows_cfg, \
    merge_rows_cfg
from excel_list_transform.handle_tableio import read_table_named, \
    write_table_named


def rename_columns_name(indata: NameData,
                        cfg: ConfigXlsListTransfName) -> NameData:
    """Rename columns in the list in indata with column name refs.

    Renames are applied in order. A rename must not overwrite an
    existing column (so a swap of two names is an error).
    """
    if len(cfg.s07_rename_columns) == 0:
        return indata
    ret = deepcopy(indata)
    for row in ret:
        for i in cfg.s07_rename_columns:
            from_name: str = get_checked_type(value=i['column'], istype=str)
            to_name: str = get_checked_type(value=i['name'], istype=str)
            rename_column(row=row, from_name=from_name, to_name=to_name)
    return ret


def rename_column(row: NameRow, from_name: str, to_name: str) -> None:
    """Rename one column in row, stop with error report on collision."""
    col_must_exist_name(col=from_name, row=row, param='s07_rename_columns')
    val = row.pop(from_name)
    if to_name in row:
        msg = f's07_rename_columns: cannot rename "{from_name}" to '
        msg += f'"{to_name}", column "{to_name}" already exists.'
        print(msg, file=sys.stderr)
        sys.exit(1)
    row[to_name] = val


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


# pylint: disable-next=duplicate-code
def fix_named_empty_rows(indata: NameData) -> None:
    """Check rows and remove empty rows with column name refs."""
    for i, row in reversed(list(enumerate(indata))):
        if not isinstance(row, dict):
            msg = 'Internal error. Expected dict of columns but got '
            msg += type(row).__name__
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
                        fix_indata_empty_rows=fix_named_empty_rows)
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


def column_order_must_exist(indata: NameData,
                            cfg: ConfigXlsListTransfName) -> None:
    """Stop with error report if s10_column_order names a missing column."""
    for row in indata:
        cols_must_exist_lst(cols=cfg.s10_column_order, row=row,
                            param='s10_column_order', tinfo='a')


def transform_data_name(indata: NameData,
                        cfg: ConfigXlsListTransfName) -> NameData:
    """Transform list in the data with column name refs."""
    check_indata_name(indata=indata)
    ret = split_rows_cfg(indata=indata, cfg=cfg, tinfo='a')
    ret = merge_rows_cfg(indata=ret, cfg=cfg, tinfo='a')
    ret = split_columns(indata=ret, cfg=cfg, tinfo='a')
    ret = merge_columns(indata=ret, cfg=cfg, tinfo='a')
    ret = rename_columns_name(indata=ret, cfg=cfg)
    ret = insert_columns_name(indata=ret, cfg=cfg)
    ret = rewrite_columns(indata=ret, cfg=cfg, tinfo='a')
    column_order_must_exist(indata=ret, cfg=cfg)
    return ret


def transform_named_files_name(infilename: str, outfilename: str,
                               cfg: Config) -> None:
    """Transform list data from the named input file to output file."""
    cfgn: ConfigXlsListTransfName = \
        get_checked_type(value=cfg, istype=ConfigXlsListTransfName)
    indata = read_table_named(infilename, cfgn)
    outdata = transform_data_name(indata=indata, cfg=cfgn)
    written_file = write_table_named(data=outdata, filename=outfilename,
                                     cfg=cfgn,
                                     column_order=cfgn.s10_column_order)
    print(f'Wrote {written_file}')
