#! /usr/local/bin/python3
"""Functions for doing transform of lists excel files."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from copy import deepcopy
from excel_list_transform.handle_csv import read_csv_num, write_csv_num
from excel_list_transform.handle_excel import read_excel_num, write_excel_num
from excel_list_transform.config_enums import FileType
from excel_list_transform.commontypes import NumData, get_checked_type
from excel_list_transform.config import Config
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.check_indata_common import check_indata_common
from excel_list_transform.transform_func_common import \
    cols_must_exist_dict, cols_must_exist_lst, split_columns, merge_columns, \
    rewrite_columns
from excel_list_transform.row_split_merge import split_rows_cfg, \
    merge_rows_cfg


def remove_columns_num(indata: NumData,
                       cfg: ConfigXlsListTransfNum) -> NumData:
    """Remove columns in the list in indata with column number refs."""
    if len(cfg.s04_remove_columns) == 0:
        return indata
    ret = deepcopy(indata)
    for row in ret:
        rowlen = len(row)
        cols_must_exist_lst(cols=cfg.s04_remove_columns, row=row,
                            param='s04_remove_columns', tinfo=2)
        for i in reversed(cfg.s04_remove_columns):
            if 0 <= i < rowlen:
                row.pop(i)
    return ret


def place_columns_first(indata: NumData,
                        cfg: ConfigXlsListTransfNum) -> NumData:
    """Place columns first in the list in indata."""
    if len(cfg.s06_place_columns_first) == 0:
        return indata
    ret = deepcopy(indata)
    for row_i, row in enumerate(ret):
        rowlen = len(row)
        cols_must_exist_lst(cols=cfg.s06_place_columns_first, row=row,
                            param='s06_place_columns_first', tinfo=2)
        drow = dict(enumerate(row))
        order = deepcopy(cfg.s06_place_columns_first)
        for i in range(rowlen):
            if i not in order:
                order.append(i)
        nrow = []
        for i in order:
            nrow.append(drow[i])
        ret[row_i] = nrow
    return ret


def rename_columns_num(indata: NumData,
                       cfg: ConfigXlsListTransfNum) -> NumData:
    """Rename columns in the list in indata with column number refs."""
    assert isinstance(cfg.s07_rename_columns, list)
    if len(cfg.s07_rename_columns) == 0:
        return indata
    ret = deepcopy(indata)
    cols_must_exist_dict(rule=cfg.s07_rename_columns, row=ret[0],
                         param='s07_rename_columns', tinfo=2)
    for i in cfg.s07_rename_columns:
        colref: int = get_checked_type(i['column'], int)
        assert isinstance(colref, int)
        ret[0][colref] = get_checked_type(i['name'], str)
    return ret


def insert_columns_num(indata: NumData,
                       cfg: ConfigXlsListTransfNum) -> NumData:
    """Insert columns in the list in indata with column number refs."""
    assert isinstance(cfg.s08_insert_columns, list)
    if len(cfg.s08_insert_columns) == 0:
        return indata
    ret = deepcopy(indata)
    for row_i, row in enumerate(ret):
        for i in cfg.s08_insert_columns:
            colref: int = get_checked_type(value=i['column'], istype=int)
            if not 0 <= colref <= len(row):
                msg = f's08_insert_columns: column index {colref} out of '
                msg += f'range [0, {len(row)}].'
                print(msg, file=sys.stderr)
                sys.exit(1)
            name: str = get_checked_type(i['name'], str)
            to_insert = name if row_i == 0 else i['value']
            if colref == len(row):
                row.append(to_insert)
            else:
                row.insert(colref, to_insert)
    return ret


def fix_indata_empty_rows_num(indata: NumData) -> None:  # pylint: disable=duplicate-code # noqa: E501
    """Check rows and remove empty rows with column number refs."""
    for i, row in reversed(list(enumerate(indata))):
        assert isinstance(row, list), 'Internal error. Expected list of ' + \
            f'columns but got {type(row).__name__}'  # pylint: disable=duplicate-code # noqa: E501
        if len(row) == 0:
            del indata[i]
        elif len(row) == 1:
            if row[0] is None:
                del indata[i][0]
                del indata[i]
            elif isinstance(row[0], str) and len(row[0]) == 0:
                del indata[i][0]
                del indata[i]


def check_indata_num(indata: NumData) -> None:
    """Check that the indata is well formed with column number refs."""
    check_indata_common(indata=indata,
                        fix_indata_empty_rows=fix_indata_empty_rows_num)


def transform_data_num(indata: NumData,
                       cfg: ConfigXlsListTransfNum) -> NumData:
    """Transform list in the data with column number refs."""
    check_indata_num(indata=indata)
    ret = split_rows_cfg(indata=indata, cfg=cfg, tinfo=2)
    ret = merge_rows_cfg(indata=ret, cfg=cfg, tinfo=2)
    ret = split_columns(indata=ret, cfg=cfg, tinfo=2)
    ret = remove_columns_num(indata=ret, cfg=cfg)
    ret = merge_columns(indata=ret, cfg=cfg, tinfo=2)
    ret = place_columns_first(indata=ret, cfg=cfg)
    ret = rename_columns_num(indata=ret, cfg=cfg)
    ret = insert_columns_num(indata=ret, cfg=cfg)
    ret = rewrite_columns(indata=ret, cfg=cfg, tinfo=2)
    return ret


def transform_named_files_num(infilename: str, outfilename: str,
                              cfg: Config) -> None:
    """Transform list in the named excel file to named file."""
    cfgn: ConfigXlsListTransfNum = \
        get_checked_type(value=cfg, istype=ConfigXlsListTransfNum)
    indata = None
    if cfgn.in_type == FileType.CSV:
        indata = read_csv_num(infilename, cfgn.get_in_csv_dialect(),
                              encoding=cfgn.in_csv_encoding,
                              max_column_read=cfgn.max_column_read)
    else:
        indata = read_excel_num(infilename,
                                max_column_read=cfgn.max_column_read,
                                strip_col_names=cfgn.in_excel_col_name_strip,
                                strip_values=cfgn.in_excel_values_strip,
                                excel_lib=cfgn.in_excel_library)
    outdata = transform_data_num(indata=indata, cfg=cfgn)
    if cfgn.out_type == FileType.CSV:
        write_csv_num(data=outdata, filename=outfilename,
                      dialect=cfgn.get_out_csv_dialect(),
                      encoding=cfgn.out_csv_encoding)
    else:
        write_excel_num(data=outdata, filename=outfilename,
                        excel_lib=cfgn.out_excel_library)
    print(f'Wrote {outfilename}')
