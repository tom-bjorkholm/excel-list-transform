#! /usr/local/bin/python3
"""Test the excel_list_transform functions for row merge/split."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from typing import Any
from copy import deepcopy
from tempfile import TemporaryDirectory
import pytest
from pytest import CaptureFixture
from test_excel_list_transform.tableio_helpers import \
    configure_input_csv, configure_input_excel, configure_output_csv, \
    configure_output_excel, read_excel_num, write_csv_num, write_excel_num
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.commontypes import NumData
from excel_list_transform.transform_func import transform_named_files
from excel_list_transform.handle_tableio import read_table_num
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName


def write_testdata_file(data: NumData, cfg: ConfigExcelListTransform[Any],
                        filename: str) -> None:
    """Write test data to named file."""
    if cfg.input_table.format_name == 'Excel':
        write_excel_num(data=data, filename=filename + '.xlsx')
        return
    write_csv_num(data=data, filename=filename + '.csv',
                  encoding=cfg.input_table.character_encoding or 'utf-8')


def read_test_file(cfg: ConfigExcelListTransform[Any],
                   filename: str) -> NumData:
    """Read test data from named file."""
    if cfg.output_table.format_name == 'Excel':
        return read_excel_num(filename=filename + '.xlsx', max_col_read=20,
                              strip_col_names=False, strip_values=False)
    read_cfg = ConfigXlsListTransfNum()
    read_cfg.input_table.format_name = 'CSV'
    read_cfg.input_table.character_encoding = \
        cfg.output_table.character_encoding
    read_cfg.input_table.csv = deepcopy(cfg.output_table.csv)
    read_cfg.max_column_read = 20
    return read_table_num(filename=filename + '.csv', cfg=read_cfg)


def chk_eq_ign_str_int(lhs: Any, rhs: Any) -> None:
    """Check that lhs and rhs are equal considering str(int) same as int."""
    if isinstance(lhs, list):
        assert isinstance(rhs, list)
        assert len(lhs) == len(rhs)
        for left, right in zip(lhs, rhs):
            chk_eq_ign_str_int(lhs=left, rhs=right)
    elif isinstance(lhs, type(rhs)):
        assert lhs == rhs
    elif isinstance(lhs, int):
        chk_eq_ign_str_int(lhs=str(lhs), rhs=rhs)
    elif isinstance(rhs, int):
        chk_eq_ign_str_int(lhs=lhs, rhs=str(rhs))
    else:
        assert False, 'Unexpected types: ' + \
            f'{type(lhs).__name__} and {type(rhs).__name__}'


ex1 = {'in': [['a', 'b', 'c', 'd'],
              [4, 'e f', 'g', 2],
              [5, 'h', 'i+j', 3],
              [6, 'k', 'l', 4]],
       'out': [['a', 'b', 'c', 'd'],
               [4, 'e', 'g', 2],
               [4, 'f', 'g', 2],
               [5, 'h', 'i', 3],
               [5, 'h', 'j', 3],
               [6, 'k', 'l', 4]],
       'snum': [{'column': 1, 'separators': [' '], 'not_separators': []},
                {'column': 2, 'separators': ['+'], 'not_separators': []}],
       'sname': [{'column': 'b', 'separators': [' '], 'not_separators': []},
                 {'column': 'c', 'separators': ['+'], 'not_separators': []}],
       'mnum': [],
       'mname': [],
       'order': ['a', 'b', 'c', 'd']}
ex2 = {'in': [['a', 'b', 'c', 'd'],
              [4, 'e', 'g', 2],
              [5, 'h', 'i', 3],
              [4, 'k', 'l', 4],
              [6, 'm', 'n', 3]],
       'out': [['a', 'b', 'c', 'd'],
               [4, 'e;k', 'g;l', '2;4'],
               ['5+6', 'h+m', 'i+n', 3]],
       'snum': [],
       'sname': [],
       'mnum': [{'columns': [0], 'separator': ';'},
                {'columns': [3], 'separator': '+'}],
       'mname': [{'columns': ['a'], 'separator': ';'},
                 {'columns': ['d'], 'separator': '+'}],
       'order': ['a', 'b', 'c', 'd']}
ex3 = {'in': [['a', 'b', 'c', 'd'],
              [4, 'e', 'g+h', 2],
              [5, 'j', 'h', 3],
              [6, 'm', 'g', 4]],
       'out': [['a', 'b', 'c', 'd'],
               ['4;6', 'e;m', 'g', '2;4'],
               ['4;5', 'e;j', 'h', '2;3']],
       'snum': [{'column': 2, 'separators': ['+'], 'not_separators': []}],
       'sname': [{'column': 'c', 'separators': ['+'], 'not_separators': []}],
       'mnum': [{'columns': [2], 'separator': ';'}],
       'mname': [{'columns': ['c'], 'separator': ';'}],
       'order': ['a', 'b', 'c', 'd']}


def make_row_cfg(
        exa: Any,
        ref: ColumnRef) -> ConfigXlsListTransfName | ConfigXlsListTransfNum:
    """Make row split and merge configuration."""
    if ref == ColumnRef.BY_NAME:
        cfg = ConfigXlsListTransfName()
        cfg.s01_split_rows = deepcopy(exa['sname'])
        cfg.s02_merge_rows = deepcopy(exa['mname'])
        cfg.s10_column_order = deepcopy(exa['order'])
        return cfg
    num_cfg = ConfigXlsListTransfNum()
    num_cfg.s01_split_rows = deepcopy(exa['snum'])
    num_cfg.s02_merge_rows = deepcopy(exa['mnum'])
    num_cfg.s04_remove_columns = []
    num_cfg.s06_place_columns_first = []
    return num_cfg


@pytest.mark.parametrize('exa', [ex1, ex2, ex3])
@pytest.mark.parametrize('ref', list(ColumnRef))
@pytest.mark.parametrize('intyp', ['Excel', 'CSV'])
@pytest.mark.parametrize('outtyp', ['Excel', 'CSV'])
@pytest.mark.parametrize('inenc', ['utf-8', 'iso8859-1'])
@pytest.mark.parametrize('outenc', ['utf-8', 'iso8859-1'])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_tr_nmd_files_row_num(capsys: CaptureFixture[str], exa: Any, ref: Any,
                              intyp: Any, outtyp: Any, inenc: Any,
                              outenc: Any) -> None:
    """Test transform_name_files row operations."""
    cfg = make_row_cfg(exa=exa, ref=ref)
    if intyp == 'Excel':
        configure_input_excel(cfg)
    else:
        configure_input_csv(cfg, encoding=deepcopy(inenc))
    if outtyp == 'Excel':
        configure_output_excel(cfg)
    else:
        configure_output_csv(cfg, encoding=deepcopy(outenc))
    cfg.s03_split_columns = []
    cfg.s05_merge_columns = []
    cfg.s07_rename_columns = []
    cfg.s08_insert_columns = []
    cfg.s09_rewrite_columns = []
    with TemporaryDirectory() as dirname:
        # pylint: disable-next=duplicate-code
        cfgname = dirname + '/test.cfg'
        cfg.write(cfgname)
        infilename = dirname + '/in'
        outfilename = dirname + '/out'
        write_testdata_file(data=deepcopy(exa['in']), cfg=cfg,
                            filename=infilename)
        # pylint: disable-next=duplicate-code
        transform_named_files(infilename=infilename, outfilename=outfilename,
                              cfgfilename=cfgname)
        res = read_test_file(cfg=cfg, filename=outfilename)
        out, err = capsys.readouterr()
        assert '' == err
        assert f'Wrote {outfilename}' in out
        chk_eq_ign_str_int(res, deepcopy(exa['out']))
