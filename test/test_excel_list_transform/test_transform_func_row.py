#! /usr/local/bin/python3
"""Test the excel_list_transform functions for row merge/split."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from copy import deepcopy
from tempfile import TemporaryDirectory
import pytest
from excel_list_transform.config_enums import FileType, ColumnRef
from excel_list_transform.commontypes import NumData
from excel_list_transform.transform_func import transform_named_files
from excel_list_transform.handle_excel import write_excel_num, read_excel_num
from excel_list_transform.handle_csv import read_csv_num, write_csv_num
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName


def write_testdata_file(data: NumData, cfg: ConfigExcelListTransform,
                        filename: str) -> None:
    """Write test data to named file."""
    if cfg.in_type == FileType.EXCEL:
        write_excel_num(data=data, filename=filename + '.xlsx')
    write_csv_num(data=data, filename=filename + '.csv',
                  dialect=cfg.get_in_csv_dialect(),
                  encoding=cfg.in_csv_encoding)


def read_testdata_file(cfg: ConfigExcelListTransform,
                       filename: str) -> NumData:
    """Read test data from named file."""
    if cfg.out_type == FileType.EXCEL:
        return read_excel_num(filename=filename + '.xlsx',
                              max_column_read=20,
                              strip_col_names=False, strip_values=False)
    return read_csv_num(filename=filename + '.csv',
                        dialect=cfg.get_out_csv_dialect(),
                        encoding=cfg.out_csv_encoding,
                        max_column_read=20)


def check_equal_ignore_str_int(lhs, rhs):
    """Check that lhs and rhs are equal considering str(int) same as int."""
    if isinstance(lhs, list):
        assert isinstance(rhs, list)
        assert len(lhs) == len(rhs)
        for left, right in zip(lhs, rhs):
            check_equal_ignore_str_int(lhs=left, rhs=right)
    elif isinstance(lhs, type(rhs)):
        assert lhs == rhs
    elif isinstance(lhs, int):
        check_equal_ignore_str_int(lhs=str(lhs), rhs=rhs)
    elif isinstance(rhs, int):
        check_equal_ignore_str_int(lhs=lhs, rhs=str(rhs))
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


@pytest.mark.parametrize('exa', [ex1, ex2, ex3])
@pytest.mark.parametrize('ref', list(ColumnRef))
@pytest.mark.parametrize('intyp', list(FileType))
@pytest.mark.parametrize('outtyp', list(FileType))
@pytest.mark.parametrize('inenc', ['utf-8', 'iso8859-1'])
@pytest.mark.parametrize('outenc', ['utf-8', 'iso8859-1'])
def test_transfm_nmd_files_row_num(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments # noqa: E501
                                   exa, ref, intyp, outtyp, inenc, outenc):
    """Test transform_name_files row operations."""
    cfg = ConfigXlsListTransfNum()
    cfg.s01_split_rows = deepcopy(exa['snum'])
    cfg.s02_merge_rows = deepcopy(exa['mnum'])
    cfg.s04_remove_columns = []
    cfg.s06_place_columns_first = []
    if ref == ColumnRef.BY_NAME:
        cfg = ConfigXlsListTransfName()
        cfg.s01_split_rows = deepcopy(exa['sname'])
        cfg.s02_merge_rows = deepcopy(exa['mname'])
        cfg.s10_column_order = deepcopy(exa['order'])
    cfg.in_csv_encoding = deepcopy(inenc)
    cfg.in_csv_encoding = deepcopy(outenc)
    cfg.in_type = deepcopy(intyp)
    cfg.out_type = deepcopy(outtyp)
    cfg.s03_split_columns = []
    cfg.s05_merge_columns = []
    cfg.s07_rename_columns = []
    cfg.s08_insert_columns = []
    cfg.s09_rewrite_columns = []
    with TemporaryDirectory() as dirname:
        cfgname = dirname + '/test.cfg'  # pylint: disable=duplicate-code  # noqa: E501
        cfg.write(cfgname)
        infilename = dirname + '/in'
        outfilename = dirname + '/out'
        write_testdata_file(data=deepcopy(exa['in']), cfg=cfg,
                            filename=infilename)
        transform_named_files(infilename=infilename, outfilename=outfilename,  # pylint: disable=duplicate-code  # noqa: E501
                              cfgfilename=cfgname)
        res = read_testdata_file(cfg=cfg, filename=outfilename)
        out, err = capsys.readouterr()
        assert '' == err
        assert f'Wrote {outfilename}' in out
        check_equal_ignore_str_int(res, deepcopy(exa['out']))
