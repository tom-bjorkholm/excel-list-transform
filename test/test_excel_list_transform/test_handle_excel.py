#! /usr/local/bin/python3
"""The test cases for handle_excel."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from tempfile import TemporaryDirectory
from datetime import datetime, date, time
from decimal import Decimal
import pytest
from excel_list_transform.handle_excel import \
    read_excel_num, write_excel_num, read_excel_named, write_excel_named, \
    excel_data_strip, fix_row_openpyxl
from excel_list_transform.config_excel_list_transform import ExcelLib


@pytest.mark.parametrize('scol', [True, False])
@pytest.mark.parametrize('sval', [True, False])
@pytest.mark.parametrize('lib', [ExcelLib.OPENPYXL,
                                 ExcelLib.PYLIGHTXL,
                                 None])
def test_read_excel_num(capsys, lib, scol, sval):
    """Test reading of excel (number referenced columns)."""
    data = read_excel_num('./test/test_excel_list_transform/test_read.xlsx',
                          strip_col_names=scol, strip_values=sval,
                          max_column_read=20, excel_lib=lib)
    assert len(data) == 6
    assert data[3][1] == 'Ek'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('scol', [True, False])
@pytest.mark.parametrize('sval', [True, False])
@pytest.mark.parametrize('lib', [ExcelLib.OPENPYXL,
                                 ExcelLib.PYLIGHTXL,
                                 None])
def test_read_excel_name(capsys, lib, scol, sval):
    """Test reading of excel (name referenced columns)."""
    data = read_excel_named('./test/test_excel_list_transform/test_read.xlsx',
                            strip_col_names=scol, strip_values=sval,
                            max_column_read=20, excel_lib=lib)
    assert len(data) == 5
    assert data[2]['Efternamn'] == 'Ek'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('scol', [True, False])
@pytest.mark.parametrize('sval', [True, False])
@pytest.mark.parametrize('lib', [ExcelLib.OPENPYXL,
                                 ExcelLib.PYLIGHTXL,
                                 None])
def test_read_excel_maxcol_num(capsys, lib, scol, sval):
    """Test reading of excel (number referenced columns)."""
    data = read_excel_num('./test/test_excel_list_transform/test_read.xlsx',
                          strip_col_names=scol, strip_values=sval,
                          max_column_read=1, excel_lib=lib)
    assert len(data) == 6
    assert len(data[3]) == 1
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('lib', [ExcelLib.OPENPYXL,
                                 ExcelLib.PYLIGHTXL,
                                 None])
def test_read_excel_maxcol_name(capsys, lib):
    """Test reading of excel (name referenced columns)."""
    data = read_excel_named('./test/test_excel_list_transform/test_read.xlsx',
                            strip_col_names=False, strip_values=False,
                            max_column_read=1, excel_lib=lib)
    assert len(data) == 5
    assert len(data[3]) == 1
    assert data[3]['Förnamn'] == 'Kalle'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('lib', [ExcelLib.OPENPYXL,
                                 ExcelLib.PYLIGHTXL,
                                 ExcelLib.XLSXWRITER,
                                 None])
@pytest.mark.parametrize('zpar', [[['a', 'kalle'], ['b', None], [None, 7]],
                                  [['a', 'kalle', 2], ['b', None, 4],
                                   [None, 7, 'c']]])
def test_write_excel_1_num(capsys, zpar, lib):
    """Test writing of excel."""
    with TemporaryDirectory() as dname:
        fname = dname + '/b.xlsx'
        write_excel_num(zpar, fname, excel_lib=lib)
        data = read_excel_num(fname, max_column_read=20, strip_col_names=False,
                              strip_values=False, excel_lib=lib)
        assert len(data) == len(zpar)
        assert data == zpar
    out, err = capsys.readouterr()
    assert out == ''
    if lib != ExcelLib.XLSXWRITER:
        assert err == ''


@pytest.mark.parametrize('lib', [ExcelLib.OPENPYXL,
                                 ExcelLib.PYLIGHTXL,
                                 ExcelLib.XLSXWRITER,
                                 None])
@pytest.mark.parametrize('zpar, corder',
                         [([{'anna': 'a', 'berta': 'kalle'},
                            {'anna': 'b', 'berta': None},
                            {'anna': None, 'berta': 7}],
                           ['anna', 'berta']),
                          ([{'a': 'b', 'kalle': None, '2': 4},
                            {'a': None, 'kalle': 7, '2': 'c'}],
                           ['a', 'kalle', '2'])])
def test_write_excel_1_name(capsys, zpar, corder, lib):
    """Test writing of excel."""
    with TemporaryDirectory() as dname:
        fname = dname + '/b.xlsx'
        write_excel_named(zpar, fname, column_order=corder, excel_lib=lib)
        data = read_excel_named(fname, max_column_read=20,
                                strip_col_names=False, strip_values=False,
                                excel_lib=lib)
        assert len(data) == len(zpar)
        assert data == zpar
    out, err = capsys.readouterr()
    assert out == ''
    if lib != ExcelLib.XLSXWRITER:
        assert err == ''


#  ExcelLib.OPENPYXL as output produces incorrect files
@pytest.mark.parametrize('inlib', [ExcelLib.OPENPYXL,
                                   ExcelLib.PYLIGHTXL,
                                   None])
@pytest.mark.parametrize('outlib', [ExcelLib.PYLIGHTXL,
                                    ExcelLib.XLSXWRITER,
                                    None])
@pytest.mark.parametrize('zpar', [[['a', 'kalle'], ['b', None], [None, 7]],
                                  [['a', 'kalle', 2], ['b', None, 4],
                                   [None, 7, 'c']]])
def test_write_excel_2_num(capsys, zpar, inlib, outlib):
    """Test writing of excel."""
    with TemporaryDirectory() as dname:
        fname = dname + '/b.xlsx'
        write_excel_num(zpar, fname, excel_lib=outlib)
        data = read_excel_num(fname, max_column_read=20, strip_col_names=False,
                              strip_values=False, excel_lib=inlib)
        assert len(data) == len(zpar)
        assert data == zpar
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


#  ExcelLib.OPENPYXL as output produces incorrect files
@pytest.mark.parametrize('inlib', [ExcelLib.OPENPYXL,
                                   ExcelLib.PYLIGHTXL,
                                   None])
@pytest.mark.parametrize('outlib', [ExcelLib.PYLIGHTXL,
                                    ExcelLib.XLSXWRITER,
                                    None])
@pytest.mark.parametrize('zpar, corder',
                         [([{'anna': 'a', 'berta': 'kalle'},
                            {'anna': 'b', 'berta': None},
                            {'anna': None, 'berta': 7}],
                           ['anna', 'berta']),
                          ([{'a': 'b', 'kalle': None, '2': 4},
                            {'a': None, 'kalle': 7, '2': 'c'}],
                           ['a', 'kalle', '2'])])
def test_write_excel_2_name(capsys, zpar, corder, inlib, outlib):
    """Test writing of excel."""
    with TemporaryDirectory() as dname:
        fname = dname + '/b.xlsx'
        write_excel_named(zpar, fname, column_order=corder, excel_lib=outlib)
        data = read_excel_named(fname, max_column_read=20,
                                strip_col_names=False, strip_values=False,
                                excel_lib=inlib)
        assert len(data) == len(zpar)
        assert data == zpar
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('zpar', [[['a', 'kalle'], ['b', None], [[2, 3], 7]],
                                  [['a', 'kalle', 2], ['b', None, 4],
                                   [None, [1, 2], 'c']]])
def test_write_excel_nok(capsys, zpar):
    """Test writing of excel."""
    with TemporaryDirectory() as dname:
        fname = dname + '/b.xlsx'
        with pytest.raises(RuntimeError) as exc:
            write_excel_num(zpar, fname, excel_lib=ExcelLib.XLSXWRITER)
    out, _ = capsys.readouterr()
    assert out == ''
    assert 'Unexpected data type list' in str(exc)


@pytest.mark.parametrize('ind, stitle, svalue, outd',
                         [([['a ', ' b '], [' c', ' d '], ['e ', 'f']],
                           False, False,
                           [['a ', ' b '], [' c', ' d '], ['e ', 'f']]),
                          ([['a ', ' b '], [' c', ' d '], ['e ', 'f']],
                           True, False,
                           [['a', 'b'], [' c', ' d '], ['e ', 'f']]),
                          ([['a ', ' b '], [' c', ' d '], ['e ', 'f']],
                           False, True,
                           [['a ', ' b '], ['c', 'd'], ['e', 'f']]),
                          ([['a ', ' b '], [' c', ' d '], ['e ', 'f']],
                           True, True,
                           [['a', 'b'], ['c', 'd'], ['e', 'f']]),
                          ])
def test_excel_data_strip(capsys, ind, stitle, svalue, outd):
    """Test excel_data_strip."""
    res = excel_data_strip(data=ind, strip_col_names=stitle,
                           strip_values=svalue)
    out, err = capsys.readouterr()
    assert res == outd
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('row,maxc,expected',
                         [((1, 2, 3), 1, [1]),
                          ((2, 'a', 1.0, 'b'), 3, [2, 'a', 1.0]),
                          ((datetime(year=2025, month=6, day=7, hour=22,
                                     minute=14),
                            date(year=2014, month=12, day=25),
                            time(hour=22, minute=24, second=59)),
                           20,
                           [datetime(year=2025, month=6, day=7, hour=22,
                                     minute=14),
                            datetime(year=2014, month=12, day=25),
                            '22:24:59']),
                          ((Decimal('1.5'), 2, Decimal(4.0)), 10,
                           [1.5, 2, 4.0])])
def test_fix_row_openpyxl_ok(capsys, row, maxc, expected):
    """Test fix_row_openpyxl for OK cases."""
    ret = fix_row_openpyxl(row=row, max_column_reed=maxc)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert ret == expected


@pytest.mark.parametrize('row,maxc,expected,errtxt',
                         [((1, [3, 2], 4), 5, [1, '[3, 2]', 4],
                           'Sorry, cells with value of type ' +
                           'list are not supported\n' +
                           'Trying to convert to string.\n'),
                          (({'a': 'b'}, (4, 5), {'c': 5}), 5,
                           ["{'a': 'b'}", "(4, 5)", "{'c': 5}"],
                           'Sorry, cells with value of type ' +
                           'dict are not supported\n' +
                           'Trying to convert to string.\n'
                           'Sorry, cells with value of type ' +
                           'tuple are not supported\n' +
                           'Trying to convert to string.\n')])
def test_fix_row_openpyxl_str(capsys, row, maxc, expected, errtxt):
    """Test fix_row_openpyxl for OK cases."""
    ret = fix_row_openpyxl(row=row, max_column_reed=maxc)
    out, err = capsys.readouterr()
    assert '' == out
    assert errtxt == err
    assert ret == expected
