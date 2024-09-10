#! /usr/local/bin/python3
"""The test cases for handle_excel."""

# Copyright (c) 2024 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from tempfile import TemporaryDirectory
import pytest
from excel_list_transform.handle_excel import \
    read_excel_num, write_excel_num, read_excel_named, write_excel_named
from excel_list_transform.config_excel_list_transform import ExcelLib


@pytest.mark.parametrize('lib', [ExcelLib.OPENPYXL,
                                 ExcelLib.PYLIGHTXL,
                                 None])
def test_read_excel_num(capsys, lib):
    """Test reading of excel (number referenced columns)."""
    data = read_excel_num('./test/test_excel_list_transform/test_read.xlsx',
                          max_column_read=20, excel_lib=lib)
    assert len(data) == 6
    assert data[3][1] == 'Ek'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('lib', [ExcelLib.OPENPYXL,
                                 ExcelLib.PYLIGHTXL,
                                 None])
def test_read_excel_name(capsys, lib):
    """Test reading of excel (name referenced columns)."""
    data = read_excel_named('./test/test_excel_list_transform/test_read.xlsx',
                            max_column_read=20, excel_lib=lib)
    assert len(data) == 5
    assert data[2]['Efternamn'] == 'Ek'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('lib', [ExcelLib.OPENPYXL,
                                 ExcelLib.PYLIGHTXL,
                                 None])
def test_read_excel_maxcol_num(capsys, lib):
    """Test reading of excel (number referenced columns)."""
    data = read_excel_num('./test/test_excel_list_transform/test_read.xlsx',
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
        data = read_excel_num(fname, max_column_read=20, excel_lib=lib)
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
        data = read_excel_named(fname, max_column_read=20, excel_lib=lib)
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
        data = read_excel_num(fname, max_column_read=20, excel_lib=inlib)
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
        data = read_excel_named(fname, max_column_read=20, excel_lib=inlib)
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
