#! /usr/local/bin/python3
"""The test cases for handle_csv."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


import csv
from tempfile import TemporaryDirectory
import pytest
from excel_list_transform.handle_csv import read_csv_num, write_csv_num, \
    read_csv_named, write_csv_named, read_csv_named_use_num


def test_read_csv_num(capsys):
    """Test reading of excel."""
    dial = csv.excel
    dial.lineterminator = '\n'
    dial.delimiter = ','
    data = read_csv_num('./test/test_excel_list_transform/test_read.csv',
                        dialect=csv.excel, encoding='utf_8_sig',
                        max_column_read=20)
    assert len(data) == 6
    assert data[3][1] == 'Ek'
    assert data[0][0] == 'Förnamn'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


def test_read_csv_named(capsys):
    """Test reading of excel."""
    dial = csv.excel
    dial.lineterminator = '\n'
    dial.delimiter = ','
    data = read_csv_named('./test/test_excel_list_transform/test_read.csv',
                          dialect=csv.excel, encoding='utf_8_sig',
                          max_column_read=20)
    assert len(data) == 5
    assert data[2]['Efternamn'] == 'Ek'
    assert data[0]['Förnamn'] == 'Anna'
    assert data[4]['Tel'] == '+4678901'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


def test_read_csv_named_use_num(capsys):
    """Test reading of excel."""
    dial = csv.excel
    dial.lineterminator = '\n'
    dial.delimiter = ','
    file = './test/test_excel_list_transform/test_read.csv'
    data = read_csv_named_use_num(file, dialect=csv.excel,
                                  encoding='utf_8_sig',
                                  max_column_read=20)
    assert len(data) == 5
    assert data[2]['Efternamn'] == 'Ek'
    assert data[0]['Förnamn'] == 'Anna'
    assert data[4]['Tel'] == '+4678901'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


def test_read_csv_maxcol_num(capsys):
    """Test reading of excel."""
    dial = csv.excel
    dial.lineterminator = '\n'
    dial.delimiter = ','
    data = read_csv_num('./test/test_excel_list_transform/test_read.csv',
                        dialect=csv.excel, encoding='utf_8_sig',
                        max_column_read=1)
    assert len(data) == 6
    assert len(data[3]) == 1
    assert data[0][0] == 'Förnamn'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


def test_read_csv_maxcol_named(capsys):
    """Test reading of excel."""
    dial = csv.excel
    dial.lineterminator = '\n'
    dial.delimiter = ','
    data = read_csv_named('./test/test_excel_list_transform/test_read.csv',
                          dialect=csv.excel, encoding='utf_8_sig',
                          max_column_read=1)
    assert len(data) == 5
    assert len(data[3]) == 1
    assert data[0]['Förnamn'] == 'Anna'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('enc', ['utf-8', 'iso8859-1'])
@pytest.mark.parametrize('dial', [csv.excel, csv.excel_tab, csv.unix_dialect])
@pytest.mark.parametrize('zpar', [[['a', 'kalle'], ['b', None], [None, 7],
                                   ['Hör', 'ÅÄÖåäö']],
                                  [['a', 'kalle', 2], ['b', None, 4],
                                   [None, 7, 'c'],
                                   ['Förnamn', 'gåta', 'ÅÄÖä']]])
def test_write_csv_num(capsys, zpar, dial, enc):
    """Test writing of csv."""
    with TemporaryDirectory() as dname:
        fname = dname + '/b.csv'
        dial.lineterminator = '\n'
        write_csv_num(zpar, fname, dialect=dial, encoding=enc)
        data = read_csv_num(fname, dialect=dial, encoding=enc,
                            max_column_read=20)
        assert len(data) == len(zpar)
        for i, row in enumerate(data):
            assert len(row) == len(zpar[i])
            for j, elem in enumerate(row):
                if elem is not None and elem.isdigit():
                    assert int(elem) == zpar[i][j]
                else:
                    assert elem == zpar[i][j]
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('enc', ['utf-8', 'iso8859-1'])
@pytest.mark.parametrize('dial', [csv.excel, csv.excel_tab, csv.unix_dialect])
@pytest.mark.parametrize('zpar, corder',
                         [([{'a': 'b', 'kalle': None},
                            {'a': None, 'kalle': 7},
                            {'a': 'Hör', 'kalle': 'ÅÄÖåäö'}],
                           ['a', 'kalle']),
                          ([{'a': 'b', 'kalle': None, '2': 4},
                            {'a': None, 'kalle': 7, '2': 'c'},
                            {'a': 'Förnamn', 'kalle': 'Gåta', '2': 'ÅÄÖä'}],
                           ['a', 'kalle', '2'])])
def test_write_csv_named(capsys, zpar, dial, corder, enc):
    """Test writing of csv."""
    with TemporaryDirectory() as dname:
        fname = dname + '/b.csv'
        dial.lineterminator = '\n'
        write_csv_named(zpar, fname, dialect=dial, encoding=enc,
                        column_order=corder)
        data = read_csv_named(fname, dialect=dial, encoding=enc,
                              max_column_read=20)
        assert len(data) == len(zpar)
        for i, row in enumerate(data):
            assert len(row) == len(zpar[i])
            for key, elem in row.items():
                if elem is not None and elem.isdigit():
                    assert int(elem) == zpar[i][key]
                else:
                    assert elem == zpar[i][key]
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


def test_write_csv_num_encmiss(capsys):
    """Test write csv num read with other encoding."""
    data = [['a', 'b'], ['Hör', 'vadå']]
    with TemporaryDirectory() as dname:
        fname = dname + '/b.csv'
        write_csv_num(data=data, filename=fname,
                      dialect=csv.unix_dialect, encoding='utf8')
        rdata = read_csv_num(filename=fname, dialect=csv.unix_dialect,
                             encoding='iso8859-1', max_column_read=20)
        out, err = capsys.readouterr()
        assert rdata[0][0] == data[0][0]
        assert rdata[0][1] == data[0][1]
        assert 'HÃ¶r' == rdata[1][0]
        assert 'vadÃ¥' == rdata[1][1]
        assert '' == err
        assert '' == out


def test_write_csv_name_encmiss(capsys):
    """Test write csv num read with other encoding."""
    data = [{'a': 'a', 'b': 'b'},
            {'a': 'Hör', 'b': 'vadå'}]
    with TemporaryDirectory() as dname:
        fname = dname + '/b.csv'
        write_csv_named(data=data, filename=fname,
                        dialect=csv.unix_dialect, encoding='utf8',
                        column_order=['a', 'b'])
        rdata = read_csv_named(filename=fname, dialect=csv.unix_dialect,
                               encoding='iso8859-1', max_column_read=20)
        out, err = capsys.readouterr()
        assert rdata[0]['a'] == data[0]['a']
        assert rdata[0]['b'] == data[0]['b']
        assert 'HÃ¶r' == rdata[1]['a']
        assert 'vadÃ¥' == rdata[1]['b']
        assert '' == err
        assert '' == out


def test_write_csv_name_encmis2(capsys):
    """Test write csv num read with other encoding."""
    data = [{'a': 'a', 'b': 'b'},
            {'a': 'Hör', 'b': 'vadå'}]
    with TemporaryDirectory() as dname:
        fname = dname + '/b.csv'
        write_csv_named(data=data, filename=fname,
                        dialect=csv.unix_dialect, encoding='utf8',
                        column_order=['a', 'b'])
        rdata = read_csv_named(filename=fname, dialect=csv.unix_dialect,
                               encoding='iso8859-1', max_column_read=1)
        out, err = capsys.readouterr()
        assert rdata[0]['a'] == data[0]['a']
        assert 'HÃ¶r' == rdata[1]['a']
        assert '' == err
        assert '' == out
