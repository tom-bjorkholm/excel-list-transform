#! /usr/local/bin/python3
"""The test cases for generate_txt."""

# Copyright (c) 2024 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from tempfile import TemporaryDirectory
from collections import namedtuple
from copy import deepcopy
from typing import Optional
from csv import excel as csv_excel_dialect
import pytest
from excel_list_transform.generate_cfg import generate_examplecfg
from excel_list_transform.config_enums import ColumnRef, ExcelLib
from excel_list_transform.handle_excel import read_excel_num, write_excel_num
from excel_list_transform.handle_csv import read_csv_num
from excel_list_transform.transform_func import transform_named_files
from excel_list_transform.transform_cmd import transform_cmd


class ExampleData:  # pylint: disable=too-many-instance-attributes
    """Data for testing configurations by using them."""

    def __init__(self):
        """Create test data."""
        self.phone_col_name = 'Mobiltelefonnummer'
        self.sail_num_col_name = 'Segelnummer (endast siffror)'
        self.first_name_col_name = 'Förnamn'
        self.last_name_col_name = 'Efternamn'
        self.email_col_name = 'Epostadress'
        self.nat_col_name = 'Nationalitetsbokstäver i seglet'
        self.form_columns = \
            ['ID', 'Starttid', 'Slutförandetid', 'E-post login', 'Namn login',
             'Ändrades senast klockan', self.first_name_col_name,
             self.last_name_col_name, self.email_col_name,
             self.phone_col_name, 'Klubb', 'Klass',
             self.nat_col_name, self.sail_num_col_name,
             'Ansvarsförsäkring', 'Ansluten till SSF',
             'Båten uppfyller klassreglerna', 'Visselpipa', 'Publicering',
             'Anhörig']
        self.phone_numbers_in = \
            ['070-123 45 67', '+467023456', '46070345678', '+4607045678912',
             '467056789123']
        self.phone_numbers_out = \
            ['+46701234567', '+467023456', '+4670345678', '+467045678912',
             '+467056789123']
        assert len(self.phone_numbers_in) == len(self.phone_numbers_out)
        #  self.rows = len(self.phone_numbers_in)
        self.rrs_col_order = ['Class', 'Division', 'Nationality',
                              'Sail Number', 'Boat Name', 'First Name',
                              'Last Name', 'Club Name', 'Email', 'Phone',
                              'WhatsApp']
        self.rrs_col_map = {
            'Class': 'Klass',
            'Division': None,
            'Nationality': 'Nationalitetsbokstäver i seglet',
            'Sail Number': self.sail_num_col_name,
            'Boat Name': None,
            'First Name': self.first_name_col_name,
            'Last Name': self.last_name_col_name,
            'Club Name': 'Klubb',
            'Email': self.email_col_name,
            'Phone': self.phone_col_name,
            'WhatsApp': None
        }
        self.sw_col_order = ['Class', 'Division', 'Nat', 'SailNo', 'Boat',
                             'HelmName', 'Club', 'HelmEmail', 'HelmPhone']
        self.sw_to_rrs = {
            'Class': ['Class'],
            'Division': ['Division'],
            'Nat': ['Nationality'],
            'SailNo': ['Sail Number'],
            'Boat': ['Boat Name'],
            'HelmName': ['First Name', 'Last Name'],
            'Club': ['Club Name'],
            'HelmEmail': ['Email'],
            'HelmPhone': ['Phone']
        }

    def form_data(self) -> list[list[str]]:
        """Get data as from office form for registration."""
        ret = []
        nats = ['USA', 'SWE', 'NOR']
        natindex = 0
        ret.append(self.form_columns)
        for i, phone in enumerate(self.phone_numbers_in):
            row = []
            for j, col in enumerate(self.form_columns):
                val = None
                if col == self.phone_col_name:
                    val = phone
                elif col == self.sail_num_col_name:
                    val = str(100+i)
                elif col == self.nat_col_name:
                    val = nats[natindex]
                    natindex += 1
                    if natindex >= len(nats):
                        natindex = 0
                else:
                    val = col[0:3] + '-' + str(i) + '-' + str(j)
                if col == self.email_col_name:
                    val += '@example.com'
                row.append(val)
            ret.append(row)
        return ret

    def get_sized_form_data(self, size: int) -> list[list[str]]:
        """Get for lines one by one for large file speed test."""
        nats = ['USA', 'SWE', 'NOR', 'ESP']
        ret = []
        ret.append(self.form_columns)
        for num in range(size):
            row = []
            for j, col in enumerate(self.form_columns):
                val = None
                if col == self.phone_col_name:
                    len_phone = len(self.phone_numbers_in)
                    val = self.phone_numbers_in[num % len_phone]
                elif col == self.sail_num_col_name:
                    val = str(100+num)
                elif col == self.nat_col_name:
                    val = nats[num % len(nats)]
                else:
                    val = col[0:3] + '-' + str(num) + '-' + str(j)
                    if col == self.email_col_name:
                        val += '@example.com'
                row.append(val)
            ret.append(row)
        return ret

    def form_data_col_for_rrs_data_col(self,
                                       rrs_data_col: int) -> Optional[int]:
        """Get form data column number matching rrs data column number."""
        rrs_col_name = self.rrs_col_order[rrs_data_col]
        return self.form_columns.index(self.rrs_col_map[rrs_col_name])

    def rrs_data(self) -> list[list[str]]:
        """Get data as imorted in rrs and exported from rrs."""
        ret = []
        ret.append(self.rrs_col_order)
        fdata = self.form_data()
        for row_num, frow in enumerate(fdata[1:]):
            row = []
            for icol, vcol in enumerate(self.rrs_col_order):
                if vcol == 'Phone':
                    row.append(self.phone_numbers_out[row_num])
                elif self.rrs_col_map[vcol] is None:
                    row.append(None)
                else:
                    row.append(frow[self.form_data_col_for_rrs_data_col(icol)])
            ret.append(row)
        return ret

    def rrs_data_out(self) -> list[list[str]]:
        """Get data as exported from rrs."""
        ret = deepcopy(self.rrs_data())
        ret[0][9] = 'Mobile Phone'
        for row in ret[1:]:
            phone = str(row[9])
            if phone[0] != '+':
                row[9] = '+' + phone
        return ret

    def rrs_data_col_for_sw_data_col(self, sw_data_col: int) -> list[int]:
        """Get rrs data column number matching sw data column number."""
        sw_col_name = self.sw_col_order[sw_data_col]
        rrs_cols = self.sw_to_rrs[sw_col_name]
        ret = []
        for i in rrs_cols:
            ret.append(self.rrs_col_order.index(i))
        return ret

    def sw_data(self) -> list[list[str]]:
        """Get data as imorted in sail wave."""
        ret = []
        ret.append(self.sw_col_order)
        rdata = self.rrs_data()
        for rrow in rdata[1:]:
            row = []
            for icol, _ in enumerate(self.sw_col_order):
                rcol_idx = self.rrs_data_col_for_sw_data_col(icol)
                val_list = [rrow[i] for i in rcol_idx if rrow[i] is not None]
                val = None
                if len(val_list) == 1:
                    val = val_list[0]
                elif len(val_list) > 1:
                    val = ' '.join(val_list)
                row.append(val)
            ret.append(row)
        return ret


FileNames = namedtuple('FileNames', ['indata', 'cfg', 'out'])


def openpyxl_reader(filename, max_column_read=40):
    """Read excel with openpyxl."""
    return read_excel_num(filename=filename, max_column_read=max_column_read,
                          excel_lib=ExcelLib.OPENPYXL)


def csv_reader(filename, max_column_read=40):
    """Read CSV with excel dialect."""
    return read_csv_num(filename=filename, dialect=csv_excel_dialect,
                        max_column_read=max_column_read)


@pytest.mark.parametrize('refcol', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
@pytest.mark.parametrize('cfg, indgen, resgen, reader, outname',
                         [('forms_to_rrs', ExampleData.form_data,
                           ExampleData.rrs_data, openpyxl_reader, 'out.xlsx'),
                          ('forms_to_sw', ExampleData.form_data,
                           ExampleData.sw_data, csv_reader, 'out.csv'),
                          ('rrs_to_sw', ExampleData.rrs_data_out,
                           ExampleData.sw_data, csv_reader, 'out.csv')])
def test_cfg_gen_used(capsys,  # pylint: disable=too-many-arguments disable=too-many-locals # noqa: E501
                      refcol, cfg, indgen, resgen, reader, outname):
    """Test to generate configuration and use it."""
    test_data = ExampleData()
    res = None
    with TemporaryDirectory() as dname:
        files = FileNames(indata=dname + '/a.xlsx',
                          cfg=dname + '/a.cfg',
                          out=dname + '/' + outname)
        generate_examplecfg(cfgtype=cfg, filename=files.cfg,
                            colref=refcol)
        write_excel_num(data=indgen(test_data), filename=files.indata)
        transform_named_files(infilename=files.indata, outfilename=files.out,
                              cfgfilename=files.cfg)
        res = reader(filename=files.out, max_column_read=40)
    out, err = capsys.readouterr()
    tomatch = resgen(test_data)
    assert len(res) == len(tomatch)
    for r, t in zip(res, tomatch):  # error msg is easier to read with loop
        assert r == t
    assert res == tomatch
    assert 'Wrote ' + files.out in out
    assert '' == err


@pytest.mark.parametrize('refcol', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
@pytest.mark.parametrize('cfg, indgen, resgen, reader, outname',
                         [('forms_to_rrs', ExampleData.form_data,
                           ExampleData.rrs_data, openpyxl_reader, 'out.xlsx'),
                          ('forms_to_sw', ExampleData.form_data,
                           ExampleData.sw_data, csv_reader, 'out.csv'),
                          ('rrs_to_sw', ExampleData.rrs_data_out,
                           ExampleData.sw_data, csv_reader, 'out.csv')])
def test_cfg_and_cmd(capsys,  # pylint: disable=too-many-arguments disable=too-many-locals # noqa: E501
                     refcol, cfg, indgen, resgen, reader, outname):
    """Test to generate configuration and use it."""
    test_data = ExampleData()
    res = None
    with TemporaryDirectory() as dname:
        files = FileNames(indata=dname + '/a.xlsx',
                          cfg=dname + '/a.cfg',
                          out=dname + '/' + outname)
        transform_cmd(['example', '-r', refcol.name.lower(), '-k', cfg,
                      '-o', files.cfg])
        write_excel_num(data=indgen(test_data), filename=files.indata)
        transform_cmd(['transform', '-i', files.indata, '-o', files.out,
                      '-c', files.cfg])
        res = reader(filename=files.out, max_column_read=40)
    out, err = capsys.readouterr()
    tomatch = resgen(test_data)
    assert len(res) == len(tomatch)
    for r, t in zip(res, tomatch):  # error msg is easier to read with loop
        assert r == t
    assert res == tomatch
    assert 'Wrote ' + files.out in out
    assert '' == err


@pytest.mark.parametrize('refcol', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
def test_gen_example(capsys, refcol):
    """Test generate_syntax_example."""
    with TemporaryDirectory() as dname:
        cfg = dname + '/a.cfg'
        txt = dname + '/a.txt'
        generate_examplecfg(cfgtype='example', filename=cfg, colref=refcol)
        with open(file=cfg, mode='r', encoding='utf-8') as file:
            cfgtxt = file.read()
            assert refcol.name in cfgtxt
        with open(file=txt, mode='r', encoding='utf-8') as file:
            txttxt = file.read()
            assert 's5_rename_columns' in txttxt
    out, err = capsys.readouterr()
    assert f'Wrote files {cfg} and {txt}' in out
    assert '' == err
