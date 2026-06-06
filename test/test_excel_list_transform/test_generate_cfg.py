#! /usr/local/bin/python3
"""The test cases for generate_txt."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from collections.abc import Callable, Iterable
from collections import namedtuple
from tempfile import TemporaryDirectory
from copy import deepcopy
import pytest
from pytest import CaptureFixture
from tableio import CsvDialect
from tableio_cfg_json import TioJsonCsvConfig
from test_excel_list_transform.tableio_helpers import read_excel_num, \
    write_excel_num
from excel_list_transform.generate_cfg import generate_examplecfg, \
    get_example_names
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.handle_tableio import read_table_num
from excel_list_transform.transform_func import transform_named_files
from excel_list_transform.transform_cmd import transform_cmd
from excel_list_transform.commontypes import NumData, NumRow, Value


def string_row(values: Iterable[str]) -> NumRow:
    """Return a numbered test-data row from strings."""
    row: NumRow = []
    for value in values:
        row.append(value)
    return row


class ExampleData:  # pylint: disable=too-many-instance-attributes
    """Data for testing configurations by using them."""

    def __init__(self) -> None:
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

    def form_data(self) -> NumData:
        """Get data as from office form for registration."""
        ret: NumData = []
        nats = ['USA', 'SWE', 'NOR']
        natindex = 0
        ret.append(string_row(self.form_columns))
        for i, phone in enumerate(self.phone_numbers_in):
            row: NumRow = []
            for j, col in enumerate(self.form_columns):
                val: Value = None
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

    def get_sized_form_data(self, size: int) -> NumData:
        """Get for lines one by one for large file speed test."""
        nats = ['USA', 'SWE', 'NOR', 'ESP']
        ret: NumData = []
        ret.append(string_row(self.form_columns))
        for num in range(size):
            row: NumRow = []
            for j, col in enumerate(self.form_columns):
                val: Value = None
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

    def form_col_for_rrs_col(self, rrs_data_col: int) -> int:
        """Get form data column number matching rrs data column number."""
        rrs_col_name = self.rrs_col_order[rrs_data_col]
        form_col_name = self.rrs_col_map[rrs_col_name]
        assert form_col_name is not None
        return self.form_columns.index(form_col_name)

    def rrs_data(self) -> NumData:
        """Get data as imorted in rrs and exported from rrs."""
        ret: NumData = []
        ret.append(string_row(self.rrs_col_order))
        fdata = self.form_data()
        for row_num, frow in enumerate(fdata[1:]):
            row: NumRow = []
            for icol, vcol in enumerate(self.rrs_col_order):
                if vcol == 'Phone':
                    row.append(self.phone_numbers_out[row_num])
                elif self.rrs_col_map[vcol] is None:
                    row.append(None)
                else:
                    row.append(frow[self.form_col_for_rrs_col(icol)])
            ret.append(row)
        return ret

    def rrs_data_out(self) -> NumData:
        """Get data as exported from rrs."""
        ret = deepcopy(self.rrs_data())
        ret[0][9] = 'Mobile Phone'
        for row in ret[1:]:
            phone = str(row[9])
            if phone[0] != '+':
                row[9] = '+' + phone
        return ret

    def rrs_col_for_sw_data_col(self, sw_data_col: int) -> list[int]:
        """Get rrs data column number matching sw data column number."""
        sw_col_name = self.sw_col_order[sw_data_col]
        rrs_cols = self.sw_to_rrs[sw_col_name]
        ret = []
        for i in rrs_cols:
            ret.append(self.rrs_col_order.index(i))
        return ret

    def sw_data(self) -> NumData:
        """Get data as imorted in sail wave."""
        ret: NumData = []
        ret.append(string_row(self.sw_col_order))
        rdata = self.rrs_data()
        for rrow in rdata[1:]:
            row: NumRow = []
            for icol, _ in enumerate(self.sw_col_order):
                rcol_idx = self.rrs_col_for_sw_data_col(icol)
                val_list = [str(rrow[i]) for i in rcol_idx
                            if rrow[i] is not None]
                val: Value = None
                if len(val_list) == 1:
                    val = val_list[0]
                elif len(val_list) > 1:
                    val = ' '.join(val_list)
                row.append(val)
            ret.append(row)
        return ret

    def sw_data_extacted(self) -> NumData:
        """Get data as extracted from sail wave export."""
        ret = deepcopy(self.sw_data())
        ret[0] = string_row(
            ['Class', 'Division', 'Nationality', 'Sail Number', 'Boat Name',
             'Name', 'Club Name', 'Email', 'Phone'])
        return ret


FileNames = namedtuple('FileNames', ['indata', 'cfg', 'out'])


def _assert_txt_layout(content: str) -> None:
    """Assert generated syntax text has consistent indentation."""
    assert '\n\n\n' not in content
    for line in content.splitlines():
        assert not line.startswith('    ')


def openpyxl_reader(filename: str, max_column_read: int = 40) -> NumData:
    """Read Excel output through the app TableIO reader."""
    return read_excel_num(filename=filename, max_col_read=max_column_read,
                          strip_col_names=False, strip_values=False)


def csv_delimiter(filename: str) -> str:
    """Return likely CSV delimiter from first line in a test file."""
    with open(file=filename, mode='r', encoding='utf_8_sig') as csv_file:
        first_line = csv_file.readline()
    if first_line.count(';') > first_line.count(','):
        return ';'
    return ','


def csv_reader(filename: str, max_column_read: int = 40) -> NumData:
    """Read CSV with TableIO."""
    cfg = ConfigXlsListTransfNum()
    cfg.input_table.format_name = 'CSV'
    cfg.input_table.character_encoding = 'utf_8_sig'
    cfg.input_table.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                           delimiter=csv_delimiter(filename),
                                           quotechar='"')
    cfg.max_column_read = max_column_read
    return read_table_num(filename=filename, cfg=cfg)


@pytest.mark.parametrize('refcol', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
@pytest.mark.parametrize('cfg, indgen, resgen, reader, outname',
                         [('forms_to_rrs', ExampleData.form_data,
                           ExampleData.rrs_data, openpyxl_reader, 'out.xlsx'),
                          ('forms_to_sw', ExampleData.form_data,
                           ExampleData.sw_data, csv_reader, 'out.csv'),
                          ('rrs_to_sw', ExampleData.rrs_data_out,
                           ExampleData.sw_data, csv_reader, 'out.csv')])
# pylint: disable-next=R0913,R0914,R0917
def test_cfg_gen_used(capsys: CaptureFixture[str], refcol: ColumnRef, cfg: str,
                      indgen: Callable[[ExampleData], NumData],
                      resgen: Callable[[ExampleData], NumData],
                      reader: Callable[[str, int], NumData],
                      outname: str) -> None:
    """Test to generate configuration and use it."""
    test_data = ExampleData()
    res = None
    with TemporaryDirectory() as dname:
        files = FileNames(dname + '/a.xlsx', dname + '/a.cfg',
                          dname + '/' + outname)
        generate_examplecfg(cfgtype=cfg, filename=files.cfg, colref=refcol)
        write_excel_num(data=indgen(test_data), filename=files.indata)
        transform_named_files(infilename=files.indata, outfilename=files.out,
                              cfgfilename=files.cfg)
        res = reader(files.out, 40)
    out, err = capsys.readouterr()
    tomatch = resgen(test_data)
    assert len(res) == len(tomatch)
    for r, t in zip(res, tomatch):  # error msg is easier to read with loop
        assert r == t
    assert res == tomatch
    assert 'Wrote ' + files.out in out
    assert '' == err


@pytest.mark.parametrize('example', ['cfg-example', 'example'])
@pytest.mark.parametrize('refcol', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
@pytest.mark.parametrize('cfg, indgen, resgen, reader, outname',
                         [('forms_to_rrs', ExampleData.form_data,
                           ExampleData.rrs_data, openpyxl_reader, 'out.xlsx'),
                          ('forms_to_sw', ExampleData.form_data,
                           ExampleData.sw_data, csv_reader, 'out.csv'),
                          ('rrs_to_sw', ExampleData.rrs_data_out,
                           ExampleData.sw_data, csv_reader, 'out.csv'),
                          ('sw_to_rrs', ExampleData.sw_data_extacted,
                           ExampleData.rrs_data, openpyxl_reader, 'out.xlsx')])
# pylint: disable-next=R0913,R0914,R0917
def test_cfg_and_cmd(capsys: CaptureFixture[str], refcol: ColumnRef, cfg: str,
                     indgen: Callable[[ExampleData], NumData],
                     resgen: Callable[[ExampleData], NumData],
                     reader: Callable[[str, int], NumData], outname: str,
                     example: str) -> None:
    """Test to generate configuration and use it."""
    test_data = ExampleData()
    res = None
    with TemporaryDirectory() as dname:
        files = FileNames(dname + '/a.xlsx', dname + '/a.cfg',
                          dname + '/' + outname)
        transform_cmd([example, '-r', refcol.name.lower(), '-k', cfg,
                      '-o', files.cfg])
        write_excel_num(data=indgen(test_data), filename=files.indata)
        transform_cmd(['transform', '-i', files.indata, '-o', files.out,
                      '-c', files.cfg])
        res = reader(files.out, 40)
    out, err = capsys.readouterr()
    tomatch = resgen(test_data)
    assert len(res) == len(tomatch)
    for r, t in zip(res, tomatch):  # error msg is easier to read with loop
        assert r == t
    assert res == tomatch
    assert 'Wrote ' + files.out in out
    assert '' == err


@pytest.mark.parametrize('example', ['cfg-example', 'example'])
@pytest.mark.parametrize('refcol', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
# pylint: disable-next=R0913,R0914,R0917
def test_sa_cfg_and_cmd(capsys: CaptureFixture[str], refcol: ColumnRef,
                        example: str) -> None:
    """Test to generate SailArena configuration and use it."""
    with TemporaryDirectory() as dname:
        files = FileNames('./test/test_excel_list_transform/SA_example.csv',
                          dname + '/a.cfg', dname + '/out.xlsx')
        transform_cmd([example, '-r', refcol.name.lower(), '-k', 'sa_to_rrs',
                      '-o', files.cfg])
        transform_cmd(['transform', '-i', files.indata, '-o', files.out,
                      '-c', files.cfg])
        res = openpyxl_reader(files.out, 40)
        tomatch = csv_reader('./test/test_excel_list_transform/' +
                             'SA_result.csv', 40)
        out, err = capsys.readouterr()
        assert len(res) == len(tomatch)
        for r, t in zip(res, tomatch):  # error msg is easier to read with loop
            assert r == t
        assert res == tomatch
        assert 'Wrote ' + files.out in out
        assert '' == err


@pytest.mark.parametrize('refcol', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
def test_gen_example(capsys: CaptureFixture[str], refcol: ColumnRef) -> None:
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
            assert 's07_rename_columns' in txttxt
    out, err = capsys.readouterr()
    assert f'Wrote files {cfg} and {txt}' in out
    assert '' == err


@pytest.mark.parametrize('cfgtype', get_example_names())
@pytest.mark.parametrize('refcol', [ColumnRef.BY_NAME, ColumnRef.BY_NUMBER])
def test_example_txt_layout(capsys: CaptureFixture[str], cfgtype: str,
                            refcol: ColumnRef) -> None:
    """Test generated example syntax text layout."""
    with TemporaryDirectory() as dname:
        cfg = dname + '/a.cfg'
        txt = dname + '/a.txt'
        generate_examplecfg(cfgtype=cfgtype, filename=cfg, colref=refcol)
        with open(file=txt, mode='r', encoding='utf-8') as file:
            _assert_txt_layout(file.read())
    out, err = capsys.readouterr()
    assert f'Wrote files {cfg} and {txt}' in out
    assert '' == err
