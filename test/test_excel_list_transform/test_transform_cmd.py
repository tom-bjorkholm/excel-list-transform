#! /usr/local/bin/python3
"""Test the excel_list_transform command parsing functionality."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from copy import deepcopy
from datetime import date
from importlib.metadata import version as metadata_version
from io import StringIO
from tempfile import TemporaryDirectory
from typing import Optional
import sys
from config_as_json import InvalidConfiguration
from packaging.version import Version
import pytest
from pytest import MonkeyPatch
from pytest import CaptureFixture
from test_excel_list_transform.tableio_helpers import read_excel_num, \
    write_excel_num
from excel_list_transform.transform_cmd import transform_cmd
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.version_information import ExcelListVersionReporter
from excel_list_transform.commontypes import NumData


@pytest.mark.smoke
@pytest.mark.parametrize('arg1', ['example', 'cfg-example'])
@pytest.mark.parametrize('arg2, exp_out, exp_in, exp_cfg, ref',
                         [[['-o', 'ofile', '-k', 'example',
                            '-r', 'by_name'],
                           'ofile', None, 'example', ColumnRef.BY_NAME],
                          [['-o', 'ofile', '-k', 'example',
                            '-r', 'by_number'],
                           'ofile', None, 'example', ColumnRef.BY_NUMBER],
                          [['--output', 'ofile', '--kind',
                            'example', '--reference', 'by_name'],
                           'ofile', None, 'example', ColumnRef.BY_NAME]])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_xls_lst_rfm_cmd_smk1(capsys: CaptureFixture[str],
                              monkeypatch: MonkeyPatch, arg1: str,
                              arg2: list[str], exp_out: str,
                              exp_in: Optional[str], exp_cfg: str,
                              ref: ColumnRef) -> None:
    """Test the excel_list_transform command parsing functionality."""
    full_args = [deepcopy(arg1)] + deepcopy(arg2)
    calls = {'gen': 0, 'refor': 0}

    def gen(filename: str, cfgtype: str, colref: ColumnRef) -> None:
        """Mock generate_examplecfg."""
        assert filename == exp_out
        assert cfgtype == exp_cfg
        assert colref == ref
        calls['gen'] += 1

    def refor(infilename: str, outfilename: str, cfgfilename: str) -> None:
        """Mock transform_named_files."""
        calls['refor'] += 1
        assert infilename == exp_in
        assert outfilename == exp_out
        assert cfgfilename == exp_cfg
    mod = 'excel_list_transform.transform_cmd.'
    monkeypatch.setattr(mod + 'generate_examplecfg', gen)
    monkeypatch.setattr(mod + 'transform_named_files', refor)
    transform_cmd(arguments=full_args)
    out, err = capsys.readouterr()
    assert calls['gen'] == 1
    assert calls['refor'] == 0
    assert '' == out
    assert '' == err


@pytest.mark.smoke
@pytest.mark.parametrize('args, exp_out, exp_in, exp_cfg, ref',
                         [[['transform', '-o', 'ofile', '-i', 'ifile', '-c',
                            'conf'],
                           'ofile', 'ifile', 'conf', None],
                          [['transform', '--output', 'of', '--input', 'ifi',
                            '--cfg', 'con'],
                           'of', 'ifi', 'con', None],
                          [['python3', '-m', 'module.py',
                            'transform', '--output', 'of', '--input', 'ifi',
                            '--cfg', 'con'],
                           'of', 'ifi', 'con', None]])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_xls_lst_rfm_cmd_smk2(capsys: CaptureFixture[str],
                              monkeypatch: MonkeyPatch, args: list[str],
                              exp_out: str, exp_in: str, exp_cfg: str,
                              ref: Optional[ColumnRef]) -> None:
    """Test the excel_list_transform command parsing functionality."""
    calls = {'gen': 0, 'refor': 0}

    def gen(filename: str, cfgtype: str, colref: ColumnRef) -> None:
        """Mock generate_examplecfg."""
        assert filename == exp_out
        assert cfgtype == exp_cfg
        assert colref == ref
        calls['gen'] += 1

    def refor(infilename: str, outfilename: str, cfgfilename: str) -> None:
        """Mock transform_named_files."""
        calls['refor'] += 1
        assert infilename == exp_in
        assert outfilename == exp_out
        assert cfgfilename == exp_cfg
    mod = 'excel_list_transform.transform_cmd.'
    monkeypatch.setattr(mod + 'generate_examplecfg', gen)
    monkeypatch.setattr(mod + 'transform_named_files', refor)
    transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert calls['gen'] == 0
    assert calls['refor'] == 1
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('config_error, expected',
                         [(KeyError('Missing column'), 'Missing column'),
                          (InvalidConfiguration('Bad config'),
                           'Bad config')])
def test_transform_cfg_error(config_error: Exception, expected: str,
                             monkeypatch: MonkeyPatch,
                             capsys: CaptureFixture[str]) -> None:
    """Test transform command reports readable configuration errors."""
    def fail_transform(infilename: str, outfilename: str,
                       cfgfilename: str) -> int:
        """Mock transform function that reports a configuration error."""
        assert infilename == 'ifile'
        assert outfilename == 'ofile'
        assert cfgfilename == 'conf'
        raise config_error
    mod = 'excel_list_transform.transform_cmd.transform_named_files'
    monkeypatch.setattr(mod, fail_transform)
    err_file = StringIO()
    monkeypatch.setattr('excel_list_transform.transform_cmd.sys_stderr',
                        err_file)
    with pytest.raises(SystemExit) as exc_info:
        transform_cmd(['transform', '-o', 'ofile', '-i', 'ifile',
                       '-c', 'conf'])
    out, err = capsys.readouterr()
    assert exc_info.value.code == 1
    assert out == ''
    assert err == ''
    assert err_file.getvalue() == expected + '\n'


@pytest.mark.parametrize('args, errs',
                         [[['transform', '-i', 'ifile', '-c', 'conf'],
                           'arguments are required: -o/--output'],
                          [['example', '-r', 'by_name', '-k', 'example'],
                           'arguments are required: -o/--output'],
                          [['cfg-example', '-r', 'by_name', '-k', 'example'],
                           'arguments are required: -o/--output'],
                          [['-i', 'ifile', '-o', 'ofile'],
                           "invalid choice: 'ifile'"],
                          [['example', '-k', 'example', '-r', 'by_number'],
                           'required: -o/--output'],
                          [['example', '--output', 'of', '-i', 'in',
                            '-k', 'con', '-r', 'by_name'],
                           'invalid choice: \'con\''],
                          [['cfg-example', '-k', 'example', '-r',
                            'by_number'],
                           'required: -o/--output'],
                          [['cfg-example', '--output', 'of', '-i', 'in',
                            '-k', 'con', '-r', 'by_name'],
                           'invalid choice: \'con\'']])
def test_xls_lst_rfm_cmd_err(capsys: CaptureFixture[str], args: list[str],
                             errs: str) -> None:
    """Test the excel_list_transform command parsing errors."""
    with pytest.raises(SystemExit):
        transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '' == out
    assert errs in err


@pytest.mark.parametrize('args', [['-h'], ['--help']])
def test_xls_lst_rfm_cmd_help(capsys: CaptureFixture[str],
                              args: list[str]) -> None:
    """Test the excel_list_transform command parsing help."""
    with pytest.raises(SystemExit) as _:
        transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '-h, --help' in out
    assert 'transform' in out
    assert 'example' in out
    assert 'version' in out
    assert 'Print versions of excel_list_transform' in out
    assert 'Generate example configuration file (example .cfg' in out
    assert 'Transform list in excel or CSV file. How data is' in out
    assert 'More detailed help is available for each sub-command.' \
        in out.replace('\n', ' ')
    assert '' == err


@pytest.mark.parametrize('arg1', ['example', 'cfg-example'])
@pytest.mark.parametrize('arg2', ['-h', '--help'])
def test_xlsr_cmd_ex_help(capsys: CaptureFixture[str], arg1: str,
                          arg2: str) -> None:
    """Test the excel_list_transform command parsing help."""
    args = [deepcopy(arg1), deepcopy(arg2)]
    with pytest.raises(SystemExit) as _:
        transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '-h, --help' in out
    assert ' -k ' in out
    assert ' --kind ' in out
    assert ' -r ' in out
    assert ' --reference ' in out
    assert 'example' in out
    assert '' == err


@pytest.mark.parametrize('args', [['transform', '-h'],
                                  ['transform', '--help']])
def test_xlsr_cmd_rfmt_help(capsys: CaptureFixture[str],
                            args: list[str]) -> None:
    """Test the excel_list_transform command parsing help."""
    with pytest.raises(SystemExit) as _:
        transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '-h, --help' in out
    assert ' -c, --cfg CFG' in out
    assert ' -i, --input INPUT' in out
    assert ' -o, --output OUTPUT' in out
    assert 'example' in out
    assert 'list in excel or CSV file' in out
    assert '' == err


@pytest.mark.parametrize('args', [['version', '-h'],
                                  ['version', '--help']])
def test_xlsr_cmd_vers_help(capsys: CaptureFixture[str],
                            args: list[str]) -> None:
    """Test the excel_list_transform command parsing help."""
    with pytest.raises(SystemExit) as _:
        transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '-h, --help' in out
    assert 'Print versions of excel_list_transform' in out
    assert '' == err


def test_version_cmd1(capsys: CaptureFixture[str],
                      monkeypatch: MonkeyPatch) -> None:
    """Test command to print version information."""
    def no_new_packages(_: ExcelListVersionReporter,
                        versions: dict[str, Version]) -> None:
        """Skip PyPI lookup in this command test."""
        assert versions is not None

    mod = 'excel_list_transform.version_information.'
    mod += 'ExcelListVersionReporter._print_info_on_new_pkgs'
    monkeypatch.setattr(mod, no_new_packages)
    transform_cmd(['version'])
    out, err = capsys.readouterr()
    python_version = '.'.join([str(s) for s in sys.version_info[0:3]])
    assert '' == err
    assert f'Python .............. {python_version}' in out
    assert f'excel_list_transform  {metadata_version("excel_list_transform")}'\
        in out


@pytest.mark.parametrize('ver, dat, errprint',
                         [((3, 11, 1, 0, 0),
                           date(year=2026, month=12, day=25), True),
                          ((3, 11, 1, 0, 0),
                           date(year=2024, month=12, day=25), False),
                          ((3, 10, 11, 75, 0),
                           date(year=2027, month=12, day=25), True)])
def test_cmd_ver_check_if_u(capsys: CaptureFixture[str],
                            monkeypatch: MonkeyPatch,
                            ver: tuple[int, int, int, int, int], dat: date,
                            errprint: bool) -> None:
    """Test version check if unsupported python widh old Python."""
    mod1 = 'versionreporter.versionreporter.sys.version_info'
    monkeypatch.setattr(mod1, ver)

    def mock_day(_: object) -> date:
        """Mock Version._today."""
        assert isinstance(dat, date)
        return dat
    mod2 = 'excel_list_transform.version_information.'
    mod2 += 'ExcelListVersionReporter._today'
    monkeypatch.setattr(mod2, mock_day)
    with pytest.raises(SystemExit):
        transform_cmd(['--help'])
    out, err = capsys.readouterr()
    assert '' == err
    if errprint:
        assert 'You are running an old version of Python:' in out
        assert 'This application no longer releases bug fixes ' in out
        assert 'for this old Python version.' in out
        assert 'Upgrade Python to a new version.' in out
        assert '(Download Python from https://www.python.org/downloads' in out
        assert 'After installing new Python, upgrade application with' in out
        assert ' install --upgrade ' in out
    else:
        assert 'You are running an old version of Python:' not in out
        assert 'This application no longer releases bug fixes ' not in out
        assert 'for this old Python version.' not in out
        assert 'Upgrade Python to a new version.' not in out
        assert '(Download Python from https://www.python.or' not in out
        assert 'After installing new Python, upgrade application ' not in out
        assert ' install --upgrade ' not in out


@pytest.mark.parametrize('ref', list(ColumnRef))
def test_row_spl_mrg_cmd_cfg(capsys: CaptureFixture[str],
                             ref: ColumnRef) -> None:
    """Test row split and merge config genaration."""
    indata: NumData = [['From', 'What', 'To'],
                       ['Gardener', 'Apples', 'Jones + Smith'],
                       ['Brewery', 'Beer', 'Smith + Bush'],
                       ['Dairy', 'Milk', 'Jones']]
    outdata = [['To', 'What', 'From'],
               ['Jones', 'Apples + Milk', 'Gardener + Dairy'],
               ['Smith', 'Apples + Beer', 'Gardener + Brewery'],
               ['Bush', 'Beer', 'Brewery']]
    with TemporaryDirectory() as dirname:
        cfgfile = dirname + '/a.cfg'
        infile = dirname + '/in.xlsx'
        outfile = dirname + '/out.xlsx'
        transform_cmd(['cfg-example', '-k', 'row_split_merge',
                       '-r', ref.name.lower(), '-o', cfgfile])
        with open(file=dirname + '/a.txt', mode='r', encoding='utf-8') as f:
            txt = f.read()
            assert '"s01_split_rows"' in txt
            assert '"s02_merge_rows"' in txt
            assert 'Gardener + Dairy' in txt
        write_excel_num(data=deepcopy(indata), filename=infile)
        transform_cmd(['transform', '-c', cfgfile, '-i', infile,
                       '-o', outfile])
        res = read_excel_num(filename=outfile, max_col_read=20,
                             strip_col_names=False, strip_values=False)
        assert res == outdata
    out, err = capsys.readouterr()
    assert '' == err
    assert 'Wrote files' in out
    assert 'a.cfg' in out
    assert 'a.txt' in out
    assert 'out.xlsx' in out
