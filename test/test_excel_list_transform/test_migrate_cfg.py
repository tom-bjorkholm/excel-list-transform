#! /usr/local/bin/python3
"""Test migrate_cfg."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

import json
import sys
from tempfile import TemporaryDirectory
import pytest
from pytest import MonkeyPatch
from pytest import CaptureFixture
from config_as_json import Config, config_factory_from_json
from tableio import CsvDialect
from tableio_cfg_json import TioJsonCsvConfig
from test_excel_list_transform.tableio_helpers import \
    configure_output_csv
from excel_list_transform.migrate_cfg import migrate_cfg
from excel_list_transform.migrate_cfg_warn_hook import MigrateCfgWarnHook
from excel_list_transform.config_match import MATCH_CONFIGS
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.assert_dict_equal import assert_dict_equal
from excel_list_transform.transform_cmd import transform_cmd


def configure_old_io(cfg: ConfigXlsListTransfName |
                     ConfigXlsListTransfNum) -> None:
    """Set expected TableIO settings from old migration fixtures."""
    configure_output_csv(cfg)
    cfg.input_table.character_encoding = 'utf_8_sig'
    cfg.output_table.character_encoding = 'utf-8'
    cfg.input_table.csv = TioJsonCsvConfig(dialect=CsvDialect.UNIX,
                                           delimiter=',', quotechar='"')
    cfg.output_table.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                            delimiter=',', quotechar='"')


def assert_new_config_file(filename: str) -> None:
    """Assert that a migrated file uses the new config shape."""
    with open(file=filename, mode='r', encoding='utf-8') as file:
        data = json.load(file)
    assert 'input_table' in data
    assert 'output_table' in data
    assert data['input_table'].get('implementation') is None
    assert data['output_table'].get('implementation') is None
    old_keys = ['in_type', 'out_type', 'in_csv_dialect',
                'out_csv_dialect', 'in_csv_encoding', 'out_csv_encoding',
                'in_excel_library', 'out_excel_library']
    for old_key in old_keys:
        assert old_key not in data


def read_config(filename: str) -> Config:
    """Read an app config file through the config-as-json factory."""
    return config_factory_from_json(match_configs=MATCH_CONFIGS,
                                    auto_ch_hook=MigrateCfgWarnHook(),
                                    from_json_filename=filename,
                                    stderr_file=sys.stderr)


def test_migrate_cfg1(capsys: CaptureFixture[str]) -> None:
    """Test migration of 0.7.13 config file."""
    refcfg = ConfigXlsListTransfName()
    configure_old_io(refcfg)
    refcfg.s01_split_rows = []
    refcfg.s02_merge_rows = []
    refcfg.s03_split_columns[0]['right_name'] = 'Family Name'
    refcfg.s08_insert_columns[1]['column'] = 'Something Else'
    infilename = 'test/test_excel_list_transform/bak_compat_0_7_13_name.cfg'
    with TemporaryDirectory() as dirname:
        outfilename = dirname + '/a.cfg'
        _ = read_config(infilename)
        _, err = capsys.readouterr()
        assert err == MigrateCfgWarnHook.migrate_warn_msg()
        res = migrate_cfg(infile=infilename, outfile=outfilename)
        cfg = read_config(outfilename)
        assert isinstance(cfg, ConfigXlsListTransfName)
        assert_new_config_file(outfilename)
        out, err = capsys.readouterr()
        assert_dict_equal(refcfg.__dict__, cfg.__dict__,
                          ['_hook_cfg_autochange'])
        assert cfg.s03_split_columns[0]['right_name'] == 'Family Name'
        assert cfg.s08_insert_columns[1]['column'] == 'Something Else'
        assert '' == out
        assert '' == err
        assert res == 0


def test_migrate_cfg2(capsys: CaptureFixture[str]) -> None:
    """Test migration of 0.7.13 config file."""
    refcfg = ConfigXlsListTransfNum()
    configure_old_io(refcfg)
    refcfg.s01_split_rows = []
    refcfg.s02_merge_rows = []
    refcfg.s07_rename_columns[1]['name'] = 'Family Name'
    refcfg.s08_insert_columns[1]['name'] = 'Something else'
    infilename = 'test/test_excel_list_transform/bak_compat_0_7_13_number.cfg'
    with TemporaryDirectory() as dirname:
        outfilename = dirname + '/a.cfg'
        _ = read_config(infilename)
        _, err = capsys.readouterr()
        assert err == MigrateCfgWarnHook.migrate_warn_msg()
        res = migrate_cfg(infile=infilename, outfile=outfilename)
        cfg = read_config(outfilename)
        assert isinstance(cfg, ConfigXlsListTransfNum)
        assert_new_config_file(outfilename)
        out, err = capsys.readouterr()
        assert_dict_equal(refcfg.__dict__, cfg.__dict__,
                          ['_hook_cfg_autochange'])
        assert cfg.s07_rename_columns[1]['name'] == 'Family Name'
        assert cfg.s08_insert_columns[1]['name'] == 'Something else'
        assert '' == out
        assert '' == err
        assert res == 0


def test_migrate_cfg_nok1(capsys: CaptureFixture[str]) -> None:
    """Test not OK case 1 of migrate_cfg."""
    with TemporaryDirectory() as dirname:
        infilename = dirname + '/a.cfg'
        outfilename = dirname + '/b.cfg'
        with pytest.raises(SystemExit):
            _ = migrate_cfg(infile=infilename, outfile=outfilename)
    out, err = capsys.readouterr()
    assert '' == out
    assert 'Cannot find input configuration file' in err


def test_migrate_cfg_nok2(capsys: CaptureFixture[str]) -> None:
    """Test not OK case 2 of migrate_cfg."""
    cfg = ConfigXlsListTransfNum()
    with TemporaryDirectory() as dirname:
        infilename = dirname + '/a.cfg'
        outfilename = dirname + '/b.cfg'
        cfg.write(to_json_filename=infilename)
        cfg.write(to_json_filename=outfilename)
        with pytest.raises(SystemExit):
            _ = migrate_cfg(infile=infilename, outfile=outfilename)
    out, err = capsys.readouterr()
    assert '' == out
    assert 'Output configuration file' in err
    assert 'already exists.' in err
    assert 'Cowardly refusing to overwrite existing configuration file' in err


@pytest.mark.parametrize('ipar', ['-i', '--input'])
@pytest.mark.parametrize('opar', ['-o', '--output'])
def test_migrate_cmd1(capsys: CaptureFixture[str], ipar: str,
                      opar: str) -> None:
    """Test migration of 0.7.13 config file."""
    refcfg = ConfigXlsListTransfName()
    configure_old_io(refcfg)
    refcfg.s01_split_rows = []
    refcfg.s02_merge_rows = []
    refcfg.s03_split_columns[0]['right_name'] = 'Family Name'
    refcfg.s08_insert_columns[1]['column'] = 'Something Else'
    infilename = 'test/test_excel_list_transform/bak_compat_0_7_13_name.cfg'
    with TemporaryDirectory() as dirname:
        outfilename = dirname + '/a.cfg'
        args = ['migrate-cfg', ipar, infilename, opar, outfilename]
        transform_cmd(arguments=args)
        cfg = read_config(outfilename)
        assert isinstance(cfg, ConfigXlsListTransfName)
        assert_new_config_file(outfilename)
        out, err = capsys.readouterr()
        assert_dict_equal(refcfg.__dict__, cfg.__dict__,
                          ['_hook_cfg_autochange'])
        assert cfg.s03_split_columns[0]['right_name'] == 'Family Name'
        assert cfg.s08_insert_columns[1]['column'] == 'Something Else'
        assert '' == out
        assert '' == err


@pytest.mark.parametrize('ipar', ['-i', '--input'])
@pytest.mark.parametrize('opar', ['-o', '--output'])
def test_migrate_cmd2(capsys: CaptureFixture[str], ipar: str,
                      opar: str) -> None:
    """Test migration of 0.7.13 config file."""
    refcfg = ConfigXlsListTransfNum()
    configure_old_io(refcfg)
    refcfg.s01_split_rows = []
    refcfg.s02_merge_rows = []
    refcfg.s07_rename_columns[1]['name'] = 'Family Name'
    refcfg.s08_insert_columns[1]['name'] = 'Something else'
    infilename = 'test/test_excel_list_transform/bak_compat_0_7_13_number.cfg'
    with TemporaryDirectory() as dirname:
        outfilename = dirname + '/a.cfg'
        args = ['migrate-cfg', ipar, infilename, opar, outfilename]
        transform_cmd(arguments=args)
        cfg = read_config(outfilename)
        assert isinstance(cfg, ConfigXlsListTransfNum)
        assert_new_config_file(outfilename)
        out, err = capsys.readouterr()
        assert_dict_equal(refcfg.__dict__, cfg.__dict__,
                          ['_hook_cfg_autochange'])
        assert cfg.s07_rename_columns[1]['name'] == 'Family Name'
        assert cfg.s08_insert_columns[1]['name'] == 'Something else'
        assert '' == out
        assert '' == err


@pytest.mark.parametrize('ipar', ['-i', '--input'])
@pytest.mark.parametrize('opar', ['-o', '--output'])
def test_migrate_cmd_nok1(capsys: CaptureFixture[str], ipar: str,
                          opar: str) -> None:
    """Test not OK case 1 of migrate_cmd."""
    with TemporaryDirectory() as dirname:
        infilename = dirname + '/a.cfg'
        outfilename = dirname + '/b.cfg'
        with pytest.raises(SystemExit):
            args = ['migrate-cfg', ipar, infilename, opar, outfilename]
            transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '' == out
    assert 'Cannot find input configuration file' in err


@pytest.mark.parametrize('ipar', ['-i', '--input'])
@pytest.mark.parametrize('opar', ['-o', '--output'])
def test_migrate_cmd_nok2(capsys: CaptureFixture[str], ipar: str,
                          opar: str) -> None:
    """Test not OK case 2 of migrate_cmd."""
    cfg = ConfigXlsListTransfNum()
    with TemporaryDirectory() as dirname:
        infilename = dirname + '/a.cfg'
        outfilename = dirname + '/b.cfg'
        cfg.write(to_json_filename=infilename)
        cfg.write(to_json_filename=outfilename)
        with pytest.raises(SystemExit):
            args = ['migrate-cfg', ipar, infilename, opar, outfilename]
            transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '' == out
    assert 'Output configuration file' in err
    assert 'already exists.' in err
    assert 'Cowardly refusing to overwrite existing configuration file' in err


def test_migrate_cmd_help1(capsys: CaptureFixture[str]) -> None:
    """Test help including migrate-cfg."""
    with pytest.raises(SystemExit):
        transform_cmd(['-h'])
    out, err = capsys.readouterr()
    assert '' == err
    assert ',migrate-cfg}' in out
    assert '  migrate-cfg  ' in out
    assert 'Migrate configuration file format from older format' in out
    assert 'newest format.' in out


def test_migrate_cmd_help2(capsys: CaptureFixture[str]) -> None:
    """Test help for migrate-cfg."""
    with pytest.raises(SystemExit):
        transform_cmd(['migrate-cfg', '-h'])
    out, err = capsys.readouterr()
    assert '' == err
    assert 'Migrate configuration file format from older format' in out
    assert 'newest format.' in out
    assert 'can be read by the next version of the command' in out


@pytest.mark.parametrize('ipar', ['-i', '--input'])
@pytest.mark.parametrize('opar', ['-o', '--output'])
@pytest.mark.parametrize('ival', ['ab', 'cd'])
@pytest.mark.parametrize('oval', ['ef', 'gh'])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_migrate_cmd3(capsys: CaptureFixture[str], monkeypatch: MonkeyPatch,
                      ipar: str, opar: str, ival: str, oval: str) -> None:
    """Test command line calling migrate_cfg."""
    migrate_calls = 0

    def migrate_mock(infile: str, outfile: str) -> int:
        """Mock of migrate_cfg."""
        nonlocal migrate_calls
        migrate_calls += 1
        assert infile == ival
        assert outfile == oval
        return 0

    to_mock = 'excel_list_transform.transform_cmd.migrate_cfg'
    monkeypatch.setattr(to_mock, migrate_mock)
    args = ['migrate-cfg', ipar, ival, opar, oval]
    transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert 1 == migrate_calls
    assert '' == out
    assert '' == err
