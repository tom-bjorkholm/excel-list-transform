#! /usr/local/bin/python3
"""Test the excel_list_transform command parsing functionality."""

# Copyright (c) 2024 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from copy import deepcopy
import pytest
from excel_list_transform.transform_cmd import transform_cmd
from excel_list_transform.config_enums import ColumnRef


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
def test_excel_list_rfm_cmd_smok1(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments,line-too-long # noqa: E501
                                  monkeypatch, arg1, arg2, exp_out, exp_in,
                                  exp_cfg, ref):
    """Test the excel_list_transform command parsing functionality."""
    full_args = [deepcopy(arg1)] + deepcopy(arg2)

    def gen(filename, cfgtype, colref):
        """Mock generate_examplecfg."""
        assert filename == exp_out
        assert cfgtype == exp_cfg
        assert colref == ref
        gen.num_called += 1
    gen.num_called = 0

    def refor(infilename, outfilename, cfgfilename):
        """Mock transform_named_files."""
        refor.num_called += 1
        assert infilename == exp_in
        assert outfilename == exp_out
        assert cfgfilename == exp_cfg
    refor.num_called = 0
    mod = 'excel_list_transform.transform_cmd.'
    monkeypatch.setattr(mod + 'generate_examplecfg', gen)
    monkeypatch.setattr(mod + 'transform_named_files', refor)
    transform_cmd(arguments=full_args)
    out, err = capsys.readouterr()
    assert gen.num_called == 1
    assert refor.num_called == 0
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
def test_excel_list_rfm_cmd_smok2(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments,line-too-long # noqa: E501
                                  monkeypatch, args, exp_out, exp_in, exp_cfg,
                                  ref):
    """Test the excel_list_transform command parsing functionality."""
    def gen(filename, cfgtype, colref):
        """Mock generate_examplecfg."""
        assert filename == exp_out
        assert cfgtype == exp_cfg
        assert colref == ref
        gen.num_called += 1
    gen.num_called = 0

    def refor(infilename, outfilename, cfgfilename):
        """Mock transform_named_files."""
        refor.num_called += 1
        assert infilename == exp_in
        assert outfilename == exp_out
        assert cfgfilename == exp_cfg
    refor.num_called = 0
    mod = 'excel_list_transform.transform_cmd.'
    monkeypatch.setattr(mod + 'generate_examplecfg', gen)
    monkeypatch.setattr(mod + 'transform_named_files', refor)
    transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert gen.num_called == 0
    assert refor.num_called == 1
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('args, errs',
                         [[['transform', '-i', 'ifile', '-c', 'conf'],
                           'arguments are required: -o/--output'],
                          [['example', '-r', 'by_name', '-k', 'example'],
                           'arguments are required: -o/--output'],
                          [['cfg-example', '-r', 'by_name', '-k', 'example'],
                           'arguments are required: -o/--output'],
                          [['-i', 'ifile', '-o', 'ofile'],
                           "(choose from 'example', 'cfg-example', " +
                           "'transform')"],
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
def test_excel_list_rfm_cmd_err(capsys, args, errs):
    """Test the excel_list_transform command parsing errors."""
    with pytest.raises(SystemExit):
        transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '' == out
    assert errs in err


@pytest.mark.parametrize('args', [['-h'], ['--help']])
def test_excel_list_rfm_cmd_help(capsys, args):
    """Test the excel_list_transform command parsing help."""
    with pytest.raises(SystemExit) as _:
        transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '-h, --help' in out
    assert 'transform' in out
    assert 'example' in out
    assert 'Generate example configuration file (example .cfg' in out
    assert 'Transform list in excel or CSV file. How data is' in out
    assert 'More detailed help is available for each sub-command.' \
        in out.replace('\n', ' ')
    assert '' == err


@pytest.mark.parametrize('arg1', ['example', 'cfg-example'])
@pytest.mark.parametrize('arg2', ['-h', '--help'])
def test_xlsr_cmd_example_help(capsys, arg1, arg2):
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
def test_xlsr_cmd_rfmt_help(capsys, args):
    """Test the excel_list_transform command parsing help."""
    with pytest.raises(SystemExit) as _:
        transform_cmd(arguments=args)
    out, err = capsys.readouterr()
    assert '-h, --help' in out
    assert ' -c CFG, --cfg CFG' in out
    assert ' -i INPUT, --input INPUT' in out
    assert ' -o OUTPUT, --output OUTPUT' in out
    assert 'example' in out
    assert 'list in excel or CSV file' in out
    assert '' == err
