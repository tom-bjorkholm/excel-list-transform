#! /usr/local/bin/python3
"""Test the editcfg sub-command and the descriptions it gives the editor."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

import json
import sys
from pathlib import Path
from typing import Optional
import pytest
from pytest import CaptureFixture, MonkeyPatch
from config_as_json import Config, ConfigPath, JsonType
from edit_cfg_json import ConfigLoader, Descriptions, NoEditorError, Settings
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.edit_cfg import NOT_SAVED, SAVED, cfg_loader, \
    chosen_class
from excel_list_transform.edit_descriptions import TABLE_MEMBERS, \
    cfg_descriptions
from excel_list_transform.transform_cmd import transform_cmd

CLASSES: dict[ColumnRef, type[ConfigXlsListTransfName] |
              type[ConfigXlsListTransfNum]] = {
    ColumnRef.BY_NAME: ConfigXlsListTransfName,
    ColumnRef.BY_NUMBER: ConfigXlsListTransfNum}
"""Configuration class for each kind of column reference."""


def json_paths(value: JsonType, prefix: ConfigPath) -> set[ConfigPath]:
    """Return the path of every value below prefix, '[' for list indexes."""
    children: list[tuple[str, JsonType]] = []
    if isinstance(value, dict):
        children = list(value.items())
    elif isinstance(value, list):
        children = [('[', element) for element in value]
    paths: set[ConfigPath] = set()
    for key, child in children:
        path = prefix + (key,)
        paths |= {path} | json_paths(child, path)
    return paths


def is_own_path(path: ConfigPath) -> bool:
    """Return whether a path is outside the TableIO table members."""
    return len(path) == 1 or path[0] not in TABLE_MEMBERS


def example_cfg(folder: Path, ref: ColumnRef) -> Path:
    """Generate an example configuration file and return its path."""
    cfg_file = folder / 'example.cfg'
    transform_cmd(['cfg-example', '-k', 'forms_to_rrs', '-r',
                   ref.name.lower(), '-o', str(cfg_file)])
    return cfg_file


@pytest.mark.parametrize('ref', list(ColumnRef))
def test_descr_match_members(ref: ColumnRef) -> None:
    """Test that every own member and rule key has a description.

    Also test that there is no description of a path the configuration
    does not have. The default configuration holds every kind of rule
    key, so its JSON has every path a configuration can have.
    """
    config = CLASSES[ref]()
    paths = json_paths(json.loads(config.as_json_string(sys.stderr)), ())
    described = set(cfg_descriptions(ref).keys())
    assert {p for p in paths if is_own_path(p)} == \
        {p for p in described if is_own_path(p)}


@pytest.mark.parametrize('ref', list(ColumnRef))
@pytest.mark.parametrize('member', TABLE_MEMBERS)
def test_descr_table_members(ref: ColumnRef, member: str) -> None:
    """Test that the TableIO members of both tables are described."""
    described = cfg_descriptions(ref)
    assert (member, 'format_name') in described
    assert (member, 'csv', 'dialect') in described


@pytest.mark.parametrize('text, expected',
                         [(None, ConfigXlsListTransfName),
                          ('not JSON', ConfigXlsListTransfName),
                          ('[1, 2]', ConfigXlsListTransfName),
                          ('{}', ConfigXlsListTransfName),
                          ('{"column_ref": "by_name"}',
                           ConfigXlsListTransfName),
                          ('{"column_ref": "BY_NUMBER"}',
                           ConfigXlsListTransfNum)])
def test_chosen_class(text: Optional[str], expected: type[Config]) -> None:
    """Test which configuration class a JSON text selects."""
    assert chosen_class(text) is expected


@pytest.mark.parametrize('ref', list(ColumnRef))
def test_cfg_loader(capsys: CaptureFixture[str], tmp_path: Path,
                    ref: ColumnRef) -> None:
    """Test that the loader constructs the class the file selects."""
    text = example_cfg(tmp_path, ref).read_text(encoding='utf-8')
    capsys.readouterr()
    config = cfg_loader(from_json_data_text=text)
    assert isinstance(config, CLASSES[ref])
    assert config.as_json_string(sys.stderr) == \
        CLASSES[ref](from_json_data_text=text).as_json_string(sys.stderr)


@pytest.mark.parametrize('ref, own_text',
                         [(ColumnRef.BY_NAME, 'referencing columns by name'),
                          (ColumnRef.BY_NUMBER,
                           'referencing columns by number')])
@pytest.mark.parametrize('sub_cmd', ['editcfg', 'edit-cfg'])
def test_editcfg_dump(capsys: CaptureFixture[str], tmp_path: Path,
                      ref: ColumnRef, own_text: str, sub_cmd: str) -> None:
    """Test editcfg in the non-interactive editor that prints the model."""
    cfg_file = example_cfg(tmp_path, ref)
    capsys.readouterr()
    transform_cmd([sub_cmd, '--ui', 'dump', '-i', str(cfg_file)[:-4]])
    out, err = capsys.readouterr()
    assert '' == err
    assert own_text in out
    assert cfg_descriptions(ref)[('column_ref',)] in out
    assert out.endswith(NOT_SAVED + '\n')


def test_editcfg_no_file(capsys: CaptureFixture[str], tmp_path: Path) -> None:
    """Test that a missing input file is reported with a way forward."""
    with pytest.raises(SystemExit) as exc_info:
        transform_cmd(['editcfg', '-i', str(tmp_path / 'missing')])
    assert exc_info.value.code == 1
    out, err = capsys.readouterr()
    assert '' == out
    assert 'missing.cfg does not exist' in err
    assert '"cfg-example" sub-command' in err


def test_editcfg_bad_json(capsys: CaptureFixture[str], tmp_path: Path) -> None:
    """Test that a file that is not JSON is reported, not edited."""
    cfg_file = tmp_path / 'bad.cfg'
    cfg_file.write_text('not JSON', encoding='utf-8')
    with pytest.raises(SystemExit) as exc_info:
        transform_cmd(['editcfg', '--ui', 'dump', '-i', str(cfg_file)])
    assert exc_info.value.code == 1
    out, err = capsys.readouterr()
    assert NOT_SAVED not in out
    assert f'Cannot open configuration file {cfg_file} for editing.\n' \
        in err
    assert 'Expecting value: line 1 column 1' in err


@pytest.mark.parametrize('out_args, exp_out',
                         [([], None), (['-o', 'new'], 'new.cfg'),
                          (['--output', 'new.cfg'], 'new.cfg')])
@pytest.mark.parametrize('ui_args, exp_ui',
                         [([], None), (['--ui', 'dump'], 'dump')])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_editcfg_args(capsys: CaptureFixture[str], monkeypatch: MonkeyPatch,
                      tmp_path: Path, out_args: list[str],
                      exp_out: Optional[str], ui_args: list[str],
                      exp_ui: Optional[str]) -> None:
    """Test what editcfg passes to the editor and reports after a save."""
    cfg_file = example_cfg(tmp_path, ColumnRef.BY_NUMBER)
    capsys.readouterr()
    calls: list[Config] = []

    # pylint: disable-next=too-many-arguments
    def fake_edit(config: Config, *, ui_name: Optional[str],
                  descriptions: Descriptions, in_file: str,
                  loader: ConfigLoader, out_file: Optional[str],
                  settings: Settings) -> Config:
        """Check the arguments of edit_in_ui and pretend a save."""
        assert ui_name == exp_ui
        assert descriptions == cfg_descriptions(ColumnRef.BY_NUMBER)
        assert in_file == str(cfg_file)
        assert loader is cfg_loader
        assert out_file == exp_out
        assert settings.file_extension == '.cfg'
        assert settings.extension_enforced
        calls.append(config)
        return config

    monkeypatch.setattr('excel_list_transform.edit_cfg.edit_in_ui', fake_edit)
    transform_cmd(['editcfg', '-i', str(cfg_file)] + out_args + ui_args)
    out, err = capsys.readouterr()
    assert '' == err
    assert SAVED + '\n' == out
    assert len(calls) == 1
    assert isinstance(calls[0], ConfigXlsListTransfNum)


def test_editcfg_no_editor(capsys: CaptureFixture[str],
                           monkeypatch: MonkeyPatch, tmp_path: Path) -> None:
    """Test that a machine with no editor to open is reported."""
    cfg_file = example_cfg(tmp_path, ColumnRef.BY_NAME)
    capsys.readouterr()

    def no_editor(config: Config, **_: object) -> None:
        """Refuse like edit_in_ui on a machine with no editor."""
        assert config is not None
        raise NoEditorError('No editor can run here.')

    monkeypatch.setattr('excel_list_transform.edit_cfg.edit_in_ui', no_editor)
    with pytest.raises(SystemExit) as exc_info:
        transform_cmd(['editcfg', '-i', str(cfg_file)])
    assert exc_info.value.code == 1
    out, err = capsys.readouterr()
    assert '' == out
    assert 'No editor can run here.\n' == err


def test_editcfg_help(capsys: CaptureFixture[str]) -> None:
    """Test help text of the editcfg sub-command."""
    with pytest.raises(SystemExit):
        transform_cmd(['editcfg', '-h'])
    out, err = capsys.readouterr()
    assert '' == err
    assert '--ui {' in out
    assert 'excel-list-transform[textual]' in out
    assert '"cfg-example" sub-command' in out


def test_cmd_return_code(monkeypatch: MonkeyPatch) -> None:
    """Test that a failing sub-command gives its return code as exit code."""
    def failed_migrate(infile: str, outfile: str) -> int:
        """Pretend a failed migration."""
        assert infile == 'old'
        assert outfile == 'new'
        return 3

    monkeypatch.setattr('excel_list_transform.transform_cmd.migrate_cfg',
                        failed_migrate)
    with pytest.raises(SystemExit) as exc_info:
        transform_cmd(['migrate-cfg', '-i', 'old', '-o', 'new'])
    assert exc_info.value.code == 3
