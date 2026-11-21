#! /usr/local/bin/python3
"""Test that constructors and factories forward ok_to_use_defaults."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

import json
import sys
from typing import Callable, Optional, TextIO
import pytest
from pytest import CaptureFixture, MonkeyPatch
from config_as_json import Config, ConfigMissingKey, PathOrStr
from tableio import Capabilities, FileAccess
from tableio_cfg_json import TioJsonConfig
import excel_list_transform.config_excel_list_transform as config_mod
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.edit_cfg import cfg_loader

LOADERS: list[tuple[Callable[..., Config], type[ConfigXlsListTransfName] |
                    type[ConfigXlsListTransfNum]]] = [
    (ConfigXlsListTransfName, ConfigXlsListTransfName),
    (ConfigXlsListTransfNum, ConfigXlsListTransfNum),
    (cfg_loader, ConfigXlsListTransfName),
    (cfg_loader, ConfigXlsListTransfNum)]
"""Each way to load a configuration, with the class it should load."""


def lacking_member(config_class: type[ConfigXlsListTransfName] |
                   type[ConfigXlsListTransfNum]) -> str:
    """Return JSON text of a default configuration without max_column_read."""
    data = json.loads(config_class().as_json_string(sys.stderr))
    del data['max_column_read']
    return json.dumps(data)


def spy_table_config(monkeypatch: MonkeyPatch) -> \
        list[tuple[Optional[str], bool]]:
    """Record each TioJsonConfig that the table factories parse JSON into.

    Returns:
        The member_name and ok_to_use_defaults of each such construction.
    """
    calls: list[tuple[Optional[str], bool]] = []

    # pylint: disable-next=too-many-ancestors
    class SpyTableConfig(TioJsonConfig):
        """TioJsonConfig that records the JSON parsing constructions.

        A subclass and not a function, as the nested_configs() declaration
        requires the config_type to be a type.
        """

        # pylint: disable-next=too-many-arguments
        def __init__(self, *, capabilities: Capabilities,
                     file_access: FileAccess,
                     from_json_data_text: Optional[str],
                     from_json_filename: Optional[PathOrStr],
                     stderr_file: TextIO, member_name: Optional[str],
                     ok_to_use_defaults: bool) -> None:
            """Record the construction and construct the real object."""
            if from_json_data_text is not None:
                calls.append((member_name, ok_to_use_defaults))
            super().__init__(capabilities=capabilities,
                             file_access=file_access,
                             from_json_data_text=from_json_data_text,
                             from_json_filename=from_json_filename,
                             stderr_file=stderr_file, member_name=member_name,
                             ok_to_use_defaults=ok_to_use_defaults)
    monkeypatch.setattr(config_mod, 'TioJsonConfig', SpyTableConfig)
    return calls


@pytest.mark.parametrize('factory', [config_mod.input_table_factory,
                                     config_mod.output_table_factory])
@pytest.mark.parametrize('use_defaults', [False, True])
def test_factory_forwards(monkeypatch: MonkeyPatch,
                          factory: Callable[..., TioJsonConfig],
                          use_defaults: bool) -> None:
    """Test that a table factory passes ok_to_use_defaults on."""
    calls = spy_table_config(monkeypatch)
    factory(from_json_data_text='{"format_name": "CSV"}',
            ok_to_use_defaults=use_defaults)
    assert calls == [(None, use_defaults)]


@pytest.mark.parametrize('config_class', [ConfigXlsListTransfName,
                                          ConfigXlsListTransfNum])
@pytest.mark.parametrize('use_defaults', [False, True])
def test_nested_forwarded(monkeypatch: MonkeyPatch,
                          config_class: type[ConfigXlsListTransfName] |
                          type[ConfigXlsListTransfNum],
                          use_defaults: bool) -> None:
    """Test that a constructor passes ok_to_use_defaults to the factories."""
    text = config_class().as_json_string(sys.stderr)
    calls = spy_table_config(monkeypatch)
    config_class(from_json_data_text=text, ok_to_use_defaults=use_defaults)
    assert calls == [('input_table', use_defaults),
                     ('output_table', use_defaults)]


@pytest.mark.parametrize('construct, config_class', LOADERS)
def test_kept_value(capsys: CaptureFixture[str],
                    construct: Callable[..., Config],
                    config_class: type[ConfigXlsListTransfName] |
                    type[ConfigXlsListTransfNum]) -> None:
    """Test that ok_to_use_defaults keeps the default of a missing value."""
    config = construct(from_json_data_text=lacking_member(config_class),
                       stderr_file=sys.stderr, ok_to_use_defaults=True)
    assert isinstance(config, config_class)
    assert config.kept_value_paths() == [('max_column_read',)]
    assert config.as_json_string(sys.stderr) == \
        config_class().as_json_string(sys.stderr)
    assert capsys.readouterr() == ('', '')


@pytest.mark.parametrize('construct, config_class', LOADERS)
@pytest.mark.parametrize('keywords', [{}, {'ok_to_use_defaults': False}])
def test_missing_is_error(capsys: CaptureFixture[str],
                          construct: Callable[..., Config],
                          config_class: type[ConfigXlsListTransfName] |
                          type[ConfigXlsListTransfNum],
                          keywords: dict[str, bool]) -> None:
    """Test that a missing value is an error without ok_to_use_defaults."""
    with pytest.raises(ConfigMissingKey):
        construct(from_json_data_text=lacking_member(config_class),
                  stderr_file=sys.stderr, **keywords)
    assert 'max_column_read' in capsys.readouterr().err
