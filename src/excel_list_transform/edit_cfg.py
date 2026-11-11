#! /usr/local/bin/python3
"""Edit a configuration file in whichever configuration editor can run."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from dataclasses import replace
from json import JSONDecodeError, loads as json_loads
from os import path
from pathlib import Path
from typing import Optional, TextIO
import sys
from config_as_json import Config, PathOrStr
from edit_cfg_json import ConfigLoadError, NoEditorError, Settings, \
    derived_loader, edit_in_ui, load_settings
from excel_list_transform.config_match import MATCH_CONFIGS
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.edit_descriptions import cfg_descriptions
from excel_list_transform.file_extension import fix_file_extension

CFG_EXTENSION = '.cfg'
"""File extension of configuration files of this application."""

NO_FILE = 'Configuration file {name} does not exist. Use the ' + \
    '"cfg-example" sub-command to create a configuration file to edit.'
"""Message for an input file that does not exist."""

CANNOT_OPEN = 'Cannot open configuration file {name} for editing.'
"""First line of the message for an input file that cannot be read."""

SAVED = 'The configuration was saved.'
"""Message after an editing session that saved the configuration."""

NOT_SAVED = 'The configuration was not saved.'
"""Message after an editing session that saved nothing."""


def chosen_class(text: Optional[str]) -> type[ConfigXlsListTransfName] | \
        type[ConfigXlsListTransfNum]:
    """Return the configuration class that one JSON text selects.

    Text that is not a JSON object, or that selects no class, gives the
    class referencing columns by name, so that the configuration class
    itself reports what is wrong with the text.

    Args:
        text: JSON text of a configuration file, or None for no file.

    Returns:
        The configuration class to construct.
    """
    if text is None:
        return ConfigXlsListTransfName
    try:
        data = json_loads(text)
    except JSONDecodeError:
        return ConfigXlsListTransfName
    if not isinstance(data, dict):
        return ConfigXlsListTransfName
    for match in MATCH_CONFIGS:
        if match.match_func(text, sys.stderr):
            assert issubclass(match.config_class, (ConfigXlsListTransfName,
                                                   ConfigXlsListTransfNum))
            return match.config_class
    return ConfigXlsListTransfName


def cfg_loader(*, from_json_data_text: Optional[str] = None,
               from_json_filename: Optional[PathOrStr] = None,
               ok_to_use_defaults: bool = False,
               stderr_file: TextIO = sys.stderr) -> Config:
    """Construct the configuration class that the JSON text selects.

    This is the loader given to the editor. As the editor asks it again
    when saving, a save that would change which class the file selects
    (by changing column_ref) is refused instead of written.

    Args:
        from_json_data_text: JSON text to apply, or None for the defaults.
        from_json_filename: File to read (never passed by the editor).
        ok_to_use_defaults: Whether defaults may fill in missing members.
        stderr_file: Stream used for user-facing diagnostics.

    Returns:
        A configuration that references columns by name or by number.
    """
    loader = derived_loader(chosen_class(from_json_data_text))
    return loader(from_json_data_text=from_json_data_text,
                  from_json_filename=from_json_filename,
                  ok_to_use_defaults=ok_to_use_defaults,
                  stderr_file=stderr_file)


def editor_settings() -> Settings:
    """Return the user's editor settings with this application's file names.

    Raises:
        ConfigLoadError: The user's editor settings file cannot be read.
    """
    return replace(load_settings(), file_extension=CFG_EXTENSION,
                   extension_enforced=True)


def initial_config(filename: str) -> ConfigXlsListTransfName | \
        ConfigXlsListTransfNum:
    """Return a default configuration of the class the named file selects.

    Undecodable bytes are replaced, as the editor reports them when it
    reads the file.
    """
    text = Path(filename).read_text(encoding='utf-8', errors='replace')
    return chosen_class(text)()


def edit_cfg(infile: str, outfile: Optional[str],
             ui_name: Optional[str]) -> int:
    """Edit a configuration file in a configuration editor.

    Args:
        infile: Configuration file to edit (.cfg is added if needed).
        outfile: File to save to, or None to save to the input file.
        ui_name: User interface to edit in, or None for the best one.

    Returns:
        0 when the editor ran, 1 when it could not be opened.
    """
    in_name = fix_file_extension(infile, CFG_EXTENSION, for_reading=True)
    if not path.isfile(in_name):
        print(NO_FILE.format(name=in_name), file=sys.stderr)
        return 1
    out_name = None if outfile is None else \
        fix_file_extension(outfile, CFG_EXTENSION)
    config = initial_config(in_name)
    try:
        saved = edit_in_ui(config, ui_name=ui_name,
                           descriptions=cfg_descriptions(config.column_ref),
                           in_file=in_name, loader=cfg_loader,
                           out_file=out_name, settings=editor_settings())
    except ConfigLoadError as error:
        print(CANNOT_OPEN.format(name=in_name), error, sep='\n',
              file=sys.stderr)
        return 1
    except NoEditorError as error:
        print(str(error), file=sys.stderr)
        return 1
    print(NOT_SAVED if saved is None else SAVED)
    return 0
