#! /usr/local/bin/python3
"""Get correct configuration type."""

# Copyright (c) 2024 Tom Björkholm
# MIT License


from typing import Optional, TypeAlias
import sys
from json import loads as json_loads
from json import JSONDecodeError
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.config_xls_list_refmt_name import \
    ConfigXlsListRefmtName
from excel_list_transform.config_xls_list_refmt_num import \
    ConfigXlsListRefmtNum
from excel_list_transform.str_to_enum import string_to_enum_best_match
from excel_list_transform.commontypes import JsonType
from excel_list_transform.file_must_exist import file_must_exist


Configs: TypeAlias = \
    dict[ColumnRef,
         type[ConfigXlsListRefmtName] | type[ConfigXlsListRefmtNum]]
DerivedConfig: TypeAlias = ConfigXlsListRefmtName | ConfigXlsListRefmtNum

_CONFIGS: Configs = {ColumnRef.BY_NAME: ConfigXlsListRefmtName,
                     ColumnRef.BY_NUMBER: ConfigXlsListRefmtNum}


def config_factory_from_enum(numerator: ColumnRef) -> DerivedConfig:
    """Get correct configuration type for numerator value."""
    return _CONFIGS[numerator]()


def _config_factory_get_text(from_json_text: Optional[str] = None,
                             from_json_filename: Optional[str] = None) -> str:
    """Get JSON text to use."""
    if from_json_filename is None and from_json_text is None:
        msg = 'Either JSON text or JSON file needed. Both cannot be None.'
        print(msg, file=sys.stderr)
        raise RuntimeError(msg)
    if from_json_filename is not None and from_json_text is not None:
        msg = 'Either JSON text or JSON file needed. Both cannot be given.'
        print(msg, file=sys.stderr)
        raise RuntimeError(msg)
    if from_json_text is not None:
        return from_json_text
    assert from_json_filename is not None
    file_must_exist(filename=from_json_filename,
                    with_content_txt='configuration JSON input')
    with open(from_json_filename, "r", encoding='UTF-8') as file:
        text = file.read()
        return text


def _config_factory_exit(msg: str,
                         exc: Optional[JSONDecodeError] |
                         Optional[UnicodeDecodeError]) -> None:
    """Report config factory error and exit."""
    msg2 = '\nDid you specify an incorrect configuration file?\n'
    totmsg = msg + msg2 + ('' if exc is None else str(exc)) + '\n'
    print(totmsg, file=sys.stderr)
    sys.exit(1)


def config_factory_from_json(from_json_text: Optional[str] = None,
                             from_json_filename: Optional[str] = None) \
        -> DerivedConfig:
    """Get correct configuration type for JSON data."""
    text = _config_factory_get_text(from_json_text=from_json_text,
                                    from_json_filename=from_json_filename)
    data: JsonType = None
    try:
        data = json_loads(text)
    except JSONDecodeError as exc:
        msg = 'Configuration JSON cannot be decoded.'
        _config_factory_exit(msg, exc)
    except UnicodeDecodeError as exc:
        msg = 'Invalid UTF-8 in configuration data.\n'
        _config_factory_exit(msg, exc)
    if data is None or not isinstance(data, dict):
        msg = 'JSON data is not valid configuration. Top level not dict'
        _config_factory_exit(msg, None)
    assert data is not None
    assert isinstance(data, dict)
    if 'column_ref' not in data:
        msg = 'JSON data is not valid configuration. No key "column_ref".'
        _config_factory_exit(msg, None)
    refpar: JsonType = data['column_ref']
    assert isinstance(refpar, str)
    numerator = string_to_enum_best_match(inp=refpar,
                                          num_type=ColumnRef)
    return _CONFIGS[numerator](from_json_text=from_json_text,
                               from_json_filename=from_json_filename)
