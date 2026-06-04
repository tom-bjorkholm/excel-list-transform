#! /usr/local/bin/python3
"""Function for doing transform of lists excel files."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


from typing import Callable, Mapping
from config_as_json import Config
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.file_extension import fix_file_extension
from excel_list_transform.config_factory import config_factory_from_json
from excel_list_transform.transform_func_num import transform_named_files_num
from excel_list_transform.transform_func_named import \
    transform_named_files_name
from excel_list_transform.file_must_exist import file_must_exist


def transform_named_files(infilename: str, outfilename: str,
                          cfgfilename: str) -> int:
    """Transform list in the named excel file to named file."""
    fixed_cfgfilename = fix_file_extension(filename=cfgfilename,
                                           ext_to_add='.cfg')
    file_must_exist(fixed_cfgfilename)
    cfg = config_factory_from_json(from_json_filename=fixed_cfgfilename)
    dispatch: Mapping[ColumnRef,
                      Callable[[str, str, Config], None]] = \
        {ColumnRef.BY_NUMBER: transform_named_files_num,
         ColumnRef.BY_NAME: transform_named_files_name}
    dispatch[cfg.column_ref](infilename, outfilename, cfg)
    return 0
