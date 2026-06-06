#! /usr/local/bin/python3
"""Functions for transforming list data in table files."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from typing import Callable, Mapping
from config_as_json import Config, config_factory_from_json
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.file_extension import fix_file_extension
from excel_list_transform.config_match import MATCH_CONFIGS
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.migrate_cfg_warn_hook import EltMigrateCfgWarnHook
from excel_list_transform.transform_func_num import transform_named_files_num
from excel_list_transform.transform_func_named import \
    transform_named_files_name
from excel_list_transform.file_must_exist import file_must_exist


def transform_named_files(infilename: str, outfilename: str,
                          cfgfilename: str) -> int:
    """Transform list data from the named input file to output file."""
    fixed_cfgfilename = fix_file_extension(filename=cfgfilename,
                                           ext_to_add='.cfg')
    file_must_exist(fixed_cfgfilename)
    cfg = config_factory_from_json(match_configs=MATCH_CONFIGS,
                                   auto_ch_hook=EltMigrateCfgWarnHook(),
                                   from_json_filename=fixed_cfgfilename,
                                   stderr_file=sys.stderr)
    assert isinstance(cfg, (ConfigXlsListTransfName, ConfigXlsListTransfNum))
    dispatch: Mapping[ColumnRef,
                      Callable[[str, str, Config], None]] = \
        {ColumnRef.BY_NUMBER: transform_named_files_num,
         ColumnRef.BY_NAME: transform_named_files_name}
    dispatch[cfg.column_ref](infilename, outfilename, cfg)
    return 0
