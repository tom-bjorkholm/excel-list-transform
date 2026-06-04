#! /usr/local/bin/python3
"""Read and write configuration of CSV splitting."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from typing import Optional, TextIO
from config_as_json import ConfigAutoChangeHook, PathOrStr
from excel_list_transform.config_enums import SplitWhere, ColumnRef
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform, RuleOrder, ColInfo
from excel_list_transform.migrate_cfg_warn_hook import MigrateCfgWarnHook


class ConfigXlsListTransfName(ConfigExcelListTransform[str]):  # pylint: disable=too-many-instance-attributes, line-too-long # noqa: E501
    """Class with configuration for excel list transform."""

    def __init__(self, from_json_text: Optional[str] = None,
                 from_json_filename: Optional[PathOrStr] = None,
                 auto_ch_hook: ConfigAutoChangeHook =
                 MigrateCfgWarnHook(),
                 stderr_file: Optional[TextIO] = None) -> None:
        """Construct configuration for excel list transform."""
        err_file = sys.stderr if stderr_file is None else stderr_file
        col_to_use = ['street', 'street number', 'name', 'last name',
                      'Phone', 'Phone', 'Phone', 'Phone', 'Phone',
                      'Last Name']
        col_to_use_row = ['Club Name', 'name', 'last name']
        colinfo = ColInfo[str](split_last='right_name', insert_last=None,
                               s03=[{'column': 'name',
                                     'separator': ' ',
                                     'where': SplitWhere.RIGHTMOST,
                                     'right_name': 'last name'}],
                               s08=[{'column': 'Division', 'value': None},
                                    {'column': 'Other', 'value': 'some text'}],
                               col_to_use=col_to_use,
                               col_to_use_row=col_to_use_row, tinfo='a')
        self.s10_column_order: RuleOrder = \
            ['Class', 'Division', 'Nationality', 'Sail Number', 'Boat Name',
             'First Name', 'Last Name', 'Club Name', 'Email', 'Phone',
             'WhatsApp']
        super().__init__(col_ref=ColumnRef.BY_NAME, colinfo=colinfo, tinfo='a',
                         from_json_text=from_json_text,
                         from_json_filename=from_json_filename,
                         auto_ch_hook=auto_ch_hook, stderr_file=err_file)
        self.check_no_duplicates(self.s10_column_order, 's10_column_order')
