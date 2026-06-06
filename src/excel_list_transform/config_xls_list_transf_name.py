#! /usr/local/bin/python3
"""Read and write configuration with named column references."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


import sys
from typing import Optional, TextIO, override
from config_as_json import ConfigAutoChangeHook, MemberValidationStep, \
    PathOrStr
from excel_list_transform.config_enums import SplitWhere, ColumnRef
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform, RuleOrder, ColInfo


# pylint: disable-next=too-many-instance-attributes,duplicate-code
class ConfigXlsListTransfName(ConfigExcelListTransform[str]):
    """Class with configuration for excel list transform."""

    def __init__(self, from_json_data_text: Optional[str] = None,
                 from_json_filename: Optional[PathOrStr] = None,
                 auto_ch_hook: Optional[ConfigAutoChangeHook] = None,
                 stderr_file: TextIO = sys.stderr) -> None:
        """Construct configuration for excel list transform."""
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
            ['Class', 'Division', 'Nationality', 'MNA No.', 'Sail Number',
             'Boat Name', 'First Name', 'Last Name', 'Club Name', 'Email',
             'Phone', 'Whats App Number']
        super().__init__(col_ref=ColumnRef.BY_NAME, colinfo=colinfo, tinfo='a',
                         from_json_data_text=from_json_data_text,
                         from_json_filename=from_json_filename,
                         auto_ch_hook=auto_ch_hook, stderr_file=stderr_file)

    @override
    def get_column_val_steps(self) -> list[MemberValidationStep]:
        """Return validation steps for name-based rule members."""
        return [self._list_unique_step('s10_column_order', str)]
