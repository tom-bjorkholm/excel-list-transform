#! /usr/local/bin/python3
"""Read and write configuration of CSV splitting."""

# Copyright (c) 2024 Tom Björkholm
# MIT License


from typing import Optional
from excel_list_transform.config_enums import SplitWhere, ColumnRef
from excel_list_transform.config_excel_list_transform import \
    ConfigExcelListTransform, RuleOrder, ColInfo


class ConfigXlsListRefmtName(ConfigExcelListTransform[str]):  # pylint: disable=too-many-instance-attributes, line-too-long # noqa: E501
    """Class with configuration for excel list transform."""

    def __init__(self,
                 from_json_text: Optional[str] = None,
                 from_json_filename: Optional[str] = None) -> None:
        """Construct configuration for excel list transform."""
        col_to_use = ['street', 'street number', 'name', 'last name',
                      'Phone', 'Phone', 'Phone', 'Phone', 'Phone',
                      'Last Name']
        colinfo = ColInfo[str](split_last='right_name', insert_last=None,
                               s1=[{'column': 'name',
                                    'separator': ' ',
                                    'where': SplitWhere.RIGHTMOST,
                                    'right_name': 'last name'}],
                               s6=[{'column': 'Division', 'value': None},
                                   {'column': 'Other', 'value': 'some text'}],
                               col_to_use=col_to_use, tinfo='a')
        self.s8_column_order: RuleOrder = \
            ['Class', 'Division', 'Nationality', 'Sail Number', 'Boat Name',
             'First Name', 'Last Name', 'Club Name', 'Email', 'Phone',
             'WhatsApp']
        super().__init__(col_ref=ColumnRef.BY_NAME,
                         colinfo=colinfo, tinfo='a',
                         from_json_text=from_json_text,
                         from_json_filename=from_json_filename)
        self._duplicates_not_allowed(self.s8_column_order, 's8_column_order')
