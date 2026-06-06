#! /usr/local/bin/python3
"""Read and write configuration with numbered column references."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


import sys
from typing import Optional, TextIO, override
from config_as_json import ConfigAutoChangeHook, MemberValidationStep, \
    PathOrStr
from excel_list_transform.config_enums import SplitWhere, ColumnRef
from excel_list_transform.config_excel_list_transform \
    import ConfigExcelListTransform, RulePlace, RuleRemove, \
    SingleRuleMerge, SingleRuleSplit, SingleRule, ColInfo


def get_column(rule:  SingleRuleSplit[int] | SingleRule[int]) -> int:
    """Get column of rule."""
    ret = rule['column']
    assert isinstance(ret, int)
    return ret


def get_merge_first_column(rule: SingleRuleMerge[int]) -> int:
    """Get column of rule."""
    cols = rule['columns']
    ret = cols[0]
    assert isinstance(ret, int)
    return ret


# pylint: disable-next=too-many-instance-attributes,duplicate-code
class ConfigXlsListTransfNum(ConfigExcelListTransform[int]):
    """Class with configuration for excel list transform."""

    def __init__(self, from_json_data_text: Optional[str] = None,
                 from_json_filename: Optional[PathOrStr] = None,
                 auto_ch_hook: Optional[ConfigAutoChangeHook] = None,
                 stderr_file: TextIO = sys.stderr) -> None:
        """Construct configuration for excel list transform."""
        self.s04_remove_columns: RuleRemove = [1, 2, 3]
        self.s06_place_columns_first: RulePlace = [7, 3, 6]
        col_to_use = [15, 16, 1, 2, 5, 5, 5, 5, 5, 6]
        col_to_use_row = [7, 1, 2]
        colinfo: ColInfo[int] = \
            ColInfo[int]('store_single', 'name',
                         [{'column': 15, 'separator': ' ',
                           'where': SplitWhere.RIGHTMOST,
                           'store_single': SplitWhere.LEFTMOST}],
                         [{'column': 1, 'name': 'Division', 'value': None},
                          {'column': 7, 'name': 'Other',
                           'value': 'some text'}],
                         col_to_use, col_to_use_row, 2)
        super().__init__(col_ref=ColumnRef.BY_NUMBER, colinfo=colinfo, tinfo=2,
                         from_json_data_text=from_json_data_text,
                         from_json_filename=from_json_filename,
                         auto_ch_hook=auto_ch_hook, stderr_file=stderr_file)

    @override
    def column_shape_steps(self) -> list[MemberValidationStep]:
        """Return pre-normalization validation for number-based members."""
        return [
            self._plain_column_list_step('s04_remove_columns'),
            self._plain_column_list_step('s06_place_columns_first')]

    @override
    def get_column_val_steps(self) -> list[MemberValidationStep]:
        """Return validation steps for number-based rule members."""
        return [
            self._list_unique_step('s04_remove_columns', int),
            self._merge_order_step('s05_merge_columns'),
            self._list_unique_step('s06_place_columns_first', int)]

    def sort_sx_hook(self) -> None:
        """Sort s[0-9]_ as needed as needed (hook)."""
        self.s03_split_columns = sorted(self.s03_split_columns, key=get_column)
        self.s04_remove_columns = sorted(self.s04_remove_columns)
        self.s05_merge_columns = sorted(self.s05_merge_columns,
                                        key=get_merge_first_column)
        self.s07_rename_columns = sorted(self.s07_rename_columns,
                                         key=get_column)
        self.s08_insert_columns = sorted(self.s08_insert_columns,
                                         key=get_column)
