#! /usr/local/bin/python3
"""Read and write configuration of excel list transform."""

# Copyright (c) 2024 Tom Björkholm
# MIT License


import sys
from copy import deepcopy
from enum import Enum
from collections import Counter
from typing import Optional, Callable, TypeAlias, TypeVar, NamedTuple, Generic
from csv import Dialect
from excel_list_transform.config import Config, ParseConverter
from excel_list_transform.config_enums import FileType, SplitWhere, \
    ExcelLib, RewriteKind, CaseSensitivity, ColumnRef
from excel_list_transform.str_to_enum import string_to_enum_best_match


CsvSpec: TypeAlias = dict[str, Optional[str]]
Column = TypeVar('Column', int, str)
SingleRule: TypeAlias = dict[str, Optional[Column | str]]
Rule: TypeAlias = list[SingleRule[Column]]
SingleRuleSplit: TypeAlias = dict[str, Optional[Column | str | SplitWhere]]
RuleSplit: TypeAlias = list[SingleRuleSplit[Column]]
RuleOrder: TypeAlias = list[str]
RulePlace: TypeAlias = list[int]
RuleRemove: TypeAlias = RulePlace
RulePlaceOrOrder: TypeAlias = RulePlace | RuleOrder
SingleRuleRewrite: TypeAlias = dict[str,
                                    Optional[Column | str |
                                             list[str] |
                                             RewriteKind |
                                             CaseSensitivity]]
RuleRewrite: TypeAlias = list[SingleRuleRewrite[Column]]
SingleRuleMerge: TypeAlias = dict[str, list[Column] | str]
RuleMerge: TypeAlias = list[SingleRuleMerge[Column]]
DupCheckKeyArg: TypeAlias = \
    SingleRule[Column] | SingleRuleMerge[Column] | SingleRuleSplit[Column]
DupCheckKey: TypeAlias = \
    Callable[[DupCheckKeyArg[Column]], Column | list[Column]]
IncrCheckKey: TypeAlias = \
    Callable[[SingleRuleMerge[Column]], Column | list[Column]]
NoDupKeydType: TypeAlias = \
    Rule[Column] | RuleMerge[Column] | RuleSplit[Column]


class ColInfo(NamedTuple, Generic[Column]):
    """Information about columns to pass to ConfigExcelListTransform init."""

    split_last: str
    insert_last: Optional[str]
    s1: RuleSplit[Column]
    s6: Rule[Column]
    col_to_use: list[Column]
    tinfo: Column


class ConfigExcelListTransform(Config, Generic[Column]):  # pylint: disable=too-many-instance-attributes, line-too-long # noqa: E501
    """Class with configuration for excel list transform."""

    def __init__(self, col_ref: ColumnRef,  # pylint: disable=too-many-arguments # noqa: E501
                 colinfo: ColInfo[Column], tinfo: Column,
                 from_json_text: Optional[str] = None,
                 from_json_filename: Optional[str] = None) -> None:
        """Construct configuration for excel list transform."""
        assert isinstance(colinfo.tinfo, type(tinfo))
        self.column_ref: ColumnRef = col_ref
        col2use = deepcopy(colinfo.col_to_use)  # dont destroying caller's arg
        self.max_column_read: int = 20
        self.out_csv_dialect: CsvSpec = {'name': 'csv.excel',
                                         'delimiter': ',', 'quoting': None,
                                         'quotechar': '"',
                                         'lineterminator': None,
                                         'escapechar': None}
        self.in_csv_dialect: CsvSpec = {'name': 'csv.excel',
                                        'delimiter': ',', 'quoting': None,
                                        'quotechar': '"',
                                        'lineterminator': None,
                                        'escapechar': None}
        self.in_type: FileType = FileType.EXCEL
        self.in_excel_library: ExcelLib = ExcelLib.PYLIGHTXL
        self.out_excel_library: ExcelLib = ExcelLib.PYLIGHTXL
        self.out_type: FileType = FileType.EXCEL
        self.s1_split_columns: RuleSplit[Column] = colinfo.s1
        self.s3_merge_columns: RuleMerge[Column] = \
            [{'columns': [col2use.pop(0), col2use.pop(0)],
              'separator': ' '}]
        self.s5_rename_columns: Rule[Column] = \
            [{'column': col2use.pop(0), 'name': 'First Name'},
             {'column': col2use.pop(0), 'name': 'Last Name'}]
        self.s6_insert_columns: Rule[Column] = colinfo.s6
        self.s7_rewrite_columns: RuleRewrite[Column] = \
            [{'column': col2use.pop(0),
              'kind': RewriteKind.STRIP, 'chars': '',
              'case': CaseSensitivity.IGNORE_CASE},
             {'column': col2use.pop(0), 'kind': RewriteKind.REMOVECHARS,
              'chars': [' ', '-'], 'case': CaseSensitivity.MATCH_CASE},
             {'column': col2use.pop(0),
              'kind': RewriteKind.REGEX_SUBSTITUTE,
              'from': '^07', 'to': '+467',
              'case': CaseSensitivity.MATCH_CASE},
             {'column': col2use.pop(0),
              'kind': RewriteKind.REGEX_SUBSTITUTE,
              'from': '^+4607', 'to': '+467',
              'case': CaseSensitivity.MATCH_CASE},
             {'column': col2use.pop(0),
              'kind': RewriteKind.REGEX_SUBSTITUTE,
              'from': '^467', 'to': '+467',
              'case': CaseSensitivity.MATCH_CASE},
             {'column': col2use.pop(0), 'kind': RewriteKind.STR_SUBSTITUTE,
              'from': 'donald', 'to': 'duck',
              'case': CaseSensitivity.IGNORE_CASE}]
        super().__init__(from_json_text, from_json_filename)
        self.check_array_configs(split_last=colinfo.split_last,
                                 insert_last=colinfo.insert_last)
        self.sort_sx_hook()
        self._check_no_duplicate_single(self.s1_split_columns,
                                        's1_split_columns', colinfo.tinfo)
        self._check_no_duplicate_single(self.s5_rename_columns,
                                        's5_rename_columns', colinfo.tinfo)
        self._check_no_duplicate_single(self.s6_insert_columns,
                                        's6_insert_columns', colinfo.tinfo)
        self.check_rewrite_configs(coltype=type(colinfo.tinfo))

    def get_out_csv_dialect(self) -> type[Dialect]:
        """Get CSV dialect for output file."""
        assert self.out_csv_dialect['name'] is not None
        return self.get_csv_dialect(**self.out_csv_dialect)

    def get_in_csv_dialect(self) -> type[Dialect]:
        """Get CSV dialect for input file from Portfolio."""
        assert self.in_csv_dialect['name'] is not None
        return self.get_csv_dialect(**self.in_csv_dialect)

    def sort_sx_hook(self) -> None:
        """Sort s[0-9]_ as needed as needed (hook)."""

    @staticmethod
    def _duplicates_not_allowed(expanded_data: list[Column],
                                param_name: str) -> None:
        """Error report duplicate data."""
        dup = [str(k) for k, v in Counter(expanded_data).items() if v > 1]
        if len(dup) == 0:
            return
        msg = f'Duplicates not allowed in {param_name}. Duplicate values: '  # noqa: E713, E501
        msg += ','.join(dup)
        print(msg, file=sys.stderr)
        raise KeyError(msg)

    @staticmethod
    def get_cols_single(rule: Rule[Column] | RuleSplit[Column] |
                        RuleRewrite[Column],
                        tinfo: Column) -> list[Column]:
        """Get list of columns for modification rule."""
        cols: list[Column] = []
        for row in rule:
            rowcol = row['column']
            assert isinstance(rowcol, type(tinfo))
            cols.append(rowcol)
        return cols

    @staticmethod
    def get_cols_multi(rule: RuleMerge[Column],
                       tinfo: Column) -> list[Column]:
        """Get list of columns for merge rule."""
        cols: list[Column] = []
        for row in rule:
            rowcols = row['columns']
            for singlecol in rowcols:
                assert isinstance(singlecol, type(tinfo))
                cols.append(singlecol)
        return cols

    @staticmethod
    def _check_no_duplicate_single(rule: Rule[Column] | RuleSplit[Column],
                                   param_name: str, tinfo: Column) -> None:
        """Flag as error if column is refered to multiple times."""
        cols: list[Column] = ConfigExcelListTransform.get_cols_single(rule,
                                                                      tinfo)
        ConfigExcelListTransform._duplicates_not_allowed(cols, param_name)

    @staticmethod
    def _check_no_duplicate_multi(rule: RuleMerge[Column],
                                  param_name: str, tinfo: Column) -> None:
        """Flag as error if column is refered to multiple times."""
        cols: list[Column] = ConfigExcelListTransform.get_cols_multi(rule,
                                                                     tinfo)
        ConfigExcelListTransform._duplicates_not_allowed(cols, param_name)

    @staticmethod
    def _check_increasing_multi(rule: RuleMerge[Column], param_name: str,
                                tinfo: Column) -> None:
        """Flag as error if order is not increasing."""
        cols: list[Column] = ConfigExcelListTransform.get_cols_multi(rule,
                                                                     tinfo)
        seen: Optional[Column] = None
        for col in cols:
            assert isinstance(col, (int, str))
            if seen is not None and seen >= col:
                msg: str = f'Increasing order needed in {param_name}'
                print(msg, file=sys.stderr)
                raise KeyError(msg)
            seen = col

    @staticmethod
    def get_converter_dict(enum_type: type[Enum]) -> ParseConverter:
        """Get dict for converting to given enum_type."""
        return ParseConverter(result_type=enum_type,
                              func=string_to_enum_best_match,
                              args={'num_type': enum_type})

    def parse_converters(self) -> dict[str, ParseConverter]:
        """Get converters for use when parsing JSON.

        Overriding in derived class.
        Return None if no conversions.
        Return dict of dict for use in json decoder hook.
        Structure of return value shall be:
        {key: {'result type': res_type, 'func': function,
        'args': {arg_name: arg_value}}}.
        """
        return {'in_type': self.get_converter_dict(FileType),
                'out_type': self.get_converter_dict(FileType),
                'in_excel_library': self.get_converter_dict(ExcelLib),
                'out_excel_library': self.get_converter_dict(ExcelLib),
                'where': self.get_converter_dict(SplitWhere),
                'store_single': self.get_converter_dict(SplitWhere),
                'kind': self.get_converter_dict(RewriteKind),
                'case': self.get_converter_dict(CaseSensitivity),
                'column_ref': self.get_converter_dict(ColumnRef)}

    def check_array_configs(self, split_last: str,
                            insert_last: Optional[str]) -> None:
        """Check that keywords in configuration arrays are OK."""
        split_col_keys = ['column', 'separator', 'where', split_last]
        self.check_array_keys('s1_split_columns', self.s1_split_columns,
                              split_col_keys)
        merge_col_keys = ['columns', 'separator']
        self.check_array_keys('s3_merge_columns', self.s3_merge_columns,
                              merge_col_keys)
        rename_col_keys = ['column', 'name']
        self.check_array_keys('s5_rename_columns', self.s5_rename_columns,
                              rename_col_keys)
        insert_col_keys = ['column', 'value']
        if insert_last is not None:
            assert insert_last is not None  # keep mypy happy
            insert_col_keys.append(insert_last)
        self.check_array_keys('s6_insert_columns', self.s6_insert_columns,
                              insert_col_keys)

    def check_rewrite_configs(self, coltype: type) -> None:
        """Check the rewrite column configuration."""
        rewrite_col_mand_keys: list[str] = ['column', 'kind', 'case']
        rewrite_col_opt_keys: list[str] = ['chars', 'from', 'to']
        self.check_array_keys('s7_rewrite_columns', self.s7_rewrite_columns,
                              mandatory_keys=rewrite_col_mand_keys,
                              allowed_keys=rewrite_col_opt_keys)
        template = {RewriteKind.STRIP: {'column': coltype,
                                        'kind': RewriteKind,
                                        'chars': str,
                                        'case': CaseSensitivity},
                    RewriteKind.REMOVECHARS: {'column': coltype,
                                              'kind': RewriteKind,
                                              'chars': list,
                                              'case': CaseSensitivity},
                    RewriteKind.REGEX_SUBSTITUTE: {'column': coltype,
                                                   'kind': RewriteKind,
                                                   'from': str, 'to': str,
                                                   'case': CaseSensitivity},
                    RewriteKind.STR_SUBSTITUTE: {'column': coltype,
                                                 'kind': RewriteKind,
                                                 'from': str, 'to': str,
                                                 'case': CaseSensitivity}}
        self.check_array_dicts(name_of_cfg='s7_rewrite_columns',
                               array=self.s7_rewrite_columns,
                               kind_key='kind', kind_type=RewriteKind,
                               dict_of_templates=template)
