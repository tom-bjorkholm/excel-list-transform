#! /usr/local/bin/python3
"""Read and write configuration of excel list transform."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from copy import deepcopy
from enum import Enum
from typing import Optional, TypeVar, NamedTuple, Generic, Mapping, \
    NoReturn, Sequence, TypedDict, TextIO, override
from config_as_json import Config, ConfigAutoChangeHook, ConfigNesting, \
    ConfigNestingKind, InvalidConfiguration, ListIsOrderedValidator, \
    NestedConfigs, ParseConverter, PathOrStr, ProjectedMemberValidator, \
    ReadOldConfiguration, ValidationPlan, MemberValidationStep, \
    WholeConfigValidationStep, ValueTypeValidator, CallingWholeConfigValidator
from tableio import CAP_IGNORABLE, CAP_NEEDED, Capabilities, FileAccess, \
    TableBorderStyle
from tableio_cfg_json import TioJsonConfig
from excel_list_transform.config_enums import SplitWhere, RewriteKind, \
    CaseSensitivity, ColumnRef
from excel_list_transform.config_read_old import ConfigReadOld
from excel_list_transform.str_to_enum import string_to_enum_best_match
from excel_list_transform.migrate_cfg_warn_hook import EltMigrateCfgWarnHook


Column = TypeVar('Column', int, str)
type SingleRule[Column] = dict[str, Optional[Column | str]]
type Rule[Column] = list[SingleRule[Column]]
type SingleRuleSplit[Column] = dict[str, Optional[Column | str | SplitWhere]]


class SingleRuleRowSplit[Column](TypedDict):
    """Sinle rule for splitting a column."""

    column: Column
    separators: list[str]
    not_separators: list[str]


type RuleSplit[Column] = list[SingleRuleSplit[Column]]
type RuleRowSplit[Column] = list[SingleRuleRowSplit[Column]]
type RuleOrder = list[str]
type RulePlace = list[int]
type RuleRemove = RulePlace
type RulePlaceOrOrder = RulePlace | RuleOrder
type SingleRuleRewrite[Column] = dict[str,
                                      Optional[Column | str |
                                               list[str] |
                                               RewriteKind |
                                               CaseSensitivity]]
type RuleRewrite[Column] = list[SingleRuleRewrite[Column]]
type SingleRuleMerge[Column] = dict[str, list[Column] | str]
type RuleMerge[Column] = list[SingleRuleMerge[Column]]


def input_capabilities() -> Capabilities:
    """Return TableIO capabilities used for input files."""
    return Capabilities(can_read=CAP_NEEDED)


def output_capabilities() -> Capabilities:
    """Return TableIO capabilities used for output files."""
    return Capabilities(can_write=CAP_NEEDED,
                        filtered_data_range=CAP_IGNORABLE,
                        can_write_borders=CAP_IGNORABLE)


def input_table_factory(from_json_data_text: Optional[str] = None,
                        from_json_filename: Optional[PathOrStr] = None,
                        auto_ch_hook: Optional[ConfigAutoChangeHook] = None,
                        stderr_file: TextIO = sys.stderr) -> TioJsonConfig:
    """Create JSON-backed TableIO configuration for input files."""
    return TioJsonConfig(capabilities=input_capabilities(),
                         file_access=FileAccess.READ,
                         from_json_data_text=from_json_data_text,
                         from_json_filename=from_json_filename,
                         auto_ch_hook=auto_ch_hook, stderr_file=stderr_file)


def output_table_factory(from_json_data_text: Optional[str] = None,
                         from_json_filename: Optional[PathOrStr] = None,
                         auto_ch_hook: Optional[ConfigAutoChangeHook] = None,
                         stderr_file: TextIO = sys.stderr) -> TioJsonConfig:
    """Create JSON-backed TableIO configuration for output files."""
    return TioJsonConfig(capabilities=output_capabilities(),
                         file_access=FileAccess.CREATE,
                         from_json_data_text=from_json_data_text,
                         from_json_filename=from_json_filename,
                         auto_ch_hook=auto_ch_hook, stderr_file=stderr_file)


def _invalid_projection(member_name: str, message: str,
                        stderr_file: TextIO) -> NoReturn:
    """Raise a readable validation error for projected rule validation."""
    msg = f'Invalid configuration: {member_name} {message}'
    print(msg, file=stderr_file)
    raise InvalidConfiguration(msg)


def _project_rule_columns(config: Config, member_name: str,
                          member_value: object, stderr_file: TextIO) -> object:
    """Project one-column rule dictionaries to a column list."""
    _ = config
    if not isinstance(member_value, list):
        _invalid_projection(member_name, 'must be a list.', stderr_file)
    values: list[object] = []
    for index, item in enumerate(member_value):
        if not isinstance(item, dict):
            msg = f'contains a non-object rule at index {index}.'
            _invalid_projection(member_name, msg, stderr_file)
        if 'column' not in item:
            msg = f'misses key column at index {index}.'
            _invalid_projection(member_name, msg, stderr_file)
        values.append(item['column'])
    return values


def _project_merge_columns(config: Config, member_name: str,
                           member_value: object,
                           stderr_file: TextIO) -> object:
    """Project merge-rule dictionaries to one flattened column list."""
    _ = config
    if not isinstance(member_value, list):
        _invalid_projection(member_name, 'must be a list.', stderr_file)
    values: list[object] = []
    for index, item in enumerate(member_value):
        if not isinstance(item, dict):
            msg = f'contains a non-object rule at index {index}.'
            _invalid_projection(member_name, msg, stderr_file)
        if 'columns' not in item:
            msg = f'misses key columns at index {index}.'
            _invalid_projection(member_name, msg, stderr_file)
        columns = item['columns']
        if not isinstance(columns, list):
            msg = f'has non-list columns at index {index}.'
            _invalid_projection(member_name, msg, stderr_file)
        values.extend(columns)
    return values


class ColInfo(NamedTuple, Generic[Column]):
    """Information about columns to pass to ConfigExcelListTransform init."""

    split_last: str
    insert_last: Optional[str]
    s03: RuleSplit[Column]
    s08: Rule[Column]
    col_to_use: list[Column]
    col_to_use_row: list[Column]
    tinfo: Column


# pylint: disable-next=too-many-instance-attributes,too-many-public-methods
class ConfigExcelListTransform(Config, Generic[Column]):
    """Class with configuration for excel list transform."""

    # pylint: disable-next=too-many-arguments,too-many-positional-arguments
    def __init__(self, *, col_ref: ColumnRef, colinfo: ColInfo[Column],
                 tinfo: Column, from_json_data_text: Optional[str] = None,
                 from_json_filename: Optional[PathOrStr] = None,
                 auto_ch_hook: Optional[ConfigAutoChangeHook] = None,
                 stderr_file: TextIO = sys.stderr) -> None:
        """Construct configuration for excel list transform."""
        assert isinstance(colinfo.tinfo, type(tinfo))
        hook = EltMigrateCfgWarnHook() if auto_ch_hook is None \
            else auto_ch_hook
        self._colinfo: ColInfo[Column] = deepcopy(colinfo)
        self._columntype: type[Column] = type(tinfo)
        self.column_ref: ColumnRef = col_ref
        col2use = deepcopy(colinfo.col_to_use)  # dont destroying caller's arg
        col2userow = deepcopy(colinfo.col_to_use_row)
        self.max_column_read: int = 20
        self.input_table: TioJsonConfig = input_table_factory()
        self.input_table.character_encoding = 'utf_8_sig'
        self.output_table: TioJsonConfig = output_table_factory()
        self.output_table.character_encoding = 'utf-8'
        self.in_excel_col_name_strip = True
        self.in_excel_values_strip = False
        self.output_borders: bool = False
        self.output_filtered_table: bool = False
        self.s01_split_rows: RuleRowSplit[Column] = \
            [{'column': col2userow.pop(0),
              'separators': [' ', '+'],
              'not_separators': ['\\ ', '\\+', ' + ']}]
        self.s02_merge_rows: RuleMerge[Column] = \
            [{'columns': [col2userow.pop(0), col2userow.pop(0)],
              'separator': ' + '}]
        self.s03_split_columns: RuleSplit[Column] = colinfo.s03
        self.s05_merge_columns: RuleMerge[Column] = \
            [{'columns': [col2use.pop(0), col2use.pop(0)],
              'separator': ' '}]
        self.s07_rename_columns: Rule[Column] = \
            [{'column': col2use.pop(0), 'name': 'First Name'},
             {'column': col2use.pop(0), 'name': 'Last Name'}]
        self.s08_insert_columns: Rule[Column] = colinfo.s08
        self.s09_rewrite_columns: RuleRewrite[Column] = \
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
        super().__init__(from_json_data_text=from_json_data_text,
                         from_json_filename=from_json_filename,
                         auto_ch_hook=hook, stderr_file=stderr_file)

    @override
    def nested_configs(self) -> NestedConfigs:
        """Return nested TableIO config sections."""
        input_nesting = ConfigNesting(kind=ConfigNestingKind.MEMBER,
                                      config_type=TioJsonConfig,
                                      factory_function=input_table_factory)
        output_nesting = ConfigNesting(kind=ConfigNestingKind.MEMBER,
                                       config_type=TioJsonConfig,
                                       factory_function=output_table_factory)
        return {'input_table': input_nesting, 'output_table': output_nesting}

    def _read_old_config(self) -> ReadOldConfiguration:
        """Return old-format migration rules."""
        return ConfigReadOld()

    _get_read_old_configuration = _read_old_config

    @override
    def get_validation_plan(self, stderr_file: TextIO) -> ValidationPlan:
        """Return validation plan for config-as-json."""
        _ = stderr_file
        bool_names = ['in_excel_col_name_strip', 'in_excel_values_strip',
                      'output_borders', 'output_filtered_table']
        plan = [
            MemberValidationStep(member_names=['max_column_read'],
                                 validator=ValueTypeValidator(
                                     int, not_allowed_type=bool)),
            MemberValidationStep(member_names=bool_names,
                                 validator=ValueTypeValidator(bool)),
            WholeConfigValidationStep(validator=CallingWholeConfigValidator(
                method_name='validate_transform_rules'))]
        plan.extend(self._base_rule_val_steps())
        plan.extend(self.get_column_val_steps())
        return plan

    def _base_rule_val_steps(self) -> list[MemberValidationStep]:
        """Return validation steps for shared transform-rule members."""
        return [
            self._rule_unique_step('s03_split_columns'),
            self._rule_unique_step('s07_rename_columns'),
            self._rule_unique_step('s08_insert_columns')]

    def get_column_val_steps(self) -> list[MemberValidationStep]:
        """Return validation steps for column-reference-specific members."""
        return []

    def _rule_unique_step(self, member_name: str) -> MemberValidationStep:
        """Return uniqueness validation for a one-column rule member."""
        validator = ProjectedMemberValidator(
            projector=_project_rule_columns,
            validators=[
                ListIsOrderedValidator(self._columntype, is_ordered=False,
                                       unique_values=True)])
        return MemberValidationStep(member_names=[member_name],
                                    validator=validator)

    def _merge_order_step(self, member_name: str) -> MemberValidationStep:
        """Return increasing and unique validation for merge columns."""
        validator = ProjectedMemberValidator(
            projector=_project_merge_columns,
            validators=[
                ListIsOrderedValidator(self._columntype, is_ordered=True,
                                       unique_values=True)])
        return MemberValidationStep(member_names=[member_name],
                                    validator=validator)

    @staticmethod
    def _list_unique_step(member_name: str,
                          element_type: type[Column]) -> MemberValidationStep:
        """Return uniqueness validation for a plain list member."""
        validator = ListIsOrderedValidator(element_type, is_ordered=False,
                                           unique_values=True)
        return MemberValidationStep(member_names=[member_name],
                                    validator=validator)

    def validate_transform_rules(self) -> None:
        """Validate transform-rule settings."""
        self.validate_base_rules()
        self.validate_column_rules()

    def validate_base_rules(self) -> None:
        """Validate common transform-rule settings."""
        colinfo = self._colinfo
        self.check_array_configs(split_last=colinfo.split_last,
                                 insert_last=colinfo.insert_last)
        self.sort_sx_hook()
        self.check_rewrite_configs(coltype=type(colinfo.tinfo))
        self.check_split_row_cfg()
        self.check_merge_row_cfg()

    def validate_column_rules(self) -> None:
        """Validate column-reference-specific transform-rule settings."""

    def table_border_style(self) -> TableBorderStyle:
        """Return the TableIO border style requested by the config."""
        if self.output_borders:
            return TableBorderStyle.OUTER_FIRST_ROW_THICK_INNER_THIN
        return TableBorderStyle.NONE

    def sort_sx_hook(self) -> None:
        """Sort s[0-9]_ as needed (hook)."""

    @staticmethod
    def check_array_keys(name_of_cfg: str,
                         array: Sequence[Mapping[str, object]],
                         mandatory_keys: list[str],
                         allowed_keys: Optional[list[str]] = None) -> None:
        """Check allowed and mandatory keys in a list of dictionaries."""
        to_allow = deepcopy(mandatory_keys)
        if allowed_keys is not None:
            to_allow += deepcopy(allowed_keys)
        for row in array:
            for used_key in row:
                if used_key not in to_allow:
                    msg = f'Found non-allowed key "{used_key}"'
                    print(msg + f' in config of {name_of_cfg}',
                          file=sys.stderr)
                    sys.exit(1)
            for key in mandatory_keys:
                if key not in row:
                    msg = f'Missing key "{key}"'
                    print(msg + f' in config of {name_of_cfg}',
                          file=sys.stderr)
                    sys.exit(1)

    @staticmethod
    # pylint: disable-next=too-many-arguments,too-many-positional-arguments
    def check_lst_dict(paramname: str, inp: Sequence[Mapping[str, object]],
                       key: str, key_optional: bool, valtype: type,
                       min_size_list: int) -> None:
        """Check a list of dictionaries with one typed value per row."""
        errtxt = f'Error in parameter {paramname}. '
        if len(inp) < min_size_list:
            msg = f'Minimum {min_size_list} elements needed in list but '
            msg += f'only {len(inp)} found.'
            print(errtxt + msg, file=sys.stderr)
            sys.exit(1)
        for elem in inp:
            if key not in elem:
                if key_optional:
                    continue
                print(errtxt + f'Expected key {key} not in dict in list',
                      file=sys.stderr)
                sys.exit(1)
            val = elem[key]
            if not isinstance(val, valtype):
                msg = f'Value for key {key} expected to be of type '
                msg += f'{valtype.__name__} but is of type '
                msg += type(val).__name__
                print(errtxt + msg, file=sys.stderr)
                sys.exit(1)

    @staticmethod
    # pylint: disable-next=too-many-arguments,too-many-positional-arguments
    def check_lst_dict_lst(paramname: str, inp: Sequence[Mapping[str, object]],
                           key: str, key_optional: bool, valtype: type,
                           min_size_outer_list: int,
                           min_size_inner_list: int) -> None:
        """Check a list of dictionaries with one typed list per row."""
        ConfigExcelListTransform.check_lst_dict(
            paramname=paramname, inp=inp, key=key, key_optional=key_optional,
            valtype=list, min_size_list=min_size_outer_list)
        errtxt = f'Error in parameter {paramname}. '
        for elem in inp:
            if key not in elem and key_optional:
                continue
            val = elem[key]
            assert isinstance(val, list)
            if len(val) < min_size_inner_list:
                msg = f'List for key {key} shall be minimum '
                msg += f'{min_size_inner_list} elements. '
                msg += f'But it is {len(val)} elements only.'
                print(errtxt + msg, file=sys.stderr)
                sys.exit(1)
            for item in val:
                if not isinstance(item, valtype):
                    msg = f'Value for key {key} expected to be list of '
                    msg += f'{valtype.__name__}. But element in list is '
                    msg += type(item).__name__
                    print(errtxt + msg, file=sys.stderr)
                    sys.exit(1)

    @staticmethod
    def check_array_dicts(name_of_cfg: str, array: Sequence[Mapping[str,
                                                                    object]],
                          kind_key: str, kind_type: type[RewriteKind],
                          dict_of_templates: Mapping[RewriteKind,
                                                     Mapping[str, type]]
                          ) -> None:
        """Check list-of-dicts entries selected by an enum kind key."""
        for index, row in enumerate(array):
            litem = f'(list index {index})'
            if kind_key not in row:
                msg = f'Key {kind_key} not in dict in config of '
                print(msg + name_of_cfg + ' ' + litem, file=sys.stderr)
                sys.exit(1)
            kind = row[kind_key]
            assert isinstance(kind, kind_type)
            template = dict_of_templates[kind]
            for key, valtype in template.items():
                if key not in row:
                    msg = f'Key {key} not in dict in config of '
                    print(msg + name_of_cfg + ' ' + litem, file=sys.stderr)
                    sys.exit(1)
                if not isinstance(row[key], valtype):
                    msg = f'Value for key {key} = {row[key]} '
                    msg += f'is not {valtype.__name__}; it is '
                    msg += type(row[key]).__name__
                    print(msg + f' in config of {name_of_cfg} ' + litem,
                          file=sys.stderr)
                    sys.exit(1)

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
    def get_cols_multi(rule: RuleMerge[Column], tinfo: Column) -> list[Column]:
        """Get list of columns for merge rule."""
        cols: list[Column] = []
        for row in rule:
            rowcols = row['columns']
            for singlecol in rowcols:
                assert isinstance(singlecol, type(tinfo))
                cols.append(singlecol)
        return cols

    @staticmethod
    def get_converter_dict(enum_type: type[Enum]) -> ParseConverter:
        """Get dict for converting to given enum_type."""
        converter = string_to_enum_best_match
        return ParseConverter(result_type=enum_type, func=converter,
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
        return {'where': self.get_converter_dict(SplitWhere),
                'store_single': self.get_converter_dict(SplitWhere),
                'kind': self.get_converter_dict(RewriteKind),
                'case': self.get_converter_dict(CaseSensitivity),
                'column_ref': self.get_converter_dict(ColumnRef)}

    def check_array_configs(self, split_last: str,
                            insert_last: Optional[str]) -> None:
        """Check that keywords in configuration arrays are OK."""
        split_col_keys = ['column', 'separator', 'where', split_last]
        self.check_array_keys('s03_split_columns', self.s03_split_columns,
                              split_col_keys)
        merge_col_keys = ['columns', 'separator']
        self.check_array_keys('s05_merge_columns', self.s05_merge_columns,
                              merge_col_keys)
        rename_col_keys = ['column', 'name']
        self.check_array_keys('s07_rename_columns', self.s07_rename_columns,
                              rename_col_keys)
        insert_col_keys = ['column', 'value']
        if insert_last is not None:
            assert insert_last is not None  # keep mypy happy
            insert_col_keys.append(insert_last)
        self.check_array_keys('s08_insert_columns', self.s08_insert_columns,
                              insert_col_keys)

    def check_rewrite_configs(self, coltype: type) -> None:
        """Check the rewrite column configuration."""
        rewrite_col_mand_keys: list[str] = ['column', 'kind', 'case']
        rewrite_col_opt_keys: list[str] = ['chars', 'from', 'to']
        self.check_array_keys('s09_rewrite_columns', self.s09_rewrite_columns,
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
        self.check_array_dicts(name_of_cfg='s09_rewrite_columns',
                               array=self.s09_rewrite_columns, kind_key='kind',
                               kind_type=RewriteKind,
                               dict_of_templates=template)

    @staticmethod
    def check_sep_not_sep(separators: list[str],
                          not_separators: list[str]) -> None:
        """Check relationship between separators and not separators."""
        for notsep in not_separators:
            if notsep in separators:
                print('Error in s01_split_rows:\n' +
                      'Cannot have same string as both separator and ' +
                      f'not separator: {notsep}', file=sys.stderr)
                sys.exit(1)
            found: bool = False
            for sep in separators:
                if sep in notsep:
                    found = True
                    break
            if not found:
                print('Error in s01_split_rows:\n' +
                      f'Not separator "{notsep}" does not affect ' +
                      'any separator.', file=sys.stderr)
                sys.exit(1)

    def check_split_row_cfg(self) -> None:
        """Check the split row configuration."""
        keys = ['column', 'separators', 'not_separators']
        self.check_lst_dict(paramname='s01_split_rows',
                            inp=self.s01_split_rows, key='column',
                            key_optional=False, valtype=self._columntype,
                            min_size_list=0)
        self.check_array_keys('s01_split_rows', self.s01_split_rows,
                              mandatory_keys=keys, allowed_keys=None)
        self.check_lst_dict_lst(paramname='s01_split_rows',
                                inp=self.s01_split_rows, key='separators',
                                key_optional=False, valtype=str,
                                min_size_outer_list=0, min_size_inner_list=1)
        self.check_lst_dict_lst(paramname='s01_split_rows',
                                inp=self.s01_split_rows, key='not_separators',
                                key_optional=False, valtype=str,
                                min_size_outer_list=0, min_size_inner_list=0)
        for elem in self.s01_split_rows:
            sep = elem['separators']
            assert isinstance(sep, list)
            nosep = elem['not_separators']
            assert isinstance(nosep, list)
            self.check_sep_not_sep(separators=sep, not_separators=nosep)

    def check_merge_row_cfg(self) -> None:
        """Check the merge rows configuration."""
        keys = ['columns', 'separator']
        self.check_lst_dict_lst(paramname='s02_merge_rows',
                                inp=self.s02_merge_rows, key='columns',
                                key_optional=False, valtype=self._columntype,
                                min_size_outer_list=0, min_size_inner_list=1)
        self.check_lst_dict(paramname='s02_merge_rows',
                            inp=self.s02_merge_rows, key='separator',
                            key_optional=False, valtype=str, min_size_list=0)
        self.check_array_keys('s02_merge_rows', self.s02_merge_rows,
                              mandatory_keys=keys, allowed_keys=None)
