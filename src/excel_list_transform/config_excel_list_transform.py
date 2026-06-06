#! /usr/local/bin/python3
"""Read and write configuration of excel list transform."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


import sys
from copy import deepcopy
from enum import Enum
from typing import Generic, Mapping, NamedTuple, NoReturn, Optional, TextIO, \
    TypeVar, TypedDict, override
from config_as_json import Config, ConfigAutoChangeHook, ConfigNesting, \
    ConfigNestingKind, InvalidConfiguration, ListIsOrderedValidator, \
    NestedConfigs, ParseConverter, PathOrStr, ProjectedMemberValidator, \
    ReadOldConfiguration, ValidationPlan, MemberValidationStep, \
    WholeConfigValidationStep, ValueTypeValidator, \
    CallingWholeConfigValidator, MemberValidator, DictKeysValidator, \
    DictRule, DictForEachValidator, DictVariant, \
    DiscriminatedDictValidator, ListForEachValidator, \
    ListSizeValidator, ListValueTypeValidator
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
    """Single rule for splitting a column."""

    column: Column
    separators: list[str]
    not_separators: list[str]


type RuleSplit[Column] = list[SingleRuleSplit[Column]]
type RuleRowSplit[Column] = list[SingleRuleRowSplit[Column]]
type RuleOrder = list[str]
type RulePlace = list[int]
type RuleRemove = RulePlace
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


def _invalid_rule(member_name: str, message: str,
                  stderr_file: TextIO) -> NoReturn:
    """Raise a readable validation error for one transform rule."""
    msg = f'Invalid configuration: {member_name} {message}'
    print(msg, file=stderr_file)
    raise InvalidConfiguration(msg)


# pylint: disable-next=too-few-public-methods
class SplitRowSepValidator(MemberValidator):
    """Validate separator and not-separator relationships in one rule."""

    def validate_member(self, config: Config, member_name: str,
                        member_value: object,
                        stderr_file: TextIO = sys.stderr) -> Optional[object]:
        """Validate that every not-separator affects a real separator."""
        _ = config
        if not isinstance(member_value, dict):
            _invalid_rule(member_name, 'must be a dict.', stderr_file)
        separators = member_value.get('separators')
        not_separators = member_value.get('not_separators')
        if not isinstance(separators, list) or \
                not isinstance(not_separators, list):
            _invalid_rule(member_name, 'must contain separator lists.',
                          stderr_file)
        for not_sep in not_separators:
            assert isinstance(not_sep, str)
            if not_sep in separators:
                msg = 'must not use the same string as a separator and '
                msg += f'not-separator: {not_sep!r}.'
                _invalid_rule(member_name, msg, stderr_file)
            if not any(sep in not_sep for sep in separators):
                msg = f'has not-separator {not_sep!r} that does not '
                msg += 'affect any separator.'
                _invalid_rule(member_name, msg, stderr_file)
        return member_value


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
        self.output_borders: bool = True
        self.output_filtered_table: bool = True
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
        plan: ValidationPlan = [
            MemberValidationStep(member_names=['max_column_read'],
                                 validator=ValueTypeValidator(
                                     int, not_allowed_type=bool)),
            MemberValidationStep(member_names=bool_names,
                                 validator=ValueTypeValidator(bool))]
        plan.extend(self._base_rule_shape_steps())
        plan.extend(self.column_shape_steps())
        plan.append(
            WholeConfigValidationStep(validator=CallingWholeConfigValidator(
                method_name='sort_sx_hook')))
        plan.extend(self._base_rule_val_steps())
        plan.extend(self.get_column_val_steps())
        return plan

    def _base_rule_shape_steps(self) -> list[MemberValidationStep]:
        """Return structural validation steps for shared transform rules."""
        return [
            self._split_rows_step(),
            self._merge_rules_step('s02_merge_rows'),
            self._split_columns_step(),
            self._merge_rules_step('s05_merge_columns'),
            self._rename_columns_step(),
            self._insert_columns_step(),
            self._rewrite_columns_step()]

    def column_shape_steps(self) -> list[MemberValidationStep]:
        """Return pre-normalization validation for column-specific members."""
        return []

    def _base_rule_val_steps(self) -> list[MemberValidationStep]:
        """Return validation steps for shared transform-rule members."""
        return [
            self._rule_unique_step('s03_split_columns'),
            self._rule_unique_step('s07_rename_columns'),
            self._rule_unique_step('s08_insert_columns')]

    def get_column_val_steps(self) -> list[MemberValidationStep]:
        """Return validation steps for column-reference-specific members."""
        return []

    def _column_validator(self) -> MemberValidator:
        """Return validator for one configured column reference."""
        if self._columntype is int:
            return ValueTypeValidator(int, not_allowed_type=bool)
        return ValueTypeValidator(str)

    def _column_list_validator(self) -> ListForEachValidator:
        """Return validator for a list of configured column references."""
        return ListForEachValidator(
            element_validators=[self._column_validator()])

    def _plain_column_list_step(self, member_name: str) \
            -> MemberValidationStep:
        """Return validation step for a plain list of column references."""
        return MemberValidationStep(member_names=[member_name],
                                    validator=self._column_list_validator())

    def _split_rows_step(self) -> MemberValidationStep:
        """Return validation step for split-row rules."""
        values_validator = DictForEachValidator(rules=[
            DictRule(keys=['column'], validators=[self._column_validator()]),
            DictRule(keys=['separators'],
                     validators=[ListSizeValidator(1, sys.maxsize),
                                 ListValueTypeValidator(str)]),
            DictRule(keys=['not_separators'],
                     validators=[ListValueTypeValidator(str)])])
        validator = ListForEachValidator(
            element_validators=[
                DictKeysValidator(
                    mandatory_keys=['column', 'separators',
                                    'not_separators']),
                values_validator,
                SplitRowSepValidator()],
            element_type=dict)
        return MemberValidationStep(member_names=['s01_split_rows'],
                                    validator=validator)

    def _merge_rules_step(self, member_name: str) -> MemberValidationStep:
        """Return validation step for merge-row or merge-column rules."""
        values_validator = DictForEachValidator(rules=[
            DictRule(keys=['columns'],
                     validators=[ListSizeValidator(1, sys.maxsize),
                                 self._column_list_validator()]),
            DictRule(keys=['separator'],
                     validators=[ValueTypeValidator(str)])])
        validator = ListForEachValidator(
            element_validators=[
                DictKeysValidator(mandatory_keys=['columns', 'separator']),
                values_validator],
            element_type=dict)
        return MemberValidationStep(member_names=[member_name],
                                    validator=validator)

    def _split_columns_step(self) -> MemberValidationStep:
        """Return validation step for split-column rules."""
        keys = ['column', 'separator', 'where', self._colinfo.split_last]
        values_validator = DictForEachValidator(rules=[
            DictRule(keys=['column'], validators=[self._column_validator()]),
            DictRule(keys=['separator'], validators=[ValueTypeValidator(str)]),
            DictRule(keys=['where'],
                     validators=[ValueTypeValidator(SplitWhere)]),
            self._split_last_rule()])
        validator = ListForEachValidator(
            element_validators=[
                DictKeysValidator(mandatory_keys=keys),
                values_validator],
            element_type=dict)
        return MemberValidationStep(member_names=['s03_split_columns'],
                                    validator=validator)

    def _split_last_rule(self) -> DictRule:
        """Return value validation for the reference-specific split key."""
        if self._columntype is int:
            value_validator = ValueTypeValidator(SplitWhere)
        else:
            value_validator = ValueTypeValidator(str)
        return DictRule(keys=[self._colinfo.split_last],
                        validators=[value_validator])

    def _rename_columns_step(self) -> MemberValidationStep:
        """Return validation step for rename-column rules."""
        values_validator = DictForEachValidator(rules=[
            DictRule(keys=['column'], validators=[self._column_validator()]),
            DictRule(keys=['name'], validators=[ValueTypeValidator(str)])])
        validator = ListForEachValidator(
            element_validators=[
                DictKeysValidator(mandatory_keys=['column', 'name']),
                values_validator],
            element_type=dict)
        return MemberValidationStep(member_names=['s07_rename_columns'],
                                    validator=validator)

    def _insert_columns_step(self) -> MemberValidationStep:
        """Return validation step for insert-column rules."""
        keys = ['column', 'value']
        rules = [
            DictRule(keys=['column'], validators=[self._column_validator()])]
        if self._colinfo.insert_last is not None:
            insert_last = self._colinfo.insert_last
            keys.append(insert_last)
            rules.append(DictRule(keys=[insert_last],
                                  validators=[ValueTypeValidator(str)]))
        validator = ListForEachValidator(
            element_validators=[
                DictKeysValidator(mandatory_keys=keys),
                DictForEachValidator(rules=rules)],
            element_type=dict)
        return MemberValidationStep(member_names=['s08_insert_columns'],
                                    validator=validator)

    def _rewrite_columns_step(self) -> MemberValidationStep:
        """Return validation step for rewrite-column rules."""
        column_rule = DictRule(keys=['column'],
                               validators=[self._column_validator()])
        case_rule = DictRule(keys=['case'],
                             validators=[ValueTypeValidator(CaseSensitivity)])
        chars_str_rule = DictRule(keys=['chars'],
                                  validators=[ValueTypeValidator(str)])
        chars_list_rule = DictRule(keys=['chars'],
                                   validators=[ListValueTypeValidator(str)])
        str_rule = DictRule(keys=['from', 'to'],
                            validators=[ValueTypeValidator(str)])
        variants: Mapping[object, DictVariant] = {
            RewriteKind.STRIP: DictVariant(
                mandatory_keys=['column', 'case', 'chars'],
                rules=[column_rule, case_rule, chars_str_rule]),
            RewriteKind.REMOVECHARS: DictVariant(
                mandatory_keys=['column', 'case', 'chars'],
                rules=[column_rule, case_rule, chars_list_rule]),
            RewriteKind.REGEX_SUBSTITUTE: DictVariant(
                mandatory_keys=['column', 'case', 'from', 'to'],
                rules=[column_rule, case_rule, str_rule]),
            RewriteKind.STR_SUBSTITUTE: DictVariant(
                mandatory_keys=['column', 'case', 'from', 'to'],
                rules=[column_rule, case_rule, str_rule])}
        one_rule_validator = DiscriminatedDictValidator(
            discriminator_key='kind', variants=variants,
            discriminator_validator=ValueTypeValidator(RewriteKind))
        validator = ListForEachValidator(
            element_validators=[one_rule_validator], element_type=dict)
        return MemberValidationStep(member_names=['s09_rewrite_columns'],
                                    validator=validator)

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

    def table_border_style(self) -> TableBorderStyle:
        """Return the TableIO border style requested by the config."""
        if self.output_borders:
            return TableBorderStyle.OUTER_FIRST_ROW_THICK_INNER_THIN
        return TableBorderStyle.NONE

    def sort_sx_hook(self) -> None:
        """Sort s[0-9]_ as needed (hook)."""

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
