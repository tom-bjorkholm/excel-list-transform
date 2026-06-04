#! /usr/local/bin/python3
"""Read and write configuration of excel list transform."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from csv import Dialect
from copy import deepcopy
from enum import Enum
from typing import Optional, Callable, TypeVar, NamedTuple, \
    Generic, TypedDict, TextIO, override
from config_as_json import Config, ConfigAutoChangeHook, ConfigNesting, \
    ConfigNestingKind, NestedConfigs, ParseConverter, PathOrStr, \
    ReadOldConfiguration, RocfKeyMove, RocfKeyRename, ValidationPlan
from tableio import CAP_IGNORABLE, CAP_NEEDED, Capabilities, CsvDialect, \
    FileAccess, TableBorderStyle
from tableio_cfg_json import TioJsonConfig, TioJsonCsvConfig
from excel_list_transform.config_enums import FileType, SplitWhere, \
    ExcelLib, RewriteKind, CaseSensitivity, ColumnRef
from excel_list_transform.str_to_enum import string_to_enum_best_match
from excel_list_transform.migrate_cfg_warn_hook import MigrateCfgWarnHook
from excel_list_transform.config import Config as OldConfig


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
type DupCheckKeyArg[Column] = \
    SingleRule[Column] | SingleRuleMerge[Column] | SingleRuleSplit[Column]
type DupCheckKey[Column] = \
    Callable[[DupCheckKeyArg[Column]], Column | list[Column]]
type IncrCheckKey[Column] = \
    Callable[[SingleRuleMerge[Column]], Column | list[Column]]
type NoDupKeydType[Column] = \
    Rule[Column] | RuleMerge[Column] | RuleSplit[Column]


def input_capabilities() -> Capabilities:
    """Return TableIO capabilities used for input files."""
    return Capabilities(can_read=CAP_NEEDED)


def output_capabilities() -> Capabilities:
    """Return TableIO capabilities used for output files."""
    return Capabilities(can_write=CAP_NEEDED,
                        filtered_data_range=CAP_IGNORABLE,
                        can_write_borders=CAP_IGNORABLE)


class InputTableConfig(TioJsonConfig):
    """JSON-backed TableIO configuration for input files."""

    def __init__(self, from_json_data_text: Optional[str] = None,
                 from_json_filename: Optional[PathOrStr] = None,
                 auto_ch_hook: Optional[ConfigAutoChangeHook] = None,
                 stderr_file: TextIO = sys.stderr) -> None:
        """Construct input TableIO configuration."""
        super().__init__(capabilities=input_capabilities(),
                         file_access=FileAccess.READ,
                         from_json_data_text=from_json_data_text,
                         from_json_filename=from_json_filename,
                         auto_ch_hook=auto_ch_hook, stderr_file=stderr_file)


class OutputTableConfig(TioJsonConfig):
    """JSON-backed TableIO configuration for output files."""

    def __init__(self, from_json_data_text: Optional[str] = None,
                 from_json_filename: Optional[PathOrStr] = None,
                 auto_ch_hook: Optional[ConfigAutoChangeHook] = None,
                 stderr_file: TextIO = sys.stderr) -> None:
        """Construct output TableIO configuration."""
        super().__init__(capabilities=output_capabilities(),
                         file_access=FileAccess.CREATE,
                         from_json_data_text=from_json_data_text,
                         from_json_filename=from_json_filename,
                         auto_ch_hook=auto_ch_hook, stderr_file=stderr_file)


def _old_value_name(value: object) -> str:
    """Return a normalized old configuration enum/string name."""
    if isinstance(value, Enum):
        return value.name
    assert isinstance(value, str)
    return value


def old_file_type_to_format(value: object) -> str:
    """Convert old input/output file type to a TableIO format name."""
    name = _old_value_name(value).lower()
    if name == 'csv':
        return 'CSV'
    if name == 'excel':
        return 'Excel'
    raise KeyError(f'Unknown old file type: {value}')


def _old_quoting_to_tableio(value: object) -> Optional[str]:
    """Convert old csv.QUOTE_* config text to TableIO text."""
    if value is None:
        return None
    assert isinstance(value, str)
    values = {'csv.quote_all': 'all',
              'csv.quote_minimal': 'minimal',
              'csv.quote_none': 'none',
              'csv.quote_nonnumeric': 'nonnumeric'}
    lower = value.lower()
    if lower not in values:
        raise KeyError(f'Unknown old CSV quoting: {value}')
    return values[lower]


def old_csv_spec_to_tableio(value: object) -> dict[str, object]:
    """Convert old CSV dialect config to tableio-cfg-json shape."""
    assert isinstance(value, dict)
    name_obj = value.get('name')
    name = '' if name_obj is None else str(name_obj).lower()
    dialect = CsvDialect.EXCEL.name
    if name == 'csv.unix_dialect':
        dialect = CsvDialect.UNIX.name
    ret: dict[str, object] = {'dialect': dialect}
    delimiter = value.get('delimiter')
    if delimiter is None and name == 'csv.excel_tab':
        delimiter = '\t'
    if delimiter is not None:
        ret['delimiter'] = delimiter
    quoting = _old_quoting_to_tableio(value.get('quoting'))
    if quoting is not None:
        ret['quoting'] = quoting
    for old_key in ['quotechar', 'lineterminator', 'escapechar']:
        item = value.get(old_key)
        if item is not None:
            ret[old_key] = item
    return ret


class ConfigReadOld(ReadOldConfiguration):
    """Normalize old excel-list-transform configuration files."""

    _OLD_SHAPE_KEYS = ['in_type', 'out_type', 'in_csv_dialect',
                       'out_csv_dialect', 'in_csv_encoding',
                       'out_csv_encoding', 'in_excel_library',
                       'out_excel_library', 's1_split_columns',
                       's2_remove_columns', 's3_merge_columns',
                       's4_place_columns_first', 's5_rename_columns',
                       's6_insert_columns', 's7_rewrite_columns',
                       's8_column_order']

    def pre_process_json(self, json_data: dict[str, object],
                         auto_ch_hook: ConfigAutoChangeHook,
                         stderr_file: TextIO) -> dict[str, object]:
        """Add old implicit defaults before declarative migration."""
        _ = auto_ch_hook
        _ = stderr_file
        if not self._is_old_shape(json_data):
            return json_data
        for key, value in self._old_defaults().items():
            if key not in json_data:
                json_data[key] = deepcopy(value)
        return json_data

    @classmethod
    def _is_old_shape(cls, json_data: dict[str, object]) -> bool:
        """Return whether parsed JSON looks like the old flat format."""
        return any(key in json_data for key in cls._OLD_SHAPE_KEYS)

    @staticmethod
    def _old_defaults() -> dict[str, object]:
        """Return defaults needed by older supported config files."""
        return {'in_type': FileType.EXCEL.name,
                'out_type': FileType.EXCEL.name,
                'in_csv_encoding': 'utf_8_sig',
                'out_csv_encoding': 'utf-8',
                'in_excel_col_name_strip': False,
                'in_excel_values_strip': False,
                's01_split_rows': [],
                's02_merge_rows': [],
                'output_borders': False,
                'output_filtered_table': False}

    @override
    def get_keys_to_prune(self) -> list[str]:
        """Return old keys accepted and discarded during migration."""
        return ['in_excel_library', 'out_excel_library']

    @override
    def get_json_key_renames(self) -> list[RocfKeyRename]:
        """Return old transform rule key renames."""
        return [
            RocfKeyRename(old='s1_split_columns', new='s03_split_columns'),
            RocfKeyRename(old='s2_remove_columns', new='s04_remove_columns'),
            RocfKeyRename(old='s3_merge_columns', new='s05_merge_columns'),
            RocfKeyRename(old='s4_place_columns_first',
                          new='s06_place_columns_first'),
            RocfKeyRename(old='s5_rename_columns', new='s07_rename_columns'),
            RocfKeyRename(old='s6_insert_columns', new='s08_insert_columns'),
            RocfKeyRename(old='s7_rewrite_columns', new='s09_rewrite_columns'),
            RocfKeyRename(old='s8_column_order', new='s10_column_order')
        ]

    @override
    def get_json_key_moves(self) -> list[RocfKeyMove]:
        """Return old I/O settings moved into TableIO config sections."""
        return [
            RocfKeyMove(old_path=('in_type',),
                        new_path=('input_table', 'format_name'),
                        transform_value=old_file_type_to_format),
            RocfKeyMove(old_path=('out_type',),
                        new_path=('output_table', 'format_name'),
                        transform_value=old_file_type_to_format),
            RocfKeyMove(old_path=('in_csv_encoding',),
                        new_path=('input_table', 'character_encoding')),
            RocfKeyMove(old_path=('out_csv_encoding',),
                        new_path=('output_table', 'character_encoding')),
            RocfKeyMove(old_path=('in_csv_dialect',),
                        new_path=('input_table', 'csv'),
                        transform_value=old_csv_spec_to_tableio),
            RocfKeyMove(old_path=('out_csv_dialect',),
                        new_path=('output_table', 'csv'),
                        transform_value=old_csv_spec_to_tableio)
        ]


class ColInfo(NamedTuple, Generic[Column]):
    """Information about columns to pass to ConfigExcelListTransform init."""

    split_last: str
    insert_last: Optional[str]
    s03: RuleSplit[Column]
    s08: Rule[Column]
    col_to_use: list[Column]
    col_to_use_row: list[Column]
    tinfo: Column


class ConfigExcelListTransform(Config, Generic[Column]):  # pylint: disable=too-many-instance-attributes,too-many-public-methods,line-too-long # noqa: E501
    """Class with configuration for excel list transform."""

    def __init__(self, *, col_ref: ColumnRef,  # pylint: disable=too-many-arguments # noqa: E501
                 colinfo: ColInfo[Column], tinfo: Column,
                 from_json_text: Optional[str] = None,
                 from_json_filename: Optional[PathOrStr] = None,
                 auto_ch_hook: ConfigAutoChangeHook =
                 MigrateCfgWarnHook(),
                 stderr_file: Optional[TextIO] = None) -> None:
        """Construct configuration for excel list transform."""
        err_file = sys.stderr if stderr_file is None else stderr_file
        assert isinstance(colinfo.tinfo, type(tinfo))
        self._colinfo: ColInfo[Column] = deepcopy(colinfo)
        self._columntype: type[Column] = type(tinfo)
        self.column_ref: ColumnRef = col_ref
        col2use = deepcopy(colinfo.col_to_use)  # dont destroying caller's arg
        col2userow = deepcopy(colinfo.col_to_use_row)
        self.max_column_read: int = 20
        self.input_table: InputTableConfig = InputTableConfig()
        self.output_table: OutputTableConfig = OutputTableConfig()
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
        super().__init__(from_json_data_text=from_json_text,
                         from_json_filename=from_json_filename,
                         auto_ch_hook=auto_ch_hook, stderr_file=err_file)

    @override
    def nested_configs(self) -> NestedConfigs:
        """Return nested TableIO config sections."""
        return {
            'input_table': ConfigNesting(kind=ConfigNestingKind.MEMBER,
                                         config_type=InputTableConfig),
            'output_table': ConfigNesting(kind=ConfigNestingKind.MEMBER,
                                          config_type=OutputTableConfig)
        }

    @override
    def get_validation_plan(self, stderr_file: TextIO) -> ValidationPlan:
        """Return validation plan for config-as-json."""
        _ = stderr_file
        return []

    @override
    def validate(self, stderr_file: TextIO) -> None:
        """Validate nested TableIO config and transform-rule settings."""
        super().validate(stderr_file=stderr_file)
        colinfo = self._colinfo
        self.check_array_configs(split_last=colinfo.split_last,
                                 insert_last=colinfo.insert_last)
        self.sort_sx_hook()
        self._check_no_duplicate_single(self.s03_split_columns,
                                        's03_split_columns', colinfo.tinfo)
        self._check_no_duplicate_single(self.s07_rename_columns,
                                        's07_rename_columns', colinfo.tinfo)
        self._check_no_duplicate_single(self.s08_insert_columns,
                                        's08_insert_columns', colinfo.tinfo)
        self.check_rewrite_configs(coltype=type(colinfo.tinfo))
        self.check_split_row_cfg()
        self.check_merge_row_cfg()

    @override
    def as_json_string(self, stderr_file: Optional[TextIO] = None) -> str:
        """Return string with this configuration as JSON data."""
        err_file = sys.stderr if stderr_file is None else stderr_file
        return super().as_json_string(stderr_file=err_file)

    def table_border_style(self) -> TableBorderStyle:
        """Return the TableIO border style requested by the config."""
        if self.output_borders:
            return TableBorderStyle.OUTER_FIRST_ROW_THICK_INNER_THIN
        return TableBorderStyle.NONE

    @property
    def in_csv_encoding(self) -> str:
        """Return old-style input CSV encoding for internal callers."""
        return self.input_table.character_encoding or 'utf_8_sig'

    @in_csv_encoding.setter
    def in_csv_encoding(self, value: str) -> None:
        """Set input character encoding using old-style member name."""
        self.input_table.character_encoding = value

    @property
    def out_csv_encoding(self) -> str:
        """Return old-style output CSV encoding for internal callers."""
        return self.output_table.character_encoding or 'utf-8'

    @out_csv_encoding.setter
    def out_csv_encoding(self, value: str) -> None:
        """Set output character encoding using old-style member name."""
        self.output_table.character_encoding = value

    @property
    def in_type(self) -> FileType:
        """Return old-style input file type for legacy callers."""
        return self._old_file_type(self.input_table.format_name)

    @in_type.setter
    def in_type(self, value: FileType) -> None:
        """Set input format using old-style file type."""
        self.input_table.format_name = old_file_type_to_format(value)
        if value == FileType.CSV:
            self._ensure_old_csv_defaults(self.input_table)

    @property
    def out_type(self) -> FileType:
        """Return old-style output file type for legacy callers."""
        return self._old_file_type(self.output_table.format_name)

    @out_type.setter
    def out_type(self, value: FileType) -> None:
        """Set output format using old-style file type."""
        self.output_table.format_name = old_file_type_to_format(value)
        if value == FileType.CSV:
            self._ensure_old_csv_defaults(self.output_table)

    @staticmethod
    def _old_file_type(format_name: str) -> FileType:
        """Return old-style file type for a TableIO format name."""
        if format_name.lower() == 'csv':
            return FileType.CSV
        return FileType.EXCEL

    @staticmethod
    def _csv_config(table_cfg: TioJsonConfig) -> TioJsonCsvConfig:
        """Return existing or newly created nested CSV configuration."""
        if table_cfg.csv is None:
            table_cfg.csv = TioJsonCsvConfig()
        assert isinstance(table_cfg.csv, TioJsonCsvConfig)
        return table_cfg.csv

    @staticmethod
    def _ensure_old_csv_defaults(table_cfg: TioJsonConfig) -> None:
        """Provide old comma CSV defaults for old-style setters."""
        if table_cfg.csv is None:
            table_cfg.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                             delimiter=',', quotechar='"')

    @staticmethod
    def _csv_dialect_name(csv_cfg: TioJsonCsvConfig) -> str:
        """Return old-style CSV dialect name for a TableIO CSV config."""
        if csv_cfg.dialect == CsvDialect.UNIX:
            return 'csv.unix_dialect'
        if csv_cfg.delimiter == '\t':
            return 'csv.excel_tab'
        return 'csv.excel'

    @staticmethod
    def _csv_quoting_name(quoting: Optional[str]) -> Optional[str]:
        """Return old-style CSV quoting name for a TableIO quoting name."""
        if quoting is None:
            return None
        values = {'all': 'csv.quote_all',
                  'minimal': 'csv.quote_minimal',
                  'none': 'csv.quote_none',
                  'nonnumeric': 'csv.quote_nonnumeric'}
        return values[quoting]

    @classmethod
    def _csv_spec(cls, table_cfg: TioJsonConfig) -> dict[str, Optional[str]]:
        """Return old-style CSV spec for internal CSV test helpers."""
        csv_cfg = cls._csv_config(table_cfg)
        return {'name': cls._csv_dialect_name(csv_cfg),
                'delimiter': csv_cfg.delimiter,
                'quoting': cls._csv_quoting_name(csv_cfg.quoting),
                'quotechar': csv_cfg.quotechar,
                'lineterminator': csv_cfg.lineterminator,
                'escapechar': csv_cfg.escapechar}

    def get_out_csv_dialect(self) -> type[Dialect]:
        """Return old-style CSV dialect for output test helpers."""
        return OldConfig.get_csv_dialect(**self._csv_spec(self.output_table))

    def get_in_csv_dialect(self) -> type[Dialect]:
        """Return old-style CSV dialect for input test helpers."""
        return OldConfig.get_csv_dialect(**self._csv_spec(self.input_table))

    def sort_sx_hook(self) -> None:
        """Sort s[0-9]_ as needed (hook)."""

    check_array_keys = staticmethod(OldConfig.check_array_keys)
    check_lst_dict = staticmethod(OldConfig.check_lst_dict)
    check_lst_dict_lst = staticmethod(OldConfig.check_lst_dict_lst)
    check_array_dicts = staticmethod(OldConfig.check_array_dicts)
    check_no_duplicates = staticmethod(OldConfig.check_no_duplicates)

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
    def _check_no_duplicate_single(rule: Rule[Column] | RuleSplit[Column],
                                   param_name: str, tinfo: Column) -> None:
        """Flag as error if column is refered to multiple times."""
        cols: list[Column] = ConfigExcelListTransform.get_cols_single(rule,
                                                                      tinfo)
        ConfigExcelListTransform.check_no_duplicates(cols, param_name)

    @staticmethod
    def _check_no_duplicate_multi(rule: RuleMerge[Column], param_name: str,
                                  tinfo: Column) -> None:
        """Flag as error if column is refered to multiple times."""
        cols: list[Column] = ConfigExcelListTransform.get_cols_multi(rule,
                                                                     tinfo)
        ConfigExcelListTransform.check_no_duplicates(cols, param_name)

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


def _old_cfg_reader(self: object) -> ReadOldConfiguration:
    """Return old-format migration rules."""
    _ = self
    return ConfigReadOld()


setattr(ConfigExcelListTransform, '_get_read_old_configuration',
        _old_cfg_reader)
