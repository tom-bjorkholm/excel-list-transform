#! /usr/local/bin/python3
"""Generate example configuration files."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License

from copy import deepcopy
from tableio import CsvDialect
from tableio_cfg_json import TioJsonCsvConfig
from excel_list_transform.file_extension import fix_file_extension
from excel_list_transform.config_enums import ColumnRef, RewriteKind, \
    CaseSensitivity, SplitWhere
from excel_list_transform.generate_txt import generate_syntax_txt
from excel_list_transform.config_excel_list_transform import \
    Column, SingleRuleRewrite, RuleRewrite
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum

SYNTAX_PHONE_FIX: str = '''
The phone number has to be in international format '+' followed by only
digits for RRS, but the sailors filling in the form sometimes use
national format (no + and no country code). Also the sailors filling
in the form are used to group digits by dashes and spaces for easy
reading. The s09_rewrite_columns include rewrite instructions to
"fix" the phone number to the format needed by RRS.
The rewriting of phone numbers in this example is for Swedish mobile
phone numbers. For other countries s09_rewrite_columns need to be
adjusted.
'''

RRS_COLUMNS: tuple[str, ...] = (
    'Class', 'Division', 'Nationality', 'MNA No.', 'Sail Number',
    'Boat Name', 'First Name', 'Last Name', 'Club Name', 'Email', 'Phone',
    'Whats App Number')

SYNYAX_O2X_COMMON: str = '''

Column names are changed from another language (Swedish) and sometimes from
a more instructive naming to the names understood by RRS.

The events does not use divisions, MNA numbers, boat names or Whats App
numbers. Thus the form did not ask for these. They are inserted as empty
columns to please RRS.

''' + SYNTAX_PHONE_FIX

SYNTAX_ONLY_O2R_COMMON: str = '''
This is an example of taking a registration list from the table file received
from office forms or google forms and transforming it for importing into
RRS ( https://www.racingrulesofsailing.org ).
'''

BY_NAME_COMMON: str = '''

s10_column_order both tells the order the columns shall be written
in, and what columns to write out. The columns not listed here
are ignored (and thus removed from the output).

This configuration files references columns "BY_NAME", which is the
preferred way to reference columns.
'''

BY_NUMBER_COMMON: str = '''

This configuration files references columns "BY_NUMBER".
The preferred way to reference columns in configuration files
is "BY_NAME".
'''

SYNTAX_O2R_COMMON: str = SYNTAX_ONLY_O2R_COMMON + SYNYAX_O2X_COMMON

SYNTAX_SW2R_COMMON: str = '''

In this example the input data is exported as JSON or XML from SailWave
https://www.sailwave.com . The data as a list in excel format is then
extracted using extract-list https://pypi.org/project/extract-list. As
SailWave only has a name field and not first name and last name fields, and as
SailWave has no fields for MNA No. and Whats App Number, this processing of
the data is needed before it can be imported into RRS
https://www.racingrulesofsailing.org/

The Name column is split into First Name and Last Name columns.
The columns MNA No. and Whats App Number are added with no data.

''' + SYNTAX_PHONE_FIX

SYNTAX_SA2R_COMMON: str = '''

In this example the input data is exported from SailArena to RRS.
This input format is a comma separated values text file, with the
old character encoding scheme cp1252.

SailArena still exports the WhatsApp column with its old name. The only
other fixes needed are to read CSV in cp1252 encoding, add MNA No., fix the
format of the phone number and save to excel.
''' + SYNTAX_PHONE_FIX


def rewrite_phone_46_cfg(cfg: ConfigXlsListTransfName | ConfigXlsListTransfNum,
                         column: Column, append: bool) -> None:
    """Add configuration to rewrite +46 phone numbers."""
    assert (isinstance(cfg, ConfigXlsListTransfName) and
            isinstance(column, str)) or \
           (isinstance(cfg, ConfigXlsListTransfNum) and
            isinstance(column, int))
    assert isinstance(append, bool)
    rules: RuleRewrite[Column] = [
        {'kind': RewriteKind.STRIP,
         'chars': '', 'case': CaseSensitivity.IGNORE_CASE},
        {'kind': RewriteKind.REMOVECHARS,
         'chars': [' ', '-', '(', ')'], 'case': CaseSensitivity.MATCH_CASE},
        {'kind': RewriteKind.REGEX_SUBSTITUTE,
         'from': '^07', 'to': '+467', 'case': CaseSensitivity.MATCH_CASE},
        {'kind': RewriteKind.REGEX_SUBSTITUTE,
         'from': '^\\+4607', 'to': '+467', 'case': CaseSensitivity.MATCH_CASE},
        {'kind': RewriteKind.REGEX_SUBSTITUTE,
         'from': '^467', 'to': '+467', 'case': CaseSensitivity.MATCH_CASE},
        {'kind': RewriteKind.REGEX_SUBSTITUTE,
         'from': '^4607', 'to': '+467', 'case': CaseSensitivity.MATCH_CASE}
    ]
    if not append:
        cfg.s09_rewrite_columns.clear()
    for rule in rules:
        newrule: SingleRuleRewrite[Column] = deepcopy(rule)
        newrule['column'] = column
        cfg.s09_rewrite_columns.append(newrule)


def generate_syntax_sa2r_name(filename: str, colref: ColumnRef) -> None:
    """Generate config example for sa_to_rrs."""
    assert colref == ColumnRef.BY_NAME
    cfg = ConfigXlsListTransfName()
    set_input_csv(cfg=cfg, encoding='cp1252', delimiter=';')
    set_output_excel(cfg)
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = []
    cfg.s05_merge_columns = []
    cfg.s07_rename_columns = [{'column': 'WhatsApp',
                               'name': 'Whats App Number'}]
    cfg.s08_insert_columns = [{'column': 'MNA No.', 'value': None}]
    rewrite_phone_46_cfg(cfg=cfg, column='Phone', append=False)
    cfg.s10_column_order = list(RRS_COLUMNS)
    cfg.write(to_json_filename=filename)


def generate_syntax_sw2r_name(filename: str, colref: ColumnRef) -> None:
    """Generate config example for sw_to_rrs."""
    assert colref == ColumnRef.BY_NAME
    cfg = ConfigXlsListTransfName()
    set_input_excel(cfg)
    set_output_excel(cfg)
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = [{'column': 'Name',
                              'separator': ' ',
                              'where': SplitWhere.RIGHTMOST,
                              'right_name': 'Last Name'}]
    cfg.s05_merge_columns = []
    cfg.s07_rename_columns = [{'column': 'Name', 'name': 'First Name'}]
    cfg.s08_insert_columns = [{'column': 'MNA No.', 'value': None},
                              {'column': 'Whats App Number', 'value': None}]
    rewrite_phone_46_cfg(cfg=cfg, column='Phone', append=False)
    cfg.s10_column_order = list(RRS_COLUMNS)
    cfg.write(to_json_filename=filename)


TXT_SW2R_NAME = SYNTAX_SW2R_COMMON + BY_NAME_COMMON
TXT_SA2R_NAME = SYNTAX_SA2R_COMMON + BY_NAME_COMMON


def config_from_colref(colref: ColumnRef) -> ConfigXlsListTransfName | \
        ConfigXlsListTransfNum:
    """Return a default example config for the selected column reference."""
    if colref == ColumnRef.BY_NAME:
        return ConfigXlsListTransfName()
    return ConfigXlsListTransfNum()


def set_input_csv(cfg: ConfigXlsListTransfName | ConfigXlsListTransfNum,
                  encoding: str, delimiter: str) -> None:
    """Configure input as CSV."""
    cfg.input_table.format_name = 'CSV'
    cfg.input_table.character_encoding = encoding
    cfg.input_table.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                           delimiter=delimiter, quotechar='"')


def set_input_excel(cfg: ConfigXlsListTransfName | ConfigXlsListTransfNum
                    ) -> None:
    """Configure input as Excel."""
    cfg.input_table.format_name = 'Excel'


def set_output_csv(cfg: ConfigXlsListTransfName | ConfigXlsListTransfNum
                   ) -> None:
    """Configure output as CSV."""
    cfg.output_table.format_name = 'CSV'
    cfg.output_table.character_encoding = 'utf-8'
    cfg.output_table.csv = TioJsonCsvConfig(dialect=CsvDialect.EXCEL,
                                            delimiter=',', quotechar='"')


def set_output_excel(cfg: ConfigXlsListTransfName | ConfigXlsListTransfNum
                     ) -> None:
    """Configure output as Excel."""
    cfg.output_table.format_name = 'Excel'
    cfg.output_table.implementation = None


def generate_syntax_sa2r_num(filename: str, colref: ColumnRef) -> None:
    """Generate config example for sa_to_rrs."""
    assert colref == ColumnRef.BY_NUMBER
    cfg = ConfigXlsListTransfNum()
    set_input_csv(cfg=cfg, encoding='cp1252', delimiter=';')
    set_output_excel(cfg)
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = []
    cfg.s04_remove_columns = []
    cfg.s05_merge_columns = []
    cfg.s06_place_columns_first = []
    cfg.s07_rename_columns = [{'column': 10, 'name': 'Whats App Number'}]
    cfg.s08_insert_columns = [{'column': 3, 'name': 'MNA No.', 'value': None}]
    rewrite_phone_46_cfg(cfg=cfg, column=10, append=False)
    cfg.write(to_json_filename=filename)


def generate_syntax_sw2r_num(filename: str, colref: ColumnRef) -> None:
    """Generate config example for sw_to_rrs."""
    assert colref == ColumnRef.BY_NUMBER
    cfg = ConfigXlsListTransfNum()
    set_input_excel(cfg)
    set_output_excel(cfg)
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = [{'column': 5, 'separator': ' ',
                              'where': SplitWhere.RIGHTMOST,
                              'store_single': SplitWhere.RIGHTMOST}]
    cfg.s04_remove_columns = []
    cfg.s05_merge_columns = []
    cfg.s06_place_columns_first = []
    cfg.s07_rename_columns = \
        [{'column': 5, 'name': 'First Name'},
         {'column': 6, 'name': 'Last Name'}]
    cfg.s08_insert_columns = [{'column': 3, 'name': 'MNA No.', 'value': None},
                              {'column': 11, 'name': 'Whats App Number',
                              'value': None}]
    rewrite_phone_46_cfg(cfg=cfg, column=10, append=False)
    cfg.write(to_json_filename=filename)


TXT_SW2R_NUM = SYNTAX_SW2R_COMMON + BY_NUMBER_COMMON
TXT_SA2R_NUM = SYNTAX_SA2R_COMMON + BY_NUMBER_COMMON


def generate_syntax_o2r_name(filename: str, colref: ColumnRef) -> None:
    """Generate config example for office_forms_to_rrs."""
    assert colref == ColumnRef.BY_NAME
    cfg = ConfigXlsListTransfName()
    set_input_excel(cfg)
    set_output_excel(cfg)
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = []
    cfg.s05_merge_columns = []
    cfg.s07_rename_columns = [
        {'column': 'Klass', 'name': 'Class'},
        {'column': 'Nationalitetsbokstäver i seglet', 'name': 'Nationality'},
        {'column': 'Segelnummer (endast siffror)', 'name': 'Sail Number'},
        {'column': 'Förnamn', 'name': 'First Name'},
        {'column': 'Efternamn', 'name': 'Last Name'},
        {'column': 'Klubb', 'name': 'Club Name'},
        {'column': 'Epostadress', 'name': 'Email'},
        {'column': 'Mobiltelefonnummer', 'name': 'Phone'}]
    cfg.s08_insert_columns = [
        {'column': 'MNA No.', 'value': None},
        {'column': 'Division', 'value': None},
        {'column': 'Boat Name', 'value': None},
        {'column': 'Whats App Number', 'value': None}]
    rewrite_phone_46_cfg(cfg=cfg, column='Phone', append=False)
    cfg.s10_column_order = list(RRS_COLUMNS)
    cfg.write(to_json_filename=filename)


TXT_O2R_NAME = SYNTAX_O2R_COMMON + BY_NAME_COMMON


def generate_syntax_o2r_num(filename: str, colref: ColumnRef) -> None:
    """Generate config example for office_forms_to_rrs."""
    assert colref == ColumnRef.BY_NUMBER
    cfg = ConfigXlsListTransfNum()
    set_input_excel(cfg)
    set_output_excel(cfg)
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = []
    cfg.s04_remove_columns = [0, 1, 2, 3, 4, 5, 14, 15, 16, 17, 18, 19]
    cfg.s05_merge_columns = []
    cfg.s06_place_columns_first = [5, 6, 7, 0, 1, 4, 2, 3]
    cfg.s07_rename_columns = [
        {'column': 0, 'name': 'Class'},
        {'column': 1, 'name': 'Nationality'},
        {'column': 2, 'name': 'Sail Number'},
        {'column': 3, 'name': 'First Name'},
        {'column': 4, 'name': 'Last Name'},
        {'column': 5, 'name': 'Club Name'},
        {'column': 6, 'name': 'Email'},
        {'column': 7, 'name': 'Phone'}]
    cfg.s08_insert_columns = [
        {'column': 1, 'name': 'Division', 'value': None},
        {'column': 3, 'name': 'MNA No.', 'value': None},
        {'column': 5, 'name': 'Boat Name', 'value': None},
        {'column': 11, 'name': 'Whats App Number', 'value': None}]
    rewrite_phone_46_cfg(cfg=cfg, column=10, append=False)
    cfg.write(to_json_filename=filename)


TXT_O2R_NUM = SYNTAX_O2R_COMMON + BY_NUMBER_COMMON

SYNTAX_ONLY_O2S_COMMON: str = '''
This is an example of taking a registration list from the table file received
from office forms or google forms and transforming it for importing into
SailWave regatta scoring program ( https://www.sailwave.com ).
'''

SYNTAX_O2S_COMMON: str = SYNTAX_ONLY_O2S_COMMON + SYNYAX_O2X_COMMON


def generate_syntax_o2s_name(filename: str, colref: ColumnRef) -> None:
    """Generate config example for office_forms_to_sw."""
    assert colref == ColumnRef.BY_NAME
    cfg = ConfigXlsListTransfName()
    set_input_excel(cfg)
    set_output_csv(cfg)
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = []
    cfg.s05_merge_columns = [{'columns': ['Förnamn', 'Efternamn'],
                             'separator': ' '}]
    cfg.s07_rename_columns = [
        {'column': 'Klass', 'name': 'Class'},
        {'column': 'Nationalitetsbokstäver i seglet', 'name': 'Nat'},
        {'column': 'Segelnummer (endast siffror)', 'name': 'SailNo'},
        {'column': 'Förnamn', 'name': 'HelmName'},
        {'column': 'Klubb', 'name': 'Club'},
        {'column': 'Epostadress', 'name': 'HelmEmail'},
        {'column': 'Mobiltelefonnummer', 'name': 'HelmPhone'}]
    cfg.s08_insert_columns = [
        {'column': 'Division', 'value': None},
        {'column': 'Boat', 'value': None}]
    rewrite_phone_46_cfg(cfg=cfg, column='HelmPhone', append=False)
    cfg.s10_column_order = ['Class', 'Division', 'Nat', 'SailNo', 'Boat',
                            'HelmName', 'Club', 'HelmEmail', 'HelmPhone']
    cfg.write(to_json_filename=filename)


TXT_O2S_NAME = SYNTAX_O2S_COMMON + BY_NAME_COMMON


def generate_syntax_o2s_num(filename: str, colref: ColumnRef) -> None:
    """Generate config example for office_forms_to_sw."""
    assert colref == ColumnRef.BY_NUMBER
    cfg = ConfigXlsListTransfNum()
    set_input_excel(cfg)
    set_output_csv(cfg)
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = []
    cfg.s04_remove_columns = [0, 1, 2, 3, 4, 5, 14, 15, 16, 17, 18, 19]
    cfg.s05_merge_columns = [{'columns': [0, 1], 'separator': ' '}]
    cfg.s06_place_columns_first = [4, 5, 6, 0, 3, 1, 2]
    cfg.s07_rename_columns = [
        {'column': 0, 'name': 'Class'},
        {'column': 1, 'name': 'Nat'},
        {'column': 2, 'name': 'SailNo'},
        {'column': 3, 'name': 'HelmName'},
        {'column': 4, 'name': 'Club'},
        {'column': 5, 'name': 'HelmEmail'},
        {'column': 6, 'name': 'HelmPhone'}]
    cfg.s08_insert_columns = [
        {'column': 1, 'name': 'Division', 'value': None},
        {'column': 4, 'name': 'Boat', 'value': None}]
    rewrite_phone_46_cfg(cfg=cfg, column=8, append=False)
    cfg.write(to_json_filename=filename)


TXT_O2S_NUM = SYNTAX_O2S_COMMON + BY_NUMBER_COMMON


SYNTAX_R2S_COMMON: str = '''
This is an example of taking a registration list from the table file
exported from RRS ( https://www.racingrulesofsailing.org ) and
transforming it for importing into SailWave regatta scoring
program ( https://www.sailwave.com ).

The names of the columns are changed from RRS names to SailWave
names. The phone number exported from RRS is in international
format, except that the '+' is missing.

The MNA No. and Whats App Number columns from RRS are not used by SailWave.

'''


def generate_syntax_r2s_name(filename: str, colref: ColumnRef) -> None:
    """Generate config example for rrs_to_sw."""
    assert colref == ColumnRef.BY_NAME
    cfg = ConfigXlsListTransfName()
    set_input_excel(cfg)
    set_output_csv(cfg)
    cfg.s03_split_columns = []
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s05_merge_columns = [{'columns': ['First Name', 'Last Name'],
                             'separator': ' '}]
    cfg.s07_rename_columns = [
        {'column': 'Nationality', 'name': 'Nat'},
        {'column': 'Sail Number', 'name': 'SailNo'},
        {'column': 'First Name', 'name': 'HelmName'},
        {'column': 'Club Name', 'name': 'Club'},
        {'column': 'Email', 'name': 'HelmEmail'},
        {'column': 'Mobile Phone', 'name': 'HelmPhone'},
        {'column': 'Boat Name', 'name': 'Boat'}]
    cfg.s08_insert_columns = []
    cfg.s09_rewrite_columns = [
        {'column': 'HelmPhone', 'kind': RewriteKind.REGEX_SUBSTITUTE,
         'from': '^', 'to': '+', 'case': CaseSensitivity.MATCH_CASE},
        {'column': 'HelmPhone', 'kind': RewriteKind.REGEX_SUBSTITUTE,
         'from': '^\\+\\+', 'to': '+', 'case': CaseSensitivity.MATCH_CASE}]
    cfg.s10_column_order = ['Class', 'Division', 'Nat', 'SailNo', 'Boat',
                            'HelmName', 'Club', 'HelmEmail', 'HelmPhone']
    cfg.write(to_json_filename=filename)


TXT_R2S_NAME = SYNTAX_R2S_COMMON + BY_NAME_COMMON


def generate_syntax_r2s_num(filename: str, colref: ColumnRef) -> None:
    """Generate config example for rrs_to_sw."""
    assert colref == ColumnRef.BY_NUMBER
    cfg = ConfigXlsListTransfNum()
    set_input_excel(cfg)
    set_output_csv(cfg)
    cfg.s01_split_rows = []
    cfg.s02_merge_rows = []
    cfg.s03_split_columns = []
    cfg.s04_remove_columns = [3, 11]
    cfg.s05_merge_columns = [{'columns': [5, 6], 'separator': ' '}]
    cfg.s06_place_columns_first = []
    cfg.s07_rename_columns = [
        {'column': 0, 'name': 'Class'},
        {'column': 1, 'name': 'Division'},
        {'column': 2, 'name': 'Nat'},
        {'column': 3, 'name': 'SailNo'},
        {'column': 4, 'name': 'Boat'},
        {'column': 5, 'name': 'HelmName'},
        {'column': 6, 'name': 'Club'},
        {'column': 7, 'name': 'HelmEmail'},
        {'column': 8, 'name': 'HelmPhone'}]
    cfg.s08_insert_columns = []
    cfg.s09_rewrite_columns = [
        {'column': 8, 'kind': RewriteKind.REGEX_SUBSTITUTE,
         'from': '^', 'to': '+', 'case': CaseSensitivity.MATCH_CASE},
        {'column': 8, 'kind': RewriteKind.REGEX_SUBSTITUTE,
         'from': '^\\+\\+', 'to': '+', 'case': CaseSensitivity.MATCH_CASE},
        {'column': 8, 'kind': RewriteKind.REGEX_SUBSTITUTE,
         'from': '^\\+H', 'to': 'H', 'case': CaseSensitivity.MATCH_CASE}]
    cfg.write(to_json_filename=filename)


TXT_R2S_NUM = SYNTAX_R2S_COMMON + BY_NUMBER_COMMON


def generate_syntax_example(filename: str, colref: ColumnRef) -> None:
    """Generate config example for example."""
    cfg = config_from_colref(colref)
    set_input_csv(cfg=cfg, encoding='utf_8_sig', delimiter=',')
    set_output_csv(cfg)
    cfg.input_table.csv = TioJsonCsvConfig(dialect=CsvDialect.UNIX)
    cfg.write(to_json_filename=filename)


TXT_SYNTAX_EXAMPLE = '''
This example configuration file does not match any known use case.
It is written as an example configuration with 2 objectives:
to demonstrate all configuration options, and to be small.
'''


def generate_rowsplitmerge(filename: str, colref: ColumnRef) -> None:
    """Generate config example for row_split_merge."""
    cfg = config_from_colref(colref)
    set_input_excel(cfg)
    set_output_excel(cfg)
    cfg.s03_split_columns = []
    cfg.s05_merge_columns = []
    cfg.s07_rename_columns = []
    cfg.s08_insert_columns = []
    cfg.s09_rewrite_columns = []
    if colref == ColumnRef.BY_NUMBER:
        assert isinstance(cfg, ConfigXlsListTransfNum)
        cfg.s01_split_rows = [{'column': 2, 'separators': [' + '],
                               'not_separators': []}]
        cfg.s02_merge_rows = [{'columns': [2], 'separator': ' + '}]
        cfg.s04_remove_columns = []
        cfg.s06_place_columns_first = [2, 1]
    else:
        assert isinstance(cfg, ConfigXlsListTransfName)
        cfg.s01_split_rows = [{'column': 'To', 'separators': [' + '],
                               'not_separators': []}]
        cfg.s02_merge_rows = [{'columns': ['To'], 'separator': ' + '}]
        cfg.s10_column_order = ['To', 'What', 'From']
    cfg.write(to_json_filename=filename)


TXT_ROWSPLITMERGE = '''
Consider that you have a list of things to pick up from stores
for customers in the following format:

| From        | What   | To            |
|-------------|--------|---------------|
| Gardener    | Apples | Jones + Smith |
| Brewery     | Beer   | Smith + Bush  |
| Dairy       | Milk   | Jones         |

But for distributing it to customers you would like to have the
list in the following format:

| To        | What          | From               |
|-----------|---------------| -------------------|
| Jones     | Apples + Milk | Gardener + Dairy   |
| Smith     | Apples + Beer | Gardener + Brewery |
| Bush      | Beer          | Brewery            |

To do this transformation we use the "s01_split_rows" to split
rows based on the "To" column using " + " as separator.
Then we use "s02_merge_rows" to merge the rows that have
identical values in the "To" column.
Finally we specify the new column order.
'''

TXT_ROWSPLITMERGE_NAME = TXT_ROWSPLITMERGE + BY_NAME_COMMON
TXT_ROWSPLITMERGE_NUM = TXT_ROWSPLITMERGE + BY_NUMBER_COMMON


dispatch = {'forms_to_rrs': {ColumnRef.BY_NAME: generate_syntax_o2r_name,
                             ColumnRef.BY_NUMBER: generate_syntax_o2r_num},
            'forms_to_sw': {ColumnRef.BY_NAME: generate_syntax_o2s_name,
                            ColumnRef.BY_NUMBER: generate_syntax_o2s_num},
            'rrs_to_sw': {ColumnRef.BY_NAME: generate_syntax_r2s_name,
                          ColumnRef.BY_NUMBER: generate_syntax_r2s_num},
            'sw_to_rrs': {ColumnRef.BY_NAME: generate_syntax_sw2r_name,
                          ColumnRef.BY_NUMBER: generate_syntax_sw2r_num},
            'sa_to_rrs': {ColumnRef.BY_NAME: generate_syntax_sa2r_name,
                          ColumnRef.BY_NUMBER: generate_syntax_sa2r_num},
            'row_split_merge': {ColumnRef.BY_NAME: generate_rowsplitmerge,
                                ColumnRef.BY_NUMBER: generate_rowsplitmerge},
            'example': {ColumnRef.BY_NAME: generate_syntax_example,
                        ColumnRef.BY_NUMBER: generate_syntax_example}}

txts = {'forms_to_rrs': {ColumnRef.BY_NAME: TXT_O2R_NAME,
                         ColumnRef.BY_NUMBER: TXT_O2R_NUM},
        'forms_to_sw': {ColumnRef.BY_NAME: TXT_O2S_NAME,
                        ColumnRef.BY_NUMBER: TXT_O2S_NUM},
        'rrs_to_sw': {ColumnRef.BY_NAME: TXT_R2S_NAME,
                      ColumnRef.BY_NUMBER: TXT_R2S_NUM},
        'sw_to_rrs': {ColumnRef.BY_NAME: TXT_SW2R_NAME,
                      ColumnRef.BY_NUMBER: TXT_SW2R_NUM},
        'sa_to_rrs': {ColumnRef.BY_NAME: TXT_SA2R_NAME,
                      ColumnRef.BY_NUMBER: TXT_SA2R_NUM},
        'row_split_merge': {ColumnRef.BY_NAME: TXT_ROWSPLITMERGE_NAME,
                            ColumnRef.BY_NUMBER: TXT_ROWSPLITMERGE_NUM},
        'example': {ColumnRef.BY_NAME: TXT_SYNTAX_EXAMPLE,
                    ColumnRef.BY_NUMBER: TXT_SYNTAX_EXAMPLE}}


def get_example_names() -> list[str]:
    """Return names of example usages with example cfg files."""
    return list(dispatch.keys())


def generate_examplecfg(filename: str, cfgtype: str, colref: ColumnRef) -> int:
    """Generate example cfg file for one type of usage."""
    fixedname = fix_file_extension(filename=filename, ext_to_add='.cfg')
    syntaxfilename = fix_file_extension(filename=fixedname, ext_to_add='.txt',
                                        ext_to_remove='.cfg')
    generate_syntax_txt(filename=syntaxfilename,
                        example_description=txts[cfgtype][colref],
                        cfgfilename=fixedname)

    dispatch[cfgtype][colref](filename=fixedname, colref=colref)
    print(f'Wrote files {fixedname} and {syntaxfilename}')
    return 0
