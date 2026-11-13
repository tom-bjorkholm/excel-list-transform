#! /usr/local/bin/python3
"""What each configuration member means, for the configuration editor.

The editor shows the type of a member next to its value, but it cannot
work out the types of the transform rules (lists of dicts whose keys
depend on the rule). These descriptions therefore also say which keys a
rule has and which values each key takes.
"""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from edit_cfg_json import Descriptions
from tableio_cfg_json import tio_json_descriptions
from excel_list_transform.config_enums import ColumnRef

TABLE_MEMBERS = ('input_table', 'output_table')
"""Members holding a TableIO configuration, described by tableio-cfg-json."""

COMMON_TEXTS: Descriptions = {
    ('column_ref',):
        'How this file refers to columns: BY_NAME (by the title in the '
        'first row) or BY_NUMBER (by position, the leftmost column is 0). '
        'The rest of the file is written for this choice, so it cannot be '
        'changed here: a save with another value is refused. To switch, '
        'generate a new file with the "cfg-example" sub-command and its '
        '"-r" option.',
    ('max_column_read',):
        'How many columns to read from the input file, counted from the '
        'left. Columns further to the right are ignored. A whole number.',
    ('input_table',):
        'The input file: which file format to read and how. Normally only '
        '"format_name" needs to be set, for example "Excel" or "CSV". Leave '
        '"implementation" empty (null) to let the program choose one.',
    ('output_table',):
        'The output file: which file format to write and how. Normally only '
        '"format_name" needs to be set, for example "Excel" or "CSV". Leave '
        '"implementation" empty (null) to let the program choose one.',
    ('in_excel_col_name_strip',):
        'true: remove spaces at the start and end of the column titles (the '
        'first row) of an Excel input file. Such spaces cannot be seen in '
        'Excel, but make a title differ from the name written in this file. '
        'Used only when the input format is Excel.',
    ('in_excel_values_strip',):
        'true: remove spaces at the start and end of all values below the '
        'first row of an Excel input file. Used only when the input format '
        'is Excel.',
    ('output_borders',):
        'true: ask for borders around the cells of the output table. Output '
        'formats that cannot show borders ignore this.',
    ('output_filtered_table',):
        'true: make the output a filtered table, with filter buttons in the '
        'title row. Output formats that cannot do this ignore it.',
    ('s01_split_rows',):
        'Split one row into several rows. Done first of all steps. A list of '
        'rules with the keys "column", "separators", "not_separators" and '
        '"empty_parts". '
        'The value of "column" is cut at every separator, and each part gets '
        'a row of its own. The other columns are copied to every new row. '
        'An empty list splits no rows.',
    ('s01_split_rows', '['):
        'One row split rule, with the keys "column", "separators", '
        '"not_separators" and "empty_parts".',
    ('s01_split_rows', '[', 'separators'):
        'List of texts (at least one) that separate the parts of the value, '
        'for example " + ". Where several separators match at the same '
        'place, the longest one is used. What happens to empty parts is '
        'set by "empty_parts".',
    ('s01_split_rows', '[', 'separators', '['):
        'One non-empty text that separates two parts of the value.',
    ('s01_split_rows', '[', 'not_separators'):
        'List of texts that are not separators even though they contain a '
        'separator, for example "\\ " when " " is a separator. Each of them '
        'must contain a separator and must not be one. A not-separator '
        'takes precedence over a separator at the same place. May be '
        'empty.',
    ('s01_split_rows', '[', 'not_separators', '['):
        'One text that is not a separator.',
    ('s01_split_rows', '[', 'empty_parts'):
        'What to do with empty parts (parts that are empty or only '
        'whitespace, from a separator at the start or end, or from two '
        'adjacent separators). One of:\n'
        'KEEP_ALL: keep all of them. "a,,b," split at "," gives "a", "", '
        '"b" and "".\n'
        'DROP_ALL: drop all of them. "a,,b," gives "a" and "b".\n'
        'DROP_TRAILING: drop only those at the end. "a,,b," gives "a", "" '
        'and "b".\n'
        'If all parts are dropped, the row is kept once with the value "".',
    ('s02_merge_rows',):
        'Merge rows that have equal values in some columns into one row. '
        'Done after s01_split_rows. A list of rules with the keys "columns" '
        'and "separator". In every other column the different values of the '
        'merged rows are joined with "separator", and a value that is the '
        'same in all of them is written once. Empty values are left out. '
        'An empty list merges no rows.',
    ('s02_merge_rows', '['):
        'One row merge rule, with the keys "columns" and "separator".',
    ('s02_merge_rows', '[', 'separator'):
        'Text put between the different values of a column when rows are '
        'merged, for example " + ".',
    ('s03_split_columns', '[', 'separator'):
        'Text at which the value is split, for example " " to split between '
        'two words. The separator itself is not kept. Must not be empty.',
    ('s03_split_columns', '[', 'where'):
        'Where to split when the separator occurs more than once: LEFTMOST '
        '(at the first occurrence) or RIGHTMOST (at the last). Splitting '
        '"Anna Maria Smith" at " " gives "Anna" and "Maria Smith" with '
        'LEFTMOST, and "Anna Maria" and "Smith" with RIGHTMOST.',
    ('s03_split_columns', '[', 'end_separators'):
        'Separators to remove at the ends of the value before it is split, '
        'together with any whitespace between them. One of:\n'
        'KEEP: keep the value as it is.\n'
        'STRIP_LEADING: remove separators at the start.\n'
        'STRIP_TRAILING: remove separators at the end. "a,b, ," split at '
        '"," becomes "a,b" before the split.\n'
        'STRIP_BOTH_ENDS: remove separators at the start and the end.\n'
        'A value with only separators and whitespace becomes "".',
    ('s05_merge_columns', '['):
        'One column merge rule, with the keys "columns" and "separator".',
    ('s05_merge_columns', '[', 'separator'):
        'Text put between the joined values, for example " " to join a '
        'first name and a last name.',
    ('s07_rename_columns', '['):
        'One rename rule, with the keys "column" and "name".',
    ('s07_rename_columns', '[', 'name'):
        'New name (title) of the column.',
    ('s08_insert_columns', '[', 'value'):
        'Value of the new column in every row: a text, or null (None) for '
        'an empty column.',
    ('s09_rewrite_columns', '['):
        'One rewrite rule. Its keys are "kind", "column", "case", and '
        'either "chars" (for STRIP and REMOVECHARS) or "from" and "to" (for '
        'STR_SUBSTITUTE and REGEX_SUBSTITUTE).',
    ('s09_rewrite_columns', '[', 'kind'):
        'What the rule does. One of:\n'
        'STRIP: remove the characters of "chars" from the start and end of '
        'the value.\n'
        'REMOVECHARS: remove each text of "chars" wherever it occurs.\n'
        'STR_SUBSTITUTE: replace the text "from" with "to" wherever it '
        'occurs.\n'
        'REGEX_SUBSTITUTE: replace what the regular expression "from" '
        'matches with "to".\n'
        'Changing the kind also means changing the other keys of the rule.',
    ('s09_rewrite_columns', '[', 'case'):
        'Whether upper and lower case letters differ when looking for '
        '"chars" or "from": MATCH_CASE or IGNORE_CASE.',
    ('s09_rewrite_columns', '[', 'chars'):
        'For STRIP: one text. Each of its characters is removed from the '
        'start and end of the value. An empty text removes spaces and other '
        'whitespace.\n'
        'For REMOVECHARS: a list of texts. Each of them is removed wherever '
        'it occurs in the value.',
    ('s09_rewrite_columns', '[', 'chars', '['):
        'One text to remove wherever it occurs in the value.',
    ('s09_rewrite_columns', '[', 'from'):
        'For STR_SUBSTITUTE: the text to look for.\n'
        'For REGEX_SUBSTITUTE: a Python regular expression, for example '
        '"^07" for a value that starts with 07.',
    ('s09_rewrite_columns', '[', 'to'):
        'The text that replaces what "from" found. For REGEX_SUBSTITUTE it '
        'may refer to a group of the expression, such as \\1.'}
"""Descriptions that are the same for both kinds of column references."""

NAME_TEXTS: Descriptions = {
    ('s01_split_rows', '[', 'column'):
        'Name of the column whose value is split into rows.',
    ('s02_merge_rows', '[', 'columns'):
        'Names of the columns (at least one) that must hold equal values for '
        'rows to be merged.',
    ('s02_merge_rows', '[', 'columns', '['):
        'Name of one column that must hold equal values.',
    ('s03_split_columns',):
        'Split a column into two columns. A list of rules with the keys '
        '"column", "separator", "where", "end_separators" and "right_name". '
        'The left part stays in the column and the right part goes into a '
        'new column named "right_name". A value without the separator stays '
        'whole in the column, and the new column is left empty. An empty '
        'list splits no columns.',
    ('s03_split_columns', '['):
        'One column split rule, with the keys "column", "separator", '
        '"where", "end_separators" and "right_name".',
    ('s03_split_columns', '[', 'column'):
        'Name of the column to split. After the split it holds the left '
        'part.',
    ('s03_split_columns', '[', 'right_name'):
        'Name of the new column that gets the right part. Use a name that is '
        'not already in the data: a column with this name is overwritten.',
    ('s05_merge_columns',):
        'Merge several columns into one. Done after s03_split_columns. A '
        'list of rules with the keys "columns" and "separator". The values '
        'of the listed columns are joined with "separator" (empty values '
        'are left out) and stored in the first listed column. The other '
        'listed columns are removed. An empty list merges no columns.',
    ('s05_merge_columns', '[', 'columns'):
        'Names of the columns to merge, in the order their values are '
        'joined. The result is stored under the first name.',
    ('s05_merge_columns', '[', 'columns', '['):
        'Name of one column to merge.',
    ('s07_rename_columns',):
        'Rename columns. Done after the splits and merges, so a column made '
        'by them can be renamed here. A list of rules with the keys "column" '
        'and "name". Each column may be renamed only once. Renames are done '
        'in the listed order, and a rename must not give a column the name '
        'of an existing column.',
    ('s07_rename_columns', '[', 'column'):
        'Name of the column before it is renamed.',
    ('s08_insert_columns',):
        'Add new columns that have the same value in every row. Done after '
        's07_rename_columns. A list of rules with the keys "column" and '
        '"value". Where a new column appears in the output is decided by '
        's10_column_order, as for every other column.',
    ('s08_insert_columns', '['):
        'One insert rule, with the keys "column" and "value".',
    ('s08_insert_columns', '[', 'column'):
        'Name of the new column. It must not already exist in the data.',
    ('s09_rewrite_columns',):
        'Change the values of columns. Done after all other steps, using '
        'the column names as they are after renaming and inserting. A list '
        'of rules, applied in the listed order. The key "kind" decides what '
        'a rule does and which other keys it has. Empty cells are left '
        'unchanged.',
    ('s09_rewrite_columns', '[', 'column'):
        'Name of the column to rewrite.',
    ('s10_column_order',):
        'Which columns the output file gets, and in which order: a list of '
        'column names as they are after all other steps. Columns that are '
        'not listed are left out of the output. Each name may be listed '
        'only once, and must be the name of an existing column.',
    ('s10_column_order', '['):
        'Name of one output column. Its place in this list is its place in '
        'the output.'}
"""Descriptions for a configuration that references columns by name."""

NUMBER_TEXTS: Descriptions = {
    ('s01_split_rows', '[', 'column'):
        'Number of the column whose value is split into rows. Columns are '
        'numbered from the left, starting at 0.',
    ('s02_merge_rows', '[', 'columns'):
        'Numbers of the columns (at least one) that must hold equal values '
        'for rows to be merged.',
    ('s02_merge_rows', '[', 'columns', '['):
        'Number of one column that must hold equal values.',
    ('s03_split_columns',):
        'Split a column into two columns. A list of rules with the keys '
        '"column", "separator", "where", "end_separators" and '
        '"store_single". The left part replaces the column and the right '
        'part becomes a new column just to the right of it, so every column '
        'further right gets a number one higher. The column numbers of all '
        'rules are those of the input file, as the rules are applied from '
        'the rightmost column to the leftmost. Each column may be split only '
        'once. An empty list splits no columns.',
    ('s03_split_columns', '['):
        'One column split rule, with the keys "column", "separator", '
        '"where", "end_separators" and "store_single".',
    ('s03_split_columns', '[', 'column'):
        'Number of the column to split. After the split it holds the left '
        'part, and the column after it holds the right part.',
    ('s03_split_columns', '[', 'store_single'):
        'Which of the two columns gets a value that does not contain the '
        'separator: LEFTMOST (the left column) or RIGHTMOST (the right '
        'column). The other column is left empty.',
    ('s04_remove_columns',):
        'Remove columns. Done after s03_split_columns, so the numbers are '
        'those after the splits. A list of column numbers, each listed only '
        'once. An empty list removes no columns.',
    ('s04_remove_columns', '['):
        'Number of one column to remove.',
    ('s05_merge_columns',):
        'Merge several columns into one. Done after s04_remove_columns, so '
        'the numbers are those after removing columns. A list of rules with '
        'the keys "columns" and "separator". The values of the listed '
        'columns are joined with "separator" (empty values are left out) '
        'and stored in the first listed column. The other listed columns '
        'are removed. Rules may not overlap: every number of one rule must '
        'be smaller than every number of the next rule. An empty list '
        'merges no columns.',
    ('s05_merge_columns', '[', 'columns'):
        'Numbers of the columns to merge, in increasing order. The result is '
        'stored in the first (leftmost) of them.',
    ('s05_merge_columns', '[', 'columns', '['):
        'Number of one column to merge.',
    ('s06_place_columns_first',):
        'Move columns to the left edge. Done after s05_merge_columns, so the '
        'numbers are those after merging columns. A list of column numbers, '
        'each listed only once. The listed columns come first, in the '
        'listed order, followed by all other columns in their current '
        'order. An empty list moves no columns.',
    ('s06_place_columns_first', '['):
        'Number of one column to move. Its place in this list is its new '
        'place.',
    ('s07_rename_columns',):
        'Rename columns, that is change their titles in the first row. Done '
        'after s06_place_columns_first, so the numbers are those after '
        'moving columns. A list of rules with the keys "column" and "name". '
        'Each column may be renamed only once.',
    ('s07_rename_columns', '[', 'column'):
        'Number of the column to rename.',
    ('s08_insert_columns',):
        'Add new columns that have the same value in every row. Done after '
        's07_rename_columns. A list of rules with the keys "column", "name" '
        'and "value". The rules are applied in order of increasing "column", '
        'and every column to the right of an inserted column gets a number '
        'one higher.',
    ('s08_insert_columns', '['):
        'One insert rule, with the keys "column", "name" and "value".',
    ('s08_insert_columns', '[', 'column'):
        'Position of the new column, counting the columns inserted by the '
        'rules before it. It may be one more than the last column number, '
        'to add the column at the right edge.',
    ('s08_insert_columns', '[', 'name'):
        'Name (title) of the new column, written in the first row.',
    ('s09_rewrite_columns',):
        'Change the values of columns. Done after all other steps, so the '
        'numbers are those after inserting columns. A list of rules, applied '
        'in the listed order. The key "kind" decides what a rule does and '
        'which other keys it has. Empty cells are left unchanged.',
    ('s09_rewrite_columns', '[', 'column'):
        'Number of the column to rewrite.'}
"""Descriptions for a configuration that references columns by number."""


def cfg_descriptions(column_ref: ColumnRef) -> Descriptions:
    """Return what every member means, for one kind of column reference.

    Args:
        column_ref: How the edited configuration references columns.

    Returns:
        Descriptions of the members of this application and of the TableIO
        members of its input and output tables.
    """
    own = NAME_TEXTS if column_ref == ColumnRef.BY_NAME else NUMBER_TEXTS
    described = {**COMMON_TEXTS, **own}
    for member in TABLE_MEMBERS:
        described.update(tio_json_descriptions((member,)))
    return described
