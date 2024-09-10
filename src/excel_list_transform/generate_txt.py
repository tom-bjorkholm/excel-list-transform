#! /usr/local/bin/python3
"""Generate example configuration files."""

# Copyright (c) 2024 Tom Björkholm
# MIT License


def generate_syntax_txt(filename: str, example_description: str,
                        cfgfilename: str) -> None:
    """Generate a text file with configuration syntax description."""
    msg = '''
    Description of how to write/change the configuration file.
    ==========================================================

    The configuration file is in JSON syntax.
    https://en.wikipedia.org/wiki/JSON
    The keywords and the nesting is important. The order of keywords
    have no significance (the examples use alphabetical order).
    Indentation and line breaks hanve no significance.

    It is recommended that you let the command generate a configuration
    file and then edit that file to match your needs. It is NOT recommended
    that the user writes the configuration file from scratch.


    column_ref
    ==========
    The keyword column_ref is used to tell if the configuration
    references columns by name (value "BY_NAME") or by number
    (value "BY_NUMBER"). The syntax of parts of the configuration file
    is slightly different depending on the how columns are referenced.
    Thus, you first have to decide how you want to reference columns.

    Generally it is easier to write a correct configuration using "BY_NAME",
    so "BY_NAME" is recommended when possible.
    "BY_NAME" imposes a few restrictions:
     - each column need to have a unique text at the first line to use
       as column name.
     - the column names need to be stable and known.

    If you cannot meet these restrictions you will have to reference the
    columns by number "BY_NUMBER". When referencing the columns by number
    it is important to note that the column numbers change when splitting,
    merging, inserting or removing columns.

    When referencing "BY_NUMBER" the columns are numbered from left to
    right. The leftmost column is number 0.


    Type of input and output file
    ==============================
    The type of input file to read is determined by "in_type".
    The type of input file to read is determined by "out_type".
    "in_type" and "out_type" can have values "CSV" or "EXCEL".

    Excel files can be read and written using three libraries.
    "in_excel_library" and "out_excel_library" can have values
    "OPENPYXL", "XLSXWRITER" or "PYLIGHTXL". These are different
    third party libraries that can read/write excel. My experience
    is that "PYLIGHTXL" most often is able to read and write excel
    files correctly. If you have trouble reading/writing your
    particular excel file, please try different combinations of
    these libraries. "in_excel_library" is always needed in
    the configuration file but is not used if the input is CSV.
    "out_excel_library" is always needed in the configuration file
    but is not used if the output is CSV.

    Currently "PYLIGHTXL" appears to be best to write files for
    Microsoft Excel to open without complaints. But "OPENPYXL"
    appears to be best to write files to be imported by
    https://www.racingrulesofsailing.org .
    When newer versions arrive this might change...

    Comma separated values files (CSV files) may differ slightly
    depening on the programs used to read/write them and the locale
    used. "in_csv_dialect" and "out_csv_dialect" changes how CSV files
    are read and written. They are always needed in the configuration
    file, but are not used if the input and output are excel files.


    Column manipulation
    ====================
    A number or records starting with "s" and a number describe
    column manipulation to be done. These manipulations are done
    in order of the number: s1_(something) before s2_(something),
    before s3_(something). As the "s" number records might add, remove
    or rename columns, it is important to keep track of the order that
    they are applied.

    Some "s" numbers are only used if columns are referenced
    by name while other "s" numbers are only used if columns are
    referenced by number.


    "s1_split_columns"
    ==================
    The first operation that is done is splitting of columns.
    The key "s1_split_columns" have an array of splits to be
    done. (When the list of splits has more than one split,
    the least confusion is to split columns to the right before
    columns to the left. Named references also helps to avoid
    confusion.)

    Each split has the following keys: "column", "separator" and
    "where". In the case of number column references the there is
    also the key "store_single". In the case of named column references
    there is instead the key "right_name".

    "column" take as argument the column to be split. This is a
    column number in the case of "BY_NUMBER", and a column name/title
    in the sace of "BY_NAME".

    "separator" is a string of characters that indicate the position
    where the column shall be split. (For instance a single space
    if splitting between two words.)

    "where" can either have the value "RIGHTMOST" or the value
    "LEFTMOST". This is used if the separator string is present
    more than once in the column value to split. The split is then
    done at the leftmost or rightmost occurence according to the
    value of "where".

    "right_name" is only used if column references are "BY_NAME".
    After the split the left column uses the original column name,
    and "right_name" is used as the name of the new right column.

    "store_single" is used only if column references are "BY_NUMBER".
    If the column value before split does not include the separator,
    the result of the split is a single value.
    "store_single" can have the value "RIGHTMOST" or the value
    "LEFTMOST". This determines which column the single value shall
    be stored in. (The other column will be empty.)
    In the case of column references "BY_NAME" the single value is
    always stored in the column with the original name.


    "s2_remove_columns"
    ===================
    "s2_remove_columns" is only used with column references "BY_NUMBER".
    The value of "s2_remove_columns" is a list of column numbers to
    remove.
    (For columns references "BY_NAME" see "s8_column_order".)


    "s3_merge_columns"
    ==================
    The key "s3_merge_columns" have an array of merges to be done.
    Each merge have the keys "columns" and "separator"

    "columns" have a list of column references. If "BY_NAME" the
    column references are column names/titles. If "BY_NUMBER" the
    column references are column numbers.

    "separator" is a string of characters that is inserted between
    the column values being merged.


    "s4_place_columns_first"
    ========================
    "s4_place_columns_first" is only used with column references "BY_NUMBER".
    The key "s4_place_columns_first" has a value that is a list of the
    column numbers to be placed first in order. This step re-orders the
    columns.
    (For columns references "BY_NAME" see "s8_column_order".)


    "s5_rename_columns"
    ===================
    The key "s5_rename_columns" has a value that is a list of column rename
    operations. Each column rename operation has the keys "column" and "name"

    "column" is the number/name of the column before renameing. This is a
    column number in the case of "BY_NUMBER", and a column name/title
    in the sace of "BY_NAME".

    "name" is the new name/title of the column identified by "column".


    "s6_insert_columns"
    ===================
    The key "s6_insert_columns" has a value that is a list of columns to
    insert. Each column to insert is described by the keys: "column",
    "value" and for "BY_NUMBER" only "name",

    "column" is the column reference of the inserted column.
    In the case of "BY_NUMBER" this is also the position where
    the column is inserted. In the case of "BY_NAME" this is the
    name of the inserted column.

    "value" is the value that this column shall have for every row.
    The special word null can be used to state that column shall be
    empty (with no value).

    "name" is the name/title of the column in the case of "BY_NUMBER".


    "s7_rewrite_columns"
    ====================
    The key "s7_rewrite_columns" has a value that is a list of
    rewrite operations that will be applied in order. Each rewrite
    operation is described by several keys.

    "kind" key is used to specify which kind of rewrite operation
    this is. Different values for "kind" will have different other
    keys. "kind" can have the values: "STRIP", "REMOVECHARS",
    "STR_SUBSTITUTE" and "REGEX_SUBSTITUTE".

    "kind" value "STRIP" shall have the keys: "kind", "case", "chars"
    and "column". This case means that the characters in "chars" shall
    be stripped off the beginning and end of the column value. If key
    "chars" is empty string, then white space is stripped off.

    "kind" value "REMOVECHARS" shall have the keys: "kind", "case",
    "chars" and "column". This case means that the characters in "chars"
    shall be removed from the column (also when the characters are in the
    middle of the column value).

    "kind" value "REGEX_SUBSTITUTE" shall have the keys: "kind", "case",
    "from", "to" and "column". This case means that "from" value is used
    as a regular expression, and if it matches the matching part is replaced
    with the value of key "to".

    "kind" value "STR_SUBSTITUTE" shall have the keys: "kind", "case",
    "from", "to" and "column". This case means that "from" value is a string,
    and if a sub-string of the column value equals this string that part is
    replaced with the value of key "to".

    "column" identifies the column to rewrite. This is a column number in the
    case of "BY_NUMBER", and a column name/title in the sace of "BY_NAME".

    "chars" is a list of characters written as a string. These characters
    are removed or stripped off the beginning/end depending on the "kind"
    value.

    "case" specifies if comparison of column value to value of "chars" or
    value of "from" shall be case sensitive or not. Possible values are
    "MATCH_CASE" and "IGNORE_CASE".

    "from" specifies what the part to substitute shall match. The from part
    in the substitute from something to something.

    "to" specifies the string that substitution will replace "from" with.


    "s8_column_order"
    =================
    The key "s8_column_order" is used only in the "BY_NAME" case.
    The value of the "s8_column_order" key is a list of column names.
    The columns will be output in this order.
    Columns not mentioned in "s8_column_order" will not be output,
    and will thus be removed.
    (For "BY_NUMBER" see "s2_remove_columns" and "s4_place_columns_first".)
    '''
    with open(file=filename, mode='w', encoding='utf-8') as file:
        print(f'Explanation for example configuration file {cfgfilename}',
              file=file)
        print(example_description, file=file)
        print('\n\n\n', file=file)
        print(msg, file=file)
