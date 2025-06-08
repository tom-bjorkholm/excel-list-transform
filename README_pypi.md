# excel-list-transform

## Background

This python application was born out of an experience at sail racing events. At the start of the events we received an excel list with participants (from a registration web) to enter into the scoring software and into tool for online notice board. The information was present in the excel file, but the columns were all wrong. To avoid the stressful manual rework of the information in the excel list this application was created.

## What it does

This small python application:

* reads data (that is a list with columns) from an excel file or from a comma separate values (CSV) file.
* splits columns in the list (like creating "first name" and "last name" columns from "name" column)
* merges columns in the list (like creating "name" column from "first name" and "last name" columns)
* removes columns in the list
* reorders columns in the list
* renames columns in the list
* inserts columns in the list
* rewrites columns in the list (like transforming telephone numbers from local/national format to international format)
* writes the resulting data (that is a list with columns) to an excel file or to a comma separate values (CSV) file.

How this is done is governed by a configuration file. The application can create a number of example configuration files with accompanying description text files.

## Installing it

If you want to use it, install it using pip. A precondition is that you have Python installed on you computer. See version table below for information on needed Python version.
Python can be downloaded from [https://www.python.org/downloads/](https://www.python.org/downloads/).

### Installing on mac and Linux

````sh
pip3 install --upgrade excel-list-transform
````

### Installing on Microsoft Windows

````sh
pip install --upgrade excel-list-transform
````

## Version history

| Version | Date        | Python version  | Description                         |
|---------|-------------|-----------------|-------------------------------------|
| 0.5     | 10 Sep 2024 | 3.10.5 or newer | First released version              |
| 0.6     | 22 Sep 2024 | 3.12.6 or newer | Configured encoding for CSV         |
| 0.6.2   | 28 Sep 2024 | 3.12.6 or newer | Fix name in test report             |
| 0.7.0   | 03 Nov 2024 | 3.12.6 or newer | Option to strip space in excel      |
| 0.7.1   | 15 Dec 2024 | 3.12.6 or newer | Refactor tests                      |
| 0.7.3   | 06 Jan 2025 | 3.12.6 or newer | Example config for Sailwave to RRS  |
| 0.7.5   | 21 Mar 2025 | 3.12.6 or newer | Example config for Sailarena to RRS |
| 0.7.6   | 21 Mar 2025 | 3.13.2 or newer | Adapted to Python 3.13.2            |
| 0.7.7   | 09 Apr 2025 | 3.13.3 or newer | refactor phone number rewrite cfg   |
| 0.7.9   | 09 Jun 2025 | 3.10.x          | Backport 0.7.12 to Python 3.10      |
| 0.7.10  | 09 Jun 2025 | 3.11.x          | Backport 0.7.12 to Python 3.11      |
| 0.7.11  | 09 Jun 2025 | 3.12.x          | Backport 0.7.12 to Python 3.12      |
| 0.7.12  | 09 Jun 2025 | 3.13 or newer   | Add version sub-command             |

## Running the application

### Running the application on mac and Linux

````sh
python3 -m excel_list_transform --help
python3 -m excel_list_transform version
python3 -m excel_list_transform cfg-example --help
python3 -m excel_list_transform transform --help
python3 -m excel_list_transform cfg-example -k forms_to_rrs -r by_name -o example.cfg
python3 -m excel_list_transform transform -c example.cfg -i input.xlsx -o output.xlsx
````

### Running the application on Microsoft Windows

````sh
python -m excel_list_transform --help
python -m excel_list_transform version
python -m excel_list_transform cfg-example --help
python -m excel_list_transform transform --help
python -m excel_list_transform cfg-example -k forms_to_rrs -r by_name -o example.cfg
python -m excel_list_transform transform -c example.cfg -i input.xlsx -o output.xlsx
````

## Suggested way to get started

 1. Use the "cfg-example" sub-command to generate a few example configuration (.cfg) files with description (.txt) files.
 2. Read the example configuration (.cfg) files and the accompanying description (.txt) files.
 3. Find an example that is close to what you want to achieve.
 4. Modify that configuration file to achieve what you want to achieve.
 5. Use the "transform" sub-command to read your data and output it transformed or reorganized according to your modified configuration file.
 6. Read the produced output. If necessary go back to step 4 and adjust how the data is transformed.

### Example configuration files

When using the "cfg-example" sub-command to generate an example configuration file (say example.cfg) a text file describing the configuration and the syntax of the configuration file is also generated. If the example configuration file is named example.cfg, then the text file descriging the configuration is named example.txt.

You can generate several example configuration files each with an accompanying text file descriping it.

Read the text file describing the configuration file while looking at the configuration file to understand the syntax and the possible options.

## Performance

This application is written for the moderate amounts of data when registering participants for the majority of sports events. If you have millions of rows this application is not for you.

With an input file consisting of 20 columns producing an output file of 11 columns I have measured the following performance on a MacBook Air (laptop) from 2020:

* 1 000 rows processed in less than 0.5 seconds
* 10 000 rows processed in less than 4 seconds
* 80 000 rows processed in less than 30 seconds
* 120 000 rows processed in 40 seconds

Naturally your performance will be different based on computer hardware, operating system and Python version. Generally it should be reasonably fast for less than 10 000 rows, painfully slow but somewhat usable up to 100 000 rows and probably so slow that it is unusable for more than 100 000 rows.

## Description of how to write/change the configuration file

The configuration file is in JSON syntax. [https://en.wikipedia.org/wiki/JSON](https://en.wikipedia.org/wiki/JSON)

The keywords and the nesting is important. The order of keywords
have no significance (the examples use alphabetical order).
Indentation and line breaks have no significance.

The encoding for the configuration file must be UTF-8. (US-ACII is a subset of UTF-8.)

It is recommended that you let the command generate a configuration
file and then edit that file to match your needs. It is NOT recommended
that the user writes the configuration file from scratch.

## column_ref

The keyword **column_ref** is used to tell if the configuration
references columns by name (value *"BY_NAME"*) or by number
(value *"BY_NUMBER"*). The syntax of parts of the configuration file
is slightly different depending on the how columns are referenced.
Thus, you first have to decide how you want to reference columns.

Generally it is easier to write a correct configuration using *"BY_NAME"*,
so *"BY_NAME"* is recommended when possible. *"BY_NAME"* imposes a few restrictions:

* each column need to have a unique text at the first line to use as column name.
* the column names need to be stable and known.

If you cannot meet these restrictions you will have to reference the
columns by number *"BY_NUMBER"*. When referencing the columns by number
it is important to note that the column numbers change when splitting,
merging, inserting or removing columns.

When referencing *"BY_NUMBER"* the columns are numbered from left to
right. The leftmost column is number 0.

## Type of input and output file

The type of input file to read is determined by **"in_type"**.
The type of output file to write is determined by **"out_type"**.
**"in_type"** and **"out_type"** can have values *"CSV"* or *"EXCEL"*.

Excel files can be read and written using three libraries.
**"in_excel_library"** and **"out_excel_library"** can have values
*"OPENPYXL"*, *"XLSXWRITER"* or *"PYLIGHTXL"*. These are different
third party libraries that can read/write excel. My experience
is that *"PYLIGHTXL"* most often is able to read and write excel
files correctly. If you have trouble reading/writing your
particular excel file, please try different combinations of
these libraries. **"in_excel_library"** is always needed in
the configuration file but is not used if the input is CSV.
**"out_excel_library"** is always needed in the configuration file
but is not used if the output is CSV.

Currently *"PYLIGHTXL"* appears to be best to write files for
Microsoft Excel to open without complaints. But *"OPENPYXL"*
appears to be best to write files to be imported by
[https://www.racingrulesofsailing.org](https://www.racingrulesofsailing.org).
When newer versions arrive this might change...

Comma separated values files (CSV files) may differ slightly
depening on the programs used to read/write them and the locale
used. **"in_csv_dialect"** and **"out_csv_dialect"** changes how CSV files
are read and written. They are always needed in the configuration
file, but are not used if the input and output are excel files.

Comma separated values files might have different encoding for
the text in the file. [https://en.wikipedia.org/wiki/Character_encoding](https://en.wikipedia.org/wiki/Character_encoding)
This is specified with **"in_csv_encoding"** and **"out_csv_encoding"**.
Unless you know that you need another encoding leave these as
in the generated example configuration. (In version 0.5 these
configuration parameters are missing. To be compatible with 0.5
configuration files "in_csv_encoding" defaults to "utf_8_sig"
and "out_csv_encoding" defaults to "utf-8" if missing in the
configuration file.)

## Extra spaces in excel input files

When viewing an excel file in excel it is very hard to notice if
some string value in a cell has trailing white space. These trailing
trailing spaces can make the further processing of a file difficult
as the strings in the file are not what you thought they are.
(This especially is a problem if you have trailing spaces in a
cell on the first line, and refer to columns by their names.)

The configuration **"in_excel_col_name_strip"** can be set to true,
to strip off leading and trailing whitespace from all columns
values read from the first line of the excel input file.
The configuration **"in_excel_values_strip"** can be set to true,
to strip off leading and trailing whitespace from all columns
values read from from the other lines (not the first line) of
the excel input file. (In version 0.6.2 and earlier these are
missing. To be compatible with version 0.6.2 and earlier
**"in_excel_col_name_strip"** and  **"in_excel_values_strip"**
defaults to false if missing in the configuration file.)

## Column manipulation

A number or records starting with "s" and a number describe
column manipulation to be done. These manipulations are done
in order of the number: s1_(something) before s2_(something),
before s3_(something). As the "s" number records might add, remove
or rename columns, it is important to keep track of the order that
they are applied.

Some "s" numbers are only used if columns are referenced
by name while other "s" numbers are only used if columns are
referenced by number.

### "s1_split_columns"

The first operation that is done is splitting of columns.
The key **"s1_split_columns"** have an array of splits to be
done. (When the list of splits has more than one split,
the least confusion is to split columns to the right before
columns to the left. Named references also helps to avoid
confusion.)

Each split has the following keys: **"column"**, **"separator"** and
**"where"**. In the case of number column references the there is
also the key **"store_single"**. In the case of named column references
there is instead the key **"right_name"**.

**"column"** take as argument the column to be split. This is a
column number in the case of *"BY_NUMBER"*, and a column name/title
in the sace of *"BY_NAME"*.

**"separator"** is a string of characters that indicate the position
where the column shall be split. (For instance a single space
if splitting between two words.)

**"where"** can either have the value *"RIGHTMOST"* or the value
*"LEFTMOST"*. This is used if the separator string is present
more than once in the column value to split. The split is then
done at the leftmost or rightmost occurence according to the
value of **"where"**.

**"right_name"** is only used if column references are *"BY_NAME"*.
After the split the left column uses the original column name,
and **"right_name"** is used as the name of the new right column.

**"store_single"** is used only if column references are *"BY_NUMBER"*.
If the column value before split does not include the separator,
the result of the split is a single value.
**"store_single"** can have the value *"RIGHTMOST"* or the value
*"LEFTMOST"*. This determines which column the single value shall
be stored in. (The other column will be empty.)
In the case of column references *"BY_NAME"* the single value is
always stored in the column with the original name.

### "s2_remove_columns"

**"s2_remove_columns"** is only used with column references *"BY_NUMBER"*.
The value of **"s2_remove_columns"** is a list of column numbers to
remove. (For columns references *"BY_NAME"* see **"s8_column_order"**.)

### "s3_merge_columns"

The key **"s3_merge_columns"** have an array of merges to be done.
Each merge have the keys **"columns"** and **"separator"**.

**"columns"** have a list of column references. If *"BY_NAME"* the
column references are column names/titles. If *"BY_NUMBER"* the
column references are column numbers.

**"separator"** is a string of characters that is inserted between
the column values being merged.

### "s4_place_columns_first"

**"s4_place_columns_first"** is only used with column references *"BY_NUMBER"*.
The key **"s4_place_columns_first"** has a value that is a list of the
column numbers to be placed first in order. This step re-orders the
columns.
(For columns references *"BY_NAME"* see **"s8_column_order"**.)

### "s5_rename_columns"

The key **"s5_rename_columns"** has a value that is a list of column rename
operations. Each column rename operation has the keys **"column"** and **"name"**.

**"column"** is the number/name of the column before renameing. This is a
column number in the case of *"BY_NUMBER"*, and a column name/title
in the sace of *"BY_NAME"*.

**"name"** is the new name/title of the column identified by **"column"**.

### "s6_insert_columns"

The key **"s6_insert_columns"** has a value that is a list of columns to
insert. Each column to insert is described by the keys: **"column"**,
**"value"** and for *"BY_NUMBER"* only also **"name"**.

**"column"** is the column reference of the inserted column.
In the case of *"BY_NUMBER"* this is also the position where
the column is inserted. In the case of *"BY_NAME"* this is the
name of the inserted column.

**"value"** is the value that this column shall have for every row.
The special word *null* can be used to state that column shall be
empty (with no value).

**"name"** is the name/title of the column in the case of *"BY_NUMBER"*.

### "s7_rewrite_columns"

The key **"s7_rewrite_columns"** has a value that is a list of
rewrite operations that will be applied in order. Each rewrite
operation is described by several keys.

**"kind"** key is used to specify which kind of rewrite operation
this is. Different values for **"kind"** will have different other
keys. **"kind"** can have the values: *"STRIP"*, *"REMOVECHARS"*,
*"STR_SUBSTITUTE"* and *"REGEX_SUBSTITUTE"*.

**"kind"** value *"STRIP"* shall have the keys: **"kind"**, **"case"**, **"chars"**
and **"column"**. This case means that the characters in **"chars"** shall
be stripped off the beginning and end of the column value. If key
**"chars"** is empty string, then white space is stripped off.

**"kind"** value *"REMOVECHARS"* shall have the keys: **"kind"**, **"case"**,
**"chars"** and **"column"**. This case means that the characters in **"chars"**
shall be removed from the column (also when the characters are in the
middle of the column value).

**"kind"** value *"REGEX_SUBSTITUTE"* shall have the keys: **"kind"**, **"case"**,
**"from"**, **"to"** and **"column"**. This case means that **"from"** value is used
as a regular expression, and if it matches the matching part is replaced
with the value of key **"to"**.

**"kind"** value *"STR_SUBSTITUTE"* shall have the keys: **"kind"**, **"case"**,
**"from"**, **"to"** and **"column"**. This case means that **"from"** value is a string,
and if a sub-string of the column value equals this string that part is
replaced with the value of key **"to"**.

**"column"** identifies the column to rewrite. This is a column number in the
case of *"BY_NUMBER"*, and a column name/title in the sace of *"BY_NAME"*.

**"chars"** is a list of characters written as a string. These characters
are removed or stripped off the beginning/end depending on the **"kind"**
value.

**"case"** specifies if comparison of column value to value of **"chars"** or
value of **"from"** shall be case sensitive or not. Possible values are
*"MATCH_CASE"* and *"IGNORE_CASE"*.

**"from"** specifies what the part of the column value shall be to substituted
when matching. The from part in the substitute from something to something.

**"to"** specifies the string that substitution will replace **"from"** with.

### "s8_column_order"

The key **"s8_column_order"** is used only in the *"BY_NAME"* case.
The value of the **"s8_column_order"** key is a list of column names.
The columns will be output in this order.
Columns not mentioned in **"s8_column_order"** will not be output,
and will thus be removed.
(For *"BY_NUMBER"* see **"s2_remove_columns"** and **"s4_place_columns_first"**.)

## Source code

Source code and tests are available at [https://bitbucket.org/tom-bjorkholm/excel-list-transform](https://bitbucket.org/tom-bjorkholm/excel-list-transform).
