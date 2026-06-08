# excel-list-transform

## Background

This python application was born out of an experience at sail racing events.
At the start of the events we received an excel list with participants (from
a registration web) to enter into the scoring software and into tool for
online notice board. The information was present in the excel file, but the
columns were all wrong. To avoid the stressful manual rework of the
information in the excel list this application was created.

## What it does

This small python application:

- reads data (that is a list with columns) from an excel file, an ODS file,
  or from a comma separated values (CSV) file.
- split rows in the list (creating separate rows when a cell contains
  something that can be split to several values, see below)
- merge rows (if specified columns have the same value in several rows
  the rows are merged to one row)
- splits columns in the list (like creating "first name" and "last name"
  columns from "name" column)
- merges columns in the list (like creating "name" column from "first name"
  and "last name" columns)
- removes columns in the list
- reorders columns in the list
- renames columns in the list
- inserts columns in the list
- rewrites columns in the list (like transforming telephone numbers from
  local/national format to international format)
- writes the resulting data (that is a list with columns) to an excel file,
  an ODS spreadsheet file, a comma separated values (CSV) file, or
  any of a number of other file formats.

How this is done is governed by a configuration file. The application can
create a number of example configuration files with accompanying description
text files.

## Installing it

If you want to use it, install it using pip. A precondition is that you have
Python installed on your computer. See version table below for information
on needed Python version. Python can be downloaded
from [https://www.python.org/downloads/](https://www.python.org/downloads/).

### Installing on mac and Linux

````sh
pip3 install --upgrade excel-list-transform
````

### Installing on Microsoft Windows

````sh
pip install --upgrade excel-list-transform
````

## Running the application

### Running the application on mac and Linux

````sh
python3 -m excel_list_transform --help
python3 -m excel_list_transform version
python3 -m excel_list_transform cfg-example --help
python3 -m excel_list_transform transform --help
python3 -m excel_list_transform cfg-example -k forms_to_rrs -r by_name -o example.cfg
python3 -m excel_list_transform transform -c example.cfg -i input.xlsx -o output.xlsx
python3 -m excel_list_transform migrate-cfg --help
python3 -m excel_list_transform migrate-cfg -i old.cfg -o new.cfg
````

### Running the application on Microsoft Windows

````sh
python -m excel_list_transform --help
python -m excel_list_transform version
python -m excel_list_transform cfg-example --help
python -m excel_list_transform transform --help
python -m excel_list_transform cfg-example -k forms_to_rrs -r by_name -o example.cfg
python -m excel_list_transform transform -c example.cfg -i input.xlsx -o output.xlsx
python -m excel_list_transform migrate-cfg --help
python -m excel_list_transform migrate-cfg -i old.cfg -o new.cfg
````

## Suggested way to get started

 1. Use the "cfg-example" sub-command to generate a few example configuration
    (.cfg) files with description (.txt) files.
 2. Read the example configuration (.cfg) files and the accompanying
    description (.txt) files.
 3. Find an example that is close to what you want to achieve.
 4. Modify that configuration file to achieve what you want to achieve.
 5. Use the "transform" sub-command to read your data and output it transformed
    or reorganized according to your modified configuration file.
 6. Read the produced output. If necessary go back to step 4 and adjust how the
    data is transformed.

### Example configuration files

When using the "cfg-example" sub-command to generate an example configuration
file (say example.cfg) a text file describing the configuration and the syntax
of the configuration file is also generated. If the example configuration
file is named example.cfg, then the text file describing the configuration
is named example.txt.

You can generate several example configuration files each with an accompanying
text file describing it.

Read the text file describing the configuration file while looking at the
configuration file to understand the syntax and the possible options.

## Performance

This application is written for the moderate amounts of data when
registering participants for the majority of sports events, or
processing scrum backlogs in excel. If you have millions of rows
this application is not for you.

With an input file consisting of 20 columns producing an output file of
11 columns I have measured the following performance:

| number of rows |   on MacBook Air M4   |
|----------------|-----------------------|
|            100 | less than 0.2 seconds |
|           1000 | less than 0.3 seconds |
|         10 000 | less than 2.5 seconds |
|         20 000 |  less than 5 seconds  |
|         40 000 | less than 11 seconds  |
|         80 000 | less than 23 seconds  |
|        120 000 | less than 34 seconds  |

Naturally your performance will be different based on computer hardware,
operating system and Python version. Generally it should be reasonably
fast for less than 10 000 rows, painfully slow but somewhat usable up
to 100 000 rows and probably so slow that it is unusable for more
than 100 000 rows.

## Description of how to write/change the configuration file

The configuration file is in JSON syntax.
[https://en.wikipedia.org/wiki/JSON](https://en.wikipedia.org/wiki/JSON)

The keywords and the nesting is important. The order of keywords
have no significance (the examples use alphabetical order).
Indentation and line breaks have no significance.

The encoding for the configuration file must be UTF-8. (US-ASCII is a
subset of UTF-8.)

It is recommended that you let the command generate a configuration
file and then edit that file to match your needs. It is NOT recommended
that the user writes the configuration file from scratch.

## column_ref

The keyword **column_ref** is used to tell if the configuration
references columns by name (value *"BY_NAME"*) or by number
(value *"BY_NUMBER"*). The syntax of parts of the configuration file
is slightly different depending on how columns are referenced.
Thus, you first have to decide how you want to reference columns.

Generally it is easier to write a correct configuration using *"BY_NAME"*,
so *"BY_NAME"* is recommended when possible. *"BY_NAME"* imposes a few
restrictions:

- each column needs to have a unique text at the first line to use as column name.
- the column names need to be stable and known.

If you cannot meet these restrictions you will have to reference the
columns by number *"BY_NUMBER"*. When referencing the columns by number
it is important to note that the column numbers change when splitting,
merging, inserting or removing columns.

When referencing *"BY_NUMBER"* the columns are numbered from left to
right. The leftmost column is number 0.

## Type of input and output file

The type of input file to read is configured in the **"input_table"**
section. The type of output file to write is configured in the
**"output_table"** section. These sections use the TableIO JSON
configuration syntax.

The most important keyword in each table section is **"format_name"**.
Generated examples commonly use *"Excel"* or *"CSV"*, but TableIO may
support more formats. In most cases you should leave **"implementation"**
missing or set to *null*. Then TableIO chooses a suitable implementation.
If you know that you need a specific implementation, you can add it
manually.

CSV files may differ slightly depending on the programs used to read/write
them and the locale used. CSV-specific settings are configured in the
nested **"csv"** section inside **"input_table"** or **"output_table"**.
The **"character_encoding"** keyword in the table section controls the
text encoding.
[https://en.wikipedia.org/wiki/Character_encoding](https://en.wikipedia.org/wiki/Character_encoding)

Output tables can request extra presentation features. Set
**"output_borders"** to *true* to request table borders. Set
**"output_filtered_table"** to *true* to request a filtered table area.
These features are ignorable requests. TableIO will prefer an
implementation that can provide them, but the transform can still write
the output if the selected file format cannot represent them.

## Extra spaces in excel input files

When viewing an excel file in excel it is very hard to notice if
some string value in a cell has trailing whitespace. These trailing
spaces can make the further processing of a file difficult
as the strings in the file are not what you thought they are.
(This especially is a problem if you have trailing spaces in a
cell on the first line, and refer to columns by their names.)

The configuration **"in_excel_col_name_strip"** can be set to true,
to strip off leading and trailing whitespace from all columns
values read from the first line of the excel input file.
The configuration **"in_excel_values_strip"** can be set to true,
to strip off leading and trailing whitespace from all columns
values read from the other lines (not the first line) of
the excel input file. (In version 0.6.2 and earlier these are
missing. To be compatible with version 0.6.2 and earlier
**"in_excel_col_name_strip"** and  **"in_excel_values_strip"**
defaults to false if missing in the configuration file.)

## Column manipulation

A number of records starting with "s" and a number describe
column manipulation to be done. These manipulations are done
in order of the number: s01_(something) before s02_(something),
before s03_(something). As the "s" number records might add, remove
or rename columns, it is important to keep track of the order that
they are applied.

Some "s" numbers are only used if columns are referenced
by name while other "s" numbers are only used if columns are
referenced by number.

### "s01_split_rows"

The first operation is to split rows based on column values.
This operation is configured using the **"s01_split_rows"** record,
that has an array of splits to be done. Each row split has the
following keys: **"column"**, **"separators"** and **"not_separators"**.

The **"column"** keyword is used to identify the column that is
split into several rows. This is a column number in the case of
*"BY_NUMBER"*, and a column name/title in the case of *"BY_NAME"*.

New rows will be created so that the parts of the identified
column will be put in that column only one part per row.
The other columns (except the one column being split) will be
replicated identically across all rows split from this row.

**"separators"** take as argument a list of strings. If any of
these strings are present in the identified column it is seen
as the separator between the parts of the column value that go
into different rows.

**"not_separators"** take as argument a list of strings. These strings
are not regarded as separators even if they include the strings
of one or more separator. (For instance ";" could be a separator,
but using **"not_separators"** the string "\\;" could be seen as not
a separator.)

### "s02_merge_rows"

The next operation is to merge rows based on column values.
This is the opposite operation to the splitting of rows.
Row merging is configured using the **"s02_merge_rows"** record,
that has an array of merges to be done. Each row merge has
the following keys: **"columns"** and **"separator"**.

The **"columns"** keyword is used to identify the columns that need
to have identical values to merge two or more rows. The **"columns"**
keyword takes a list of columns. These are column numbers in the case of
*"BY_NUMBER"*, and column names/titles in the case of *"BY_NAME"*.

For each column that has the same value for all rows merged, that
value will be in the merged row. When rows being merged have different
values for a column, the set of unique values from different rows
will form the value for that column in the merged row.

The **"separator"** keyword is used to specify the string that is
concatenated between values for one column from different rows
(in case the column has different values in different rows).

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

To do this transformation we use the **"s01_split_rows"** to split
rows based on the "To" column using " + " as separator.
Then we use **"s02_merge_rows"** to merge the rows that have
identical values in the "To" column.

### "s03_split_columns"

The first column operation that is done is splitting of columns.
The key **"s03_split_columns"** has an array of splits to be
done. (When the list of splits has more than one split,
the least confusion is to split columns to the right before
columns to the left. Named references also helps to avoid
confusion.)

Each split has the following keys: **"column"**, **"separator"** and
**"where"**. In the case of number column references there is
also the key **"store_single"**. In the case of named column references
there is instead the key **"right_name"**.

**"column"** takes as argument the column to be split. This is a
column number in the case of *"BY_NUMBER"*, and a column name/title
in the case of *"BY_NAME"*.

**"separator"** is a string of characters that indicate the position
where the column shall be split. (For instance a single space
if splitting between two words.)

**"where"** can either have the value *"RIGHTMOST"* or the value
*"LEFTMOST"*. This is used if the separator string is present
more than once in the column value to split. The split is then
done at the leftmost or rightmost occurrence according to the
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

### "s04_remove_columns"

**"s04_remove_columns"** is only used with column references *"BY_NUMBER"*.
The value of **"s04_remove_columns"** is a list of column numbers to
remove. (For column references *"BY_NAME"* see **"s10_column_order"**.)

### "s05_merge_columns"

The key **"s05_merge_columns"** has an array of merges to be done.
Each merge have the keys **"columns"** and **"separator"**.

**"columns"** have a list of column references. If *"BY_NAME"* the
column references are column names/titles. If *"BY_NUMBER"* the
column references are column numbers.

**"separator"** is a string of characters that is inserted between
the column values being merged.

### "s06_place_columns_first"

**"s06_place_columns_first"** is only used with column references *"BY_NUMBER"*.
The key **"s06_place_columns_first"** has a value that is a list of the
column numbers to be placed first in order. This step re-orders the
columns.
(For column references *"BY_NAME"* see **"s10_column_order"**.)

### "s07_rename_columns"

The key **"s07_rename_columns"** has a value that is a list of column rename
operations. Each column rename operation has the keys **"column"** and **"name"**.

**"column"** is the number/name of the column before renaming. This is a
column number in the case of *"BY_NUMBER"*, and a column name/title
in the case of *"BY_NAME"*.

**"name"** is the new name/title of the column identified by **"column"**.

### "s08_insert_columns"

The key **"s08_insert_columns"** has a value that is a list of columns to
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

### "s09_rewrite_columns"

The key **"s09_rewrite_columns"** has a value that is a list of
rewrite operations that will be applied in order. Each rewrite
operation is described by several keys.

**"kind"** key is used to specify which kind of rewrite operation
this is. Different values for **"kind"** will have different other
keys. **"kind"** can have the values: *"STRIP"*, *"REMOVECHARS"*,
*"STR_SUBSTITUTE"* and *"REGEX_SUBSTITUTE"*.

**"kind"** value *"STRIP"* shall have the keys: **"kind"**, **"case"**, **"chars"**
and **"column"**. This case means that the characters in **"chars"** shall
be stripped off the beginning and end of the column value. If key
**"chars"** is empty string, then whitespace is stripped off.

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
case of *"BY_NUMBER"*, and a column name/title in the case of *"BY_NAME"*.

**"chars"** is a list of characters written as a string. These characters
are removed or stripped off the beginning/end depending on the **"kind"**
value.

**"case"** specifies if comparison of column value to value of **"chars"** or
value of **"from"** shall be case sensitive or not. Possible values are
*"MATCH_CASE"* and *"IGNORE_CASE"*.

**"from"** specifies what part of the column value shall be substituted
when matching.

**"to"** specifies the string that substitution will replace **"from"** with.

### "s10_column_order"

The key **"s10_column_order"** is used only in the *"BY_NAME"* case.
The value of the **"s10_column_order"** key is a list of column names.
The columns will be output in this order.
Columns not mentioned in **"s10_column_order"** will not be output,
and will thus be removed.
(For *"BY_NUMBER"* see **"s04_remove_columns"** and **"s06_place_columns_first"**.)

## Utility for recoding text files

When working with text based files like comma separated values (CSV)
the file may have an incorrect character encoding, and even worse
you may not know what character encoding was used to create the file.
As a small tool for these situations the excel-list-transform includes
a small recode utility. Using this you can by trial and error recode
from different encodings until you find which encoding works. (This
may be the way to figure out which character encoding to write into
a configuration file.)

```python
python3 -m excel_list_transform.recode --help
python3 -m excel_list_transform.recode -i a.txt -o b.txt --from cp1250 --to utf-8
```

## Source code

Source code and tests are available at [https://bitbucket.org/tom-bjorkholm/excel-list-transform](https://bitbucket.org/tom-bjorkholm/excel-list-transform).

## Test summary

- Test result: 1368 passed in 65s (0:01:05)
- No flake8 warnings.
- No mypy errors found.
- No python layout warnings.
- Built version(s): 1.0
- Build and test using Python 3.14.5
