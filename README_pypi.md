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
 * writest the resulting data (that is a list with columns) to an excel file or to a comma separate values (CSV) file.

How this is done is governed by a configuration file. The application can create a number of example configuration files with accompanying description text files.

## Installing it

If you want to use it install it using pip. A precondition is that you have Python 3.10.5 or newer installed on you computer. Python can be downloaded from [https://www.python.org/downloads/](https://www.python.org/downloads/).

### Installing on mac and Linux

````
pip3 install excel-list-transform
````

### Installing on Microsoft Windows

````
pip install excel-list-transform
````

## Running the application

### Running the application on mac and Linux

````
python3 -m excel_list_transform --help
python3 -m excel_list_transform example --help
python3 -m excel_list_transform transform --help
python3 -m excel_list_transform example -k forms_to_rrs -r by_name -o example.cfg
python3 -m excel_list_transform transform -c example.cfg -i input.xlsx -o output.xlsx
````

### Running the application on Microsoft Windows

````
python -m excel_list_transform --help
python -m excel_list_transform example --help
python -m excel_list_transform transform --help
python -m excel_list_transform example -k forms_to_rrs -r by_name -o example.cfg
python -m excel_list_transform transform -c example.cfg -i input.xlsx -o output.xlsx
````

## Suggested way to get started

 1. Use the "example" sub-command to generate a few example configuration (.cfg) files with description (.txt) files.
 2. Read the example configuration (.cfg) files and the accompanying description (.txt) files.
 3. Find an example that is close to what you want to achieve.
 4. Modify that configuration file to achieve what you want to achieve.
 5. Use the "transform" sub-command to read your data and output it transformed or reorganized according to your modified configuration file.
 6. Read the produced output. If necessary go back to step 4 and adjust how the data is transformed.

### Description of configuration file

When using the "example" sub-command to generate an example configuration file (say example.cfg) a text file describing the configuration and the syntax of the configuration file is also generated. If the example configuration file is named example.cfg, then the text file descriging the configuration is named example.txt. 

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

## Source code

Source code and tests are available at [https://bitbucket.org/tom-bjorkholm/excel-list-transform](https://bitbucket.org/tom-bjorkholm/excel-list-transform).
