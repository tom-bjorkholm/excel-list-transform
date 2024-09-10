# excel-list-transform

## Background

This python application was born out of an experience at sail racing events. At the start of the events we received an excel list with participants (from a registration web) to enter into the scoring software and into tool for online notice board. The information was present in the excel file, but the columns were all wrong. To avoid the stressful manual rework of the information in the excel list this application was created.

## Using it

If you want to use it install it using pip from [https://pypi.org/project/excel-list-transform](https://pypi.org/project/excel-list-transform). There is no need download anything from Bitbucket to use the application.

### Installing on mac and Linux

````
pip3 install excel-list-transform
````

### Installing on Microsoft Windows

````
pip install excel-list-transform
````

### Information for use

Please see [https://pypi.org/project/excel-list-transform](https://pypi.org/project/excel-list-transform) or please see README_pypi.md

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

## For developers

### Needed environment

#### OS

For running the script and running the test suite you need a mac or a Linux computer. Even if the resulting application can be installed and used on Windows, the scripts for building and testing is only implemented for mac and Linux.

#### Python version

The tests and the script for running the tests, coverage, mypy etc. requires Python version 3.12.5 or newer.

#### Zsh

The scripts are all zsh. zsh is available by default on modern macs. zsh can easily be installed on Linux (on Ubuntu: sudo apt install zsh).

### Internal APIs not quaranteed

The internal APIs in this package are not guaranteed to be stable. They can change without warning between versions.

### Building application

There are 3 scripts for building the application
 * setup_build_environment.zsh
 Run this script first to get the environment set up for building
 * doBuild.zsh
 Run this script to build an installation package (.whl) and to run the tests on it in a venv (virtual environment).
 * clean.zsh
 Deletes all files that was produced by the build to start over from a clean state.

The "testing" includes pytest, pylint, flake8 and mypy.

After running doBuild.zsh you can open reports/index.htm to see all test reports.

After running doBuild.zsh you can do manual test of the built and installed application in the virtual environment ./venv
