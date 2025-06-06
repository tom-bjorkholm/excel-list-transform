#! /usr/local/bin/python3
"""Command to transform list in excel or CSV file."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


from copy import deepcopy
from sys import argv as sys_argv
from typing import Optional, TypeAlias
import argparse
from excel_list_transform.transform_func import transform_named_files
from excel_list_transform.generate_cfg import generate_examplecfg
from excel_list_transform.generate_cfg import get_example_names
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.str_to_enum import string_to_enum_best_match
from excel_list_transform.version import Version


def gen_cfg_cmd(args: argparse.Namespace) -> int:
    """Generate example cfg file."""
    outfilename: str = args.output[0]
    cfgtype: str = args.kind[0]
    colref: ColumnRef = string_to_enum_best_match(args.reference[0],
                                                  ColumnRef)
    return generate_examplecfg(filename=outfilename,
                               cfgtype=cfgtype, colref=colref)


def rfmt_cmd(args: argparse.Namespace) -> int:
    """Generate example cfg file."""
    outfilename = args.output[0]
    infilename = args.input[0]
    cfgfilename = args.cfg[0]
    return transform_named_files(infilename=infilename,
                                 outfilename=outfilename,
                                 cfgfilename=cfgfilename)


def version_cmd(_: argparse.Namespace) -> int:
    """Print version information."""
    vers = Version()
    vers.print()
    return 0


USAGE_ORDER = '''
The normal way to use this command is:
(1) Using the "example" sub-command a few example configuration (.cfg) files
with description (.txt) files are generated.
(2) Read the example configuration (.cfg) files and the accompanying
description (.txt) files.
(3) Find an example that is close to what you want to achieve.
(4) Modify that configuration file to achieve what you want to achieve.
(5) Use the "transform" sub-command to read your data and output it
transformed or reorganized according to your modified configuration
file.
(6) Read the produced output. If necessary go back to step 4 and adjust
how the data is transformed.
'''

TXT_DESCRIPTION = '''
When generating an example configuration file a text file describing
the configuration file syntax is also generated, with the same name
as the configuration file but with extension .txt instead of .cfg.
'''

GENERAL_DESCRIPTION = '''
Transform or reorganize list in excel or CSV file. How data is transformed
is described in a configuration file. Name of input file,
output file and configuration file is given as command line
arguments.
The command can also generate a few example configuration files.
When generating an example configuration file the output file name
switch gives the name of the generated configuration file.
'''

SEE_MAIN_HELP = '''
See also help text for main command without sub-commands.
'''

SubParseAct: TypeAlias = 'argparse._SubParsersAction[argparse.ArgumentParser]'


def gen_cfg_args_named(subparsers: SubParseAct, sub_pars_name: str) -> None:
    """Add arguments for generate named example config sub-command."""
    cfg_help = 'Generate example configuration file (example .cfg file). '
    cfg_help += 'Arguments select the kind of configuration file that '
    cfg_help += 'is generated.'
    if sub_pars_name == 'example':
        cfg_help = 'example is an alias for cfg-example.'
    cfg_parser = subparsers.add_parser(sub_pars_name, help=cfg_help,
                                       epilog=USAGE_ORDER,
                                       description=cfg_help +
                                       TXT_DESCRIPTION + SEE_MAIN_HELP)
    cfg_parser.set_defaults(func=gen_cfg_cmd)
    examplekinds = get_example_names()
    kind_help = 'Kind of example to generate configuration file for.'
    kind_help += 'Possible kinds are (' + ', '.join(examplekinds) + ').'
    cfg_parser.add_argument('-k', '--kind', nargs=1, required=True,
                            help=kind_help, choices=examplekinds)
    reftypes = [e.name.lower() for e in ColumnRef]
    ref_help = 'How does configuration file reference to columns '
    ref_help += 'in the data. By column number of by column name. '
    ref_help += 'Possible values are (' + ', '.join(reftypes) + '). '
    ref_help += 'Normally "by_name" column references are easier to get right.'
    cfg_parser.add_argument('-r', '--reference', nargs=1, required=True,
                            help=ref_help, choices=reftypes)
    cfg_output_help = 'Name of configuration (output) file to create.'
    cfg_parser.add_argument('-o', '--output', nargs=1,
                            help=cfg_output_help, required=True)


def gen_cfg_args(subparsers: SubParseAct) -> None:
    """Add arguments for generate example config sub-command."""
    gen_cfg_args_named(subparsers=subparsers, sub_pars_name='example')
    gen_cfg_args_named(subparsers=subparsers, sub_pars_name='cfg-example')


def transf_args(subparsers: SubParseAct) -> None:
    """Add arguments for transform sub-command."""
    transf_help = 'Transform list in excel or CSV file. How data is '
    transf_help += 'transformed is described in a configuration file. Name '
    transf_help += 'of input file, output file and configuration file is '
    transf_help += 'given as command line arguments.'
    transf_parser = subparsers.add_parser('transform', help=transf_help,
                                          epilog=USAGE_ORDER,
                                          description=transf_help +
                                          SEE_MAIN_HELP)
    transf_parser.set_defaults(func=rfmt_cmd)
    transf_parser.add_argument('-c', '--cfg', nargs=1, required=True,
                               help='Configuation file name to use.')
    transf_parser.add_argument('-i', '--input', nargs=1,
                               help='Name of input file.', required=True)
    transf_parser.add_argument('-o', '--output', nargs=1,
                               help='Name of output file to create.',
                               required=True)


def version_args(subparsers: SubParseAct) -> None:
    """Add arguments for version sub-command."""
    version_help = 'Only print versions of excel_list_transform '
    version_help += 'and of main modules used by it and of Python.'
    version_parser = subparsers.add_parser('version', help=version_help,
                                           epilog=USAGE_ORDER,
                                           description=version_help +
                                           SEE_MAIN_HELP)
    version_parser.set_defaults(func=version_cmd)


def transform_cmd(arguments: Optional[list[str]] = None) -> None:
    """Command to transform list in excel or CSV file."""
    epimain = 'More detailed help is available for each sub-command.'
    if arguments is None:  # pragma: no cover
        arguments = sys_argv
    fixed_args = deepcopy(arguments)
    if len(fixed_args) > 2 and 'python' in fixed_args[0]:
        del fixed_args[0]
    if len(fixed_args) > 2 and '-m' == fixed_args[0]:
        del fixed_args[0]
    while len(fixed_args) >= 1 and fixed_args[0][-3:] == '.py':
        del fixed_args[0]
    desc = GENERAL_DESCRIPTION + \
        USAGE_ORDER
    parser = argparse.ArgumentParser(prog='excel_list_transform',
                                     description=desc, epilog=epimain)
    subparsers = parser.add_subparsers(dest='subparser_name', required=True)
    gen_cfg_args(subparsers)
    transf_args(subparsers)
    version_args(subparsers)
    args = parser.parse_args(args=fixed_args)
    _ = args.func(args)


if __name__ == "__main__":  # pragma: no cover
    transform_cmd()
