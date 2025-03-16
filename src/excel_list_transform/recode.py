#! /usr/local/bin/python3
"""Read with one encoding and write with another encoding."""

import sys
from copy import deepcopy
from typing import Optional
import argparse


def recode(infilename: str, inencoding: str,
           outfilename: str, outencoding: str) -> None:
    """Read with one encoding and write with another encoding."""
    try:
        with open(file=infilename, mode='r', errors='strict',
                  encoding=inencoding) as inf:
            try:
                with open(file=outfilename, mode='x', errors='strict',
                          encoding=outencoding) as outf:
                    for line in inf:
                        outf.write(line)
            except FileExistsError:
                print(f'Error: Output file {outfilename} already exists.',
                      file=sys.stderr)
                sys.exit(1)
            except LookupError:
                print(f'Error: Unrecognized output encoding {outencoding}',
                      file=sys.stderr)
                sys.exit(1)
            except ValueError:
                print('Error: file content cannot be decoded/encoded ' +
                      f'from {inencoding} to {outencoding}.',
                      file=sys.stderr)
                sys.exit(1)
    except FileNotFoundError:
        print(f'Error: Input file {infilename} not found.',
              file=sys.stderr)
        sys.exit(1)
    except LookupError:
        print(f'Error: Unrecognized input encoding {inencoding}',
              file=sys.stderr)
        sys.exit(1)
    except IOError as exc:  # pragma: no cover
        print(f'Unexpected I/O error: {exc.strerror}',
              file=sys.stderr)
        sys.exit(1)
    print(f'Wrote file {outfilename}', file=sys.stderr)


def recode_cmd(args: Optional[list[str]] = None) -> None:
    """Command to read with one encoding and write with another encoding."""
    desc = 'Command to read with one encoding and write with another encoding.'
    arguments = deepcopy(args) if args is not None else deepcopy(sys.argv)
    parser = argparse.ArgumentParser(prog='recode', description=desc)
    parser.add_argument('-i', '--input', dest='infile', type=str, nargs=1,
                        required=True, help='Input file name')
    parser.add_argument('-o', '--output', dest='outfile', type=str, nargs=1,
                        required=True, help='Output file name')
    parser.add_argument('-f', '--from', dest='fromenc', type=str, nargs=1,
                        required=True,
                        help='From encoding. Encoding in input file')
    parser.add_argument('-t', '--to', dest='toenc', type=str, nargs=1,
                        required=True,
                        help='To encoding. Encoding in output file')
    if len(arguments) > 2 and 'python' in arguments[0]:
        del arguments[0]
    if len(arguments) > 2 and '-m' == arguments[0]:
        del arguments[0]
    while len(arguments) >= 1 and arguments[0][-3:] == '.py':
        del arguments[0]
    parsed_arg = parser.parse_args(args=arguments)
    recode(infilename=parsed_arg.infile[0], inencoding=parsed_arg.fromenc[0],
           outfilename=parsed_arg.outfile[0], outencoding=parsed_arg.toenc[0])


if __name__ == '__main__':  # pragma: no cover
    recode_cmd()
