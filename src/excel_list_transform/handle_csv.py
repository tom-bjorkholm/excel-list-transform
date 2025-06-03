#! /usr/local/bin/python3
"""Handle reading and writing of an excel file."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


from csv import reader, DictReader, writer, DictWriter, Dialect
from excel_list_transform.handle_empty_column import \
    handle_empty_column_in_lst, handle_empty_column_in_dict_lst
from excel_list_transform.num_named_conversion import named_cols_from_num_cols
from excel_list_transform.commontypes import NumData, NumDataSeq, NameData


def read_csv_num(filename: str, dialect: type[Dialect], encoding: str,
                 max_column_read: int) -> NumData:
    """Read a csv file in specified dialect (numbered columns)."""
    res: NumDataSeq = []
    with open(file=filename, mode="r",
              encoding=encoding, newline='') as cfile:
        creader = reader(cfile, dialect=dialect)
        for row in creader:
            res.append(list(row)[:max_column_read])
    return handle_empty_column_in_lst(res)


def read_csv_named_use_num(filename: str, dialect: type[Dialect],
                           encoding: str,
                           max_column_read: int) -> NameData:
    """Read a csv file using numbered columns for named columns."""
    data = read_csv_num(filename=filename, dialect=dialect, encoding=encoding,
                        max_column_read=max_column_read)
    return named_cols_from_num_cols(data=data, filename=filename)


def read_csv_named(filename: str, dialect: type[Dialect], encoding: str,
                   max_column_read: int) -> NameData:
    """Read a csv file in specified dialect (named columns)."""
    res = []
    first_line: bool = True
    with open(file=filename, mode="r",
              encoding=encoding, newline='') as cfile:
        creader = DictReader(cfile, dialect=dialect)
        for row in creader:
            if first_line:
                if len(row) > max_column_read:
                    return read_csv_named_use_num(filename=filename,
                                                  dialect=dialect,
                                                  encoding=encoding,
                                                  max_column_read=  # noqa: E251, E501
                                                  max_column_read)
                first_line = False
            res.append(row)
    return handle_empty_column_in_dict_lst(res)


def write_csv_num(data: NumData, filename: str, dialect: type[Dialect],
                  encoding: str) -> None:
    """Write a csv file in specified dialect."""
    with open(file=filename, mode="w", encoding=encoding,
              newline='') as cfile:
        cwriter = writer(cfile, dialect=dialect)
        for row in data:
            cwriter.writerow(row)


def write_csv_named(data: NameData, filename: str, dialect: type[Dialect],
                    encoding: str, column_order: list[str]) -> None:
    """Write a csv file in specified dialect."""
    with open(file=filename, mode="w", encoding=encoding,
              newline='') as cfile:
        cwriter = DictWriter(f=cfile, dialect=dialect, fieldnames=column_order,
                             extrasaction='ignore')
        cwriter.writeheader()
        for row in data:
            cwriter.writerow(row)
