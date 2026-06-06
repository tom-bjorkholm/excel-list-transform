#! /usr/local/bin/python3
"""Measure speed for different data set sizes."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


from tempfile import TemporaryDirectory
import sys
from timeit import default_timer
from typing import TextIO
from test_excel_list_transform.test_generate_cfg import ExampleData
from test_excel_list_transform.tableio_helpers import write_excel_num
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.generate_cfg import generate_examplecfg
from excel_list_transform.transform_cmd import transform_cmd


def measure_single_run(dirname: str, refcol: ColumnRef, size: int,
                       kind: str) -> float:
    """Measure speed for a single command run."""
    print(f'Start measuring time for processing {size} rows...')
    start = default_timer()
    transform_cmd(['transform', '-i', dirname + '/in.xlsx',
                   '-o', dirname + '/' + refcol.name + 'out.xlsx',
                   '-c', dirname + '/' + refcol.name + '.cfg'])
    stop = default_timer()
    sec: float = stop - start
    print(f'Processing {size} rows {kind} in {sec} seconds.')
    return sec


def measure_speed_for_size(size: int) -> dict[ColumnRef, float]:
    """Measure speed for for given data set size."""
    ret: dict[ColumnRef, float] = {}
    test_data = ExampleData()
    print(f'Setting up environment and input files for {size} rows.')
    with TemporaryDirectory() as dname:
        for refcol in ColumnRef:
            cfgfile = dname + '/' + refcol.name + '.cfg'
            filename = dname + '/in.xlsx'
            generate_examplecfg(filename=cfgfile, cfgtype='forms_to_rrs',
                                colref=refcol)
        data = test_data.get_sized_form_data(size=size)
        write_excel_num(data=data, filename=filename)
        print(f'Environment and input files for {size} rows ready.')
        for i in ColumnRef:
            ret[i] = measure_single_run(dirname=dname, refcol=i, size=size,
                                        kind=i.name)
    return ret


def print_results(data: dict[int, dict[ColumnRef, float]],
                  file: TextIO) -> None:
    """Print the results of running tests."""
    print('Speed measurement of processing 20 input columns per row.',
          file=file)
    print('Processing results in 11 output columns per row.\n', file=file)
    print('Num rows\tSeconds by_name\t  Seconds by_number', file=file)
    for size, result in data.items():
        print(f'{size:8,}\t{result[ColumnRef.BY_NAME]: 15.2f}' +
              f'\t  {result[ColumnRef.BY_NUMBER]: 17.2f}',
              file=file)


def measure_speed() -> None:
    """Measure speed for different data set sizes."""
#    sizes = [10, 100, 1000, 10*1000, 20*1000, 40*1000, 60*1000,
#             80*1000, 100*1000, 120*1000]
    sizes = [100, 500, 1000, 2*1000, 3*1000, 4*1000, 5*1000]
    resulting_time: dict[int, dict[ColumnRef, float]] = {}
    for i in sizes:
        resulting_time[i] = measure_speed_for_size(i)
    print_results(data=resulting_time, file=sys.stdout)
    with open(file='performance.txt', mode='w', encoding='utf-8') as file:
        print_results(data=resulting_time, file=file)


if __name__ == '__main__':
    measure_speed()
