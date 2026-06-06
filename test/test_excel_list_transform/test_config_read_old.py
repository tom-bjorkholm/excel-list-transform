#! /usr/local/bin/python3
"""Test old excel-list-transform configuration normalization."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from enum import Enum, auto
from typing import Optional
import pytest
from tableio import CsvDialect
from excel_list_transform.config_read_old import old_csv_spec_to_tableio, \
    old_file_type_to_format


class OldFileType(Enum):
    """Old-style fake file type enum for conversion tests."""

    CSV = auto()
    EXCEL = auto()
    PDF = auto()


@pytest.mark.parametrize('value, expected',
                         [('csv', 'CSV'),
                          ('EXCEL', 'Excel'),
                          (OldFileType.CSV, 'CSV'),
                          (OldFileType.EXCEL, 'Excel')])
def test_old_file_format(value: str | OldFileType, expected: str) -> None:
    """Test old file type values converted to TableIO format names."""
    assert old_file_type_to_format(value) == expected


@pytest.mark.parametrize('value', ['PDF', OldFileType.PDF])
def test_old_file_type_bad(value: str | OldFileType) -> None:
    """Test unknown old file type values are rejected."""
    with pytest.raises(KeyError) as exc_info:
        old_file_type_to_format(value)
    assert 'Unknown old file type' in str(exc_info.value)


@pytest.mark.parametrize('quoting, expected',
                         [('csv.quote_all', 'all'),
                          ('CSV.QUOTE_MINIMAL', 'minimal'),
                          ('csv.quote_none', 'none'),
                          ('csv.quote_nonnumeric', 'nonnumeric'),
                          (None, None)])
def test_old_csv_quoting(quoting: Optional[str],
                         expected: Optional[str]) -> None:
    """Test old CSV quoting values converted into TableIO names."""
    old_spec: dict[str, object] = {'quoting': quoting}
    new_spec = old_csv_spec_to_tableio(old_spec)
    if expected is None:
        assert 'quoting' not in new_spec
    else:
        assert new_spec['quoting'] == expected


@pytest.mark.parametrize('old_spec, expected',
                         [({}, {'dialect': CsvDialect.EXCEL.name}),
                          ({'name': 'csv.unix_dialect'},
                           {'dialect': CsvDialect.UNIX.name}),
                          ({'name': 'csv.excel_tab'},
                           {'dialect': CsvDialect.EXCEL.name,
                            'delimiter': '\t'}),
                          ({'name': 'csv.excel_tab', 'delimiter': ';'},
                           {'dialect': CsvDialect.EXCEL.name,
                            'delimiter': ';'}),
                          ({'quotechar': "'", 'lineterminator': '\n',
                            'escapechar': '\\'},
                           {'dialect': CsvDialect.EXCEL.name,
                            'quotechar': "'",
                            'lineterminator': '\n',
                            'escapechar': '\\'})])
def test_old_csv_spec(old_spec: dict[str, object],
                      expected: dict[str, object]) -> None:
    """Test old CSV dialect dictionaries converted to TableIO dictionaries."""
    assert old_csv_spec_to_tableio(old_spec) == expected


def test_old_csv_bad_quoting() -> None:
    """Test unknown old CSV quoting values are rejected."""
    with pytest.raises(KeyError) as exc_info:
        old_csv_spec_to_tableio({'quoting': 'csv.quote_invalid'})
    assert 'Unknown old CSV quoting' in str(exc_info.value)
