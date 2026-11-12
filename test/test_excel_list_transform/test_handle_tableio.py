#! /usr/local/bin/python3
"""Test TableIO handling helpers."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from collections.abc import Callable
from pathlib import Path
from typing import Optional
import openpyxl
import pytest
from pytest import CaptureFixture, MonkeyPatch
from test_excel_list_transform.tableio_helpers import configure_input_excel, \
    configure_output_csv, write_excel_num, read_excel_num, \
    configure_input_csv, write_csv_num
from excel_list_transform import handle_tableio
from excel_list_transform.commontypes import NameData, NumData
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.handle_tableio import OverwriteAnswer, \
    read_table_num, read_table_named, write_table_num


def parse_overwrite_answer(answer: str) -> Optional[OverwriteAnswer]:
    """Return parsed overwrite answer through the private implementation."""
    # pylint: disable-next=protected-access
    return handle_tableio._parse_overwrite_answer(answer)


@pytest.mark.parametrize('answer', ['y', 'Y', 'ye', 'yE', 'yes', 'YES'])
def test_parse_overwrite_yes(answer: str) -> None:
    """Test overwrite answer spellings accepted as yes."""
    assert parse_overwrite_answer(answer) == OverwriteAnswer.YES


@pytest.mark.parametrize('answer', ['', 'n', 'N', 'no', 'NO'])
def test_parse_overwrite_no(answer: str) -> None:
    """Test overwrite answer spellings accepted as no."""
    assert parse_overwrite_answer(answer) == OverwriteAnswer.NO


def test_parse_overwrite_bad() -> None:
    """Test unknown overwrite answers are not accepted as yes or no."""
    assert parse_overwrite_answer('maybe') is None


def test_read_num_strip_excel(tmp_path: Path) -> None:
    """Test Excel input reading can trim columns, headers, and values."""
    filename = str(tmp_path / 'input.xlsx')
    data: NumData = [[' Name ', ' Count ', 'Extra'],
                     [' Alice ', 7, ' ignored ']]
    write_excel_num(data=data, filename=filename)
    cfg = ConfigXlsListTransfNum()
    configure_input_excel(cfg)
    cfg.max_column_read = 2
    cfg.in_excel_col_name_strip = True
    cfg.in_excel_values_strip = True
    assert read_table_num(filename=filename, cfg=cfg) == [
        ['Name', 'Count'], ['Alice', 7]]


STRIP_DATA: NumData = [[' Name ', 5, ' Phone '], [' Alice ', '  ', 7],
                       [' ', None, ' x ']]


@pytest.mark.parametrize('cols, values, expected',
                         [(False, False, STRIP_DATA),
                          (True, False, [['Name', 5, 'Phone'],
                                         [' Alice ', '  ', 7],
                                         [' ', None, ' x ']]),
                          (False, True, [[' Name ', 5, ' Phone '],
                                         ['Alice', None, 7],
                                         [None, None, 'x']]),
                          (True, True, [['Name', 5, 'Phone'],
                                        ['Alice', None, 7],
                                        [None, None, 'x']])])
def test_read_excel_strip(tmp_path: Path, cols: bool, values: bool,
                          expected: NumData) -> None:
    """Test Excel strip flags, non-str and whitespace-only cells.

    Values that are empty after strip are read as empty (None).
    """
    filename = str(tmp_path / 'input.xlsx')
    write_excel_num(data=STRIP_DATA, filename=filename)
    assert read_excel_num(filename=filename, max_col_read=3,
                          strip_col_names=cols, strip_values=values) == \
        expected


def test_read_csv_no_strip(tmp_path: Path) -> None:
    """Test that the Excel strip flags do not affect CSV input."""
    filename = str(tmp_path / 'input.csv')
    write_csv_num(data=[[' Name ', ' Phone '], [' Alice ', '']],
                  filename=filename)
    cfg = ConfigXlsListTransfNum()
    configure_input_csv(cfg)
    cfg.in_excel_col_name_strip = True
    cfg.in_excel_values_strip = True
    assert read_table_num(filename=filename, cfg=cfg) == [
        [' Name ', ' Phone '], [' Alice ', None]]


def test_read_named_dup_title(capsys: CaptureFixture[str],
                              tmp_path: Path) -> None:
    """Test that titles equal after Excel strip are reported as duplicates."""
    filename = str(tmp_path / 'input.xlsx')
    data: NumData = [[' Name', 'Name '], ['a', 'b']]
    write_excel_num(data=data, filename=filename)
    cfg = ConfigXlsListTransfNum()
    configure_input_excel(cfg)
    cfg.in_excel_col_name_strip = True
    with pytest.raises(SystemExit) as exc_info:
        _ = read_table_named(filename=filename, cfg=cfg)
    out, err = capsys.readouterr()
    assert exc_info.value.code == 1
    assert '' == out
    assert f'Duplicate input column name "Name" in file {filename}.' in err
    cfg.in_excel_col_name_strip = False
    assert read_table_named(filename=filename, cfg=cfg) == [
        {' Name': 'a', 'Name ': 'b'}]


def _empty_input(tmp_path: Path, cfg: ConfigXlsListTransfNum,
                 excel: bool) -> str:
    """Create an input file without rows and configure cfg to read it."""
    if not excel:
        configure_input_csv(cfg)
        filename = tmp_path / 'input.csv'
        filename.write_text('', encoding='utf-8')
        return str(filename)
    configure_input_excel(cfg)
    cfg.in_excel_col_name_strip = True
    cfg.in_excel_values_strip = True
    filename = tmp_path / 'input.xlsx'
    openpyxl.Workbook().save(filename)
    return str(filename)


@pytest.mark.parametrize('reader', [read_table_num, read_table_named])
@pytest.mark.parametrize('excel', [False, True])
def test_read_empty_input(capsys: CaptureFixture[str], tmp_path: Path,
                          reader: Callable[[str, ConfigXlsListTransfNum],
                                           NumData | NameData],
                          excel: bool) -> None:
    """Test that an input file without rows gives a clear error."""
    cfg = ConfigXlsListTransfNum()
    filename = _empty_input(tmp_path=tmp_path, cfg=cfg, excel=excel)
    with pytest.raises(SystemExit) as exc_info:
        _ = reader(filename, cfg)
    out, err = capsys.readouterr()
    assert exc_info.value.code == 1
    assert '' == out
    assert f'Input file {filename} has no rows.' in err


def _set_overwrite_answer(monkeypatch: MonkeyPatch, answer: str) -> None:
    """Mock the interactive overwrite answer."""
    def read_answer(prompt: str) -> str:
        """Return a fixed overwrite answer."""
        assert prompt == 'Overwrite it? [y/N] '
        return answer
    monkeypatch.setattr('builtins.input', read_answer)


def test_write_num_ow_yes(capsys: CaptureFixture[str],
                          monkeypatch: MonkeyPatch, tmp_path: Path) -> None:
    """Test numbered table writing can overwrite after user approval."""
    cfg = ConfigXlsListTransfNum()
    configure_output_csv(cfg)
    filename = str(tmp_path / 'out.csv')
    Path(filename).write_text('old', encoding='utf-8')
    _set_overwrite_answer(monkeypatch, 'yes')
    written_file = write_table_num(data=[['Name'], ['Alice']],
                                   filename=filename, cfg=cfg)
    out, err = capsys.readouterr()
    assert written_file == filename
    assert f'Output file {filename} already exists.' in out
    assert err == ''


@pytest.mark.parametrize('answer', ['', 'no', 'maybe'])
def test_write_num_ow_no(capsys: CaptureFixture[str], monkeypatch: MonkeyPatch,
                         tmp_path: Path, answer: str) -> None:
    """Test numbered table writing rejects missing overwrite approval."""
    cfg = ConfigXlsListTransfNum()
    configure_output_csv(cfg)
    filename = str(tmp_path / 'out.csv')
    Path(filename).write_text('old', encoding='utf-8')
    _set_overwrite_answer(monkeypatch, answer)
    with pytest.raises(FileExistsError) as exc_info:
        write_table_num(data=[['Name'], ['Alice']], filename=filename, cfg=cfg)
    out, err = capsys.readouterr()
    assert str(exc_info.value) == f'Output file already exists: {filename}'
    assert f'Output file {filename} already exists.' in out
    assert err == ''
    assert Path(filename).read_text(encoding='utf-8') == 'old'
