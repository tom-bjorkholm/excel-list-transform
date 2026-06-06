#! /usr/local/bin/python3
"""Test TableIO handling helpers."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from pathlib import Path
from typing import Optional
import pytest
from pytest import CaptureFixture, MonkeyPatch
from test_excel_list_transform.tableio_helpers import configure_input_excel, \
    configure_output_csv, write_excel_num
from excel_list_transform import handle_tableio
from excel_list_transform.commontypes import NumData
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum
from excel_list_transform.handle_tableio import OverwriteAnswer, \
    read_table_num, write_table_num


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
