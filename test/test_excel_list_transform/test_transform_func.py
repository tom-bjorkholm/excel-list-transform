#! /usr/local/bin/python3
"""Test the excel_list_transform functions functionality."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from pathlib import Path
import pytest
from pytest import CaptureFixture
from excel_list_transform.file_must_exist import file_must_exist


def test_file_must_exist_ok(capsys: CaptureFixture[str],
                            tmp_path: Path) -> None:
    """Test file_must_exit with existing file."""
    existing = tmp_path / 'existing.txt'
    existing.touch()
    file_must_exist(str(existing))
    out, err = capsys.readouterr()
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('fnam', ['nosuchfile.nonexistent',
                                  'nosuchdir.nonexistent/nosuchfile'])
def test_file_must_exist_nok1(tmp_path: Path, fnam: str) -> None:
    """Test 1 file_must_exit with non-existing file."""
    with pytest.raises(SystemExit) as exc:
        file_must_exist(str(tmp_path / fnam))
    assert exc.value.code == 1


def test_file_missing_content(capsys: CaptureFixture[str],
                              tmp_path: Path) -> None:
    """Test file_must_exit with description of expected file content."""
    with pytest.raises(SystemExit) as exc:
        file_must_exist(str(tmp_path / 'nosuchfile.nonexistent'),
                        with_content_txt='configuration data')
    out, err = capsys.readouterr()
    assert exc.value.code == 1
    assert out == ''
    assert 'with configuration data does not exist' in err
