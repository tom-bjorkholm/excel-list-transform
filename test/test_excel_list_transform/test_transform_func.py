#! /usr/local/bin/python3
"""Test the excel_list_transform functions functionality."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


import pytest
from pytest import CaptureFixture
from excel_list_transform.file_must_exist import file_must_exist


def test_file_must_exist_ok(capsys: CaptureFixture[str]) -> None:
    """Test file_must_exit with existing file."""
    file_must_exist('/bin/ls')
    out, err = capsys.readouterr()
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('fnam', ['/bin/nosuchfile.nonexistent',
                                  '/bin/nosuchdir.nonexistent/nosuchfile'])
def test_file_must_exist_nok1(fnam: str) -> None:
    """Test 1 file_must_exit with non-existing file."""
    with pytest.raises(SystemExit) as exc:
        file_must_exist(fnam)
    assert exc.value.code == 1


def test_file_missing_content(capsys: CaptureFixture[str]) -> None:
    """Test file_must_exit with description of expected file content."""
    with pytest.raises(SystemExit) as exc:
        file_must_exist('/bin/nosuchfile.nonexistent',
                        with_content_txt='configuration data')
    out, err = capsys.readouterr()
    assert exc.value.code == 1
    assert out == ''
    assert 'with configuration data does not exist' in err
