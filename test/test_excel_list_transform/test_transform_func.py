#! /usr/local/bin/python3
"""Test the excel_list_transform functions functionality."""

# Copyright (c) 2024 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


import pytest
from excel_list_transform.transform_func import file_must_exist


def test_file_must_exist_ok(capsys):
    """Test file_must_exit with existing file."""
    file_must_exist('/bin/ls')
    out, err = capsys.readouterr()
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('fnam', ['/bin/nosuchfile.nonexistent',
                                  '/bin/nosuchdir.nonexistent/nosuchfile'])
def test_file_must_exist_nok1(fnam):
    """Test 1 file_must_exit with non-existing file."""
    with pytest.raises(SystemExit) as exc:
        file_must_exist(fnam)
    assert exc.value.code == 1
