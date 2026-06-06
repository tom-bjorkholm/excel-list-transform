#! /usr/local/bin/python3
"""Test common types."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

import pytest
from pytest import CaptureFixture
from excel_list_transform.commontypes import num_row_to_str_list, NumRow


@pytest.mark.parametrize('ind, outd',
                         [(['a', 'x', 'some text'],
                           ['a', 'x', 'some text']),
                          (['a', 4, 3.14],
                           ['a', '4', '3.14'])])
def test_num_row_to_str_ok(capsys: CaptureFixture[str], ind: NumRow,
                           outd: list[str]) -> None:
    """Test num_row_to_str_list for OK cases."""
    ret = num_row_to_str_list(ind)
    out, err = capsys.readouterr()
    assert outd == ret
    assert '' == out
    assert '' == err


def test_num_row_to_str_nok(capsys: CaptureFixture[str]) -> None:
    """Test num_row_to_str_list for not OK case."""
    ind: NumRow = ['a', 2, None, 'x']
    with pytest.raises(TypeError) as exc:
        _ = num_row_to_str_list(ind)
    out, err = capsys.readouterr()
    assert 'Found None when expecting str' in str(exc)
    assert '' == out
    assert '' == err
