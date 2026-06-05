#! /usr/local/bin/python3
"""Test TableIO handling helpers."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from typing import Optional
import pytest
from excel_list_transform import handle_tableio
from excel_list_transform.handle_tableio import OverwriteAnswer


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
