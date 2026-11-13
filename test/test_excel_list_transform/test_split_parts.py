#! /usr/local/bin/python3
"""Test handling of empty parts and end separators when splitting."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from copy import deepcopy
import pytest
from excel_list_transform.config_enums import EmptyParts, EndSeparators
from excel_list_transform.split_parts import drop_empty_parts, \
    strip_end_separators

KEEP_ALL = EmptyParts.KEEP_ALL
DROP_ALL = EmptyParts.DROP_ALL
DROP_TRAIL = EmptyParts.DROP_TRAILING
KEEP = EndSeparators.KEEP
LEADING = EndSeparators.STRIP_LEADING
TRAILING = EndSeparators.STRIP_TRAILING
BOTH = EndSeparators.STRIP_BOTH_ENDS


@pytest.mark.parametrize('parts,mode,res',
                         [(['a', '', 'b', ''], KEEP_ALL, ['a', '', 'b', '']),
                          (['a', '', 'b', ''], DROP_ALL, ['a', 'b']),
                          (['a', '', 'b', ''], DROP_TRAIL, ['a', '', 'b']),
                          (['a', 'b', '', ''], DROP_TRAIL, ['a', 'b']),
                          (['', 'a'], DROP_TRAIL, ['', 'a']),
                          (['', 'a'], DROP_ALL, ['a']),
                          (['a', ' ', '\t', 'b', ' \n'], DROP_ALL,
                           ['a', 'b']),
                          (['a', ' ', 'b', ' ', '  '], DROP_TRAIL,
                           ['a', ' ', 'b']),
                          ([' a ', ' b'], DROP_ALL, [' a ', ' b']),
                          (['', ''], KEEP_ALL, ['', '']),
                          (['', ''], DROP_ALL, ['']),
                          (['', ' '], DROP_TRAIL, ['']),
                          ([''], KEEP_ALL, ['']),
                          ([''], DROP_ALL, ['']),
                          (['   '], DROP_ALL, ['']),
                          (['   '], DROP_TRAIL, ['']),
                          (['   '], KEEP_ALL, ['   ']),
                          (['abc'], DROP_ALL, ['abc']),
                          ([], KEEP_ALL, ['']),
                          ([], DROP_ALL, ['']),
                          ([], DROP_TRAIL, [''])])
def test_drop_empty_parts(parts: list[str], mode: EmptyParts,
                          res: list[str]) -> None:
    """Test dropping empty parts, and that the input is not modified."""
    orig = deepcopy(parts)
    assert drop_empty_parts(parts=parts, mode=mode) == res
    assert parts == orig


@pytest.mark.parametrize('value,sep,mode,res',
                         [('a,b,', ',', KEEP, 'a,b,'),
                          ('a,b,', ',', TRAILING, 'a,b'),
                          ('a,b,,', ',', TRAILING, 'a,b'),
                          ('a,b, ,', ',', TRAILING, 'a,b'),
                          ('a,b , ', ',', TRAILING, 'a,b '),
                          (',a,b', ',', LEADING, 'a,b'),
                          (', ,a,b', ',', LEADING, 'a,b'),
                          (' ,a,b', ',', LEADING, 'a,b'),
                          (',a,b,', ',', LEADING, 'a,b,'),
                          (',a,b,', ',', TRAILING, ',a,b'),
                          (',a,b,', ',', BOTH, 'a,b'),
                          (', ,a,,b, ,', ',', BOTH, 'a,,b'),
                          ('a,,b', ',', BOTH, 'a,,b'),
                          (' a ', ',', BOTH, ' a '),
                          ('a', ',', TRAILING, 'a'),
                          (',', ',', KEEP, ','),
                          (',', ',', LEADING, ''),
                          (',', ',', TRAILING, ''),
                          (', ,', ',', BOTH, ''),
                          ('', ',', KEEP, ''),
                          ('', ',', BOTH, ''),
                          ('   ', ',', TRAILING, ''),
                          ('   ', ',', KEEP, '   '),
                          ('Anna Smith  ', ' ', TRAILING, 'Anna Smith'),
                          ('  Anna Smith', ' ', LEADING, 'Anna Smith'),
                          ('a, b, ', ', ', TRAILING, 'a, b'),
                          ('a, b,', ', ', TRAILING, 'a, b,'),
                          ('a,,,', ',,', TRAILING, 'a,,,'),
                          ('a,,,,', ',,', TRAILING, 'a'),
                          ('a.*b.*', '.*', TRAILING, 'a.*b')])
def test_strip_end_seps(value: str, sep: str, mode: EndSeparators,
                        res: str) -> None:
    """Test stripping of separators and empty parts at value ends."""
    assert strip_end_separators(value=value, separator=sep, mode=mode) == res


@pytest.mark.parametrize('mode', list(EndSeparators))
def test_strip_end_seps_empty(mode: EndSeparators) -> None:
    """Test that an empty separator is rejected."""
    with pytest.raises(AssertionError):
        _ = strip_end_separators(value='a,b', separator='', mode=mode)
