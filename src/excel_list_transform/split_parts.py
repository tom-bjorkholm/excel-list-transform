#! /usr/local/bin/python3
"""Handle empty parts and end separators when splitting values.

A part is the text between two separators, or between a separator and
the start or end of the value. A part is empty if it is '' or consists
only of whitespace.
"""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from excel_list_transform.config_enums import EmptyParts, EndSeparators


def _is_empty(part: str) -> bool:
    """Return True if the part is '' or only whitespace."""
    return not part.strip()


def drop_empty_parts(parts: list[str], mode: EmptyParts) -> list[str]:
    """Return the parts left after dropping empty parts according to mode.

    Never returns an empty list: if all parts are dropped (or parts is
    empty) the result is [''], so that the split row is kept with ''.
    """
    kept = list(parts)
    if mode == EmptyParts.DROP_ALL:
        kept = [part for part in kept if not _is_empty(part)]
    elif mode == EmptyParts.DROP_TRAILING:
        while kept and _is_empty(kept[-1]):
            kept.pop()
    return kept or ['']


def strip_end_separators(value: str, separator: str,
                         mode: EndSeparators) -> str:
    """Strip separators and empty parts at the ends given by mode.

    The value is split at every separator. STRIP_TRAILING cuts the value
    directly after the last non-empty part, STRIP_LEADING cuts it
    directly before the first non-empty part. A value without non-empty
    parts (only separators and whitespace) becomes '' unless mode is KEEP.
    @param separator  Non-empty string (enforced by config validation).
    """
    assert separator
    if mode == EndSeparators.KEEP:
        return value
    parts = value.split(separator)
    filled = [index for index, part in enumerate(parts)
              if not _is_empty(part)]
    if not filled:
        return ''
    first = 0 if mode == EndSeparators.STRIP_TRAILING else filled[0]
    last = len(parts) if mode == EndSeparators.STRIP_LEADING \
        else filled[-1] + 1
    return separator.join(parts[first:last])
