#! /usr/local/bin/python3
"""Enumerations used in configuration of excel list transform."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


from enum import Enum, auto


class SplitWhere(Enum):
    """Use leftmost or rightmost separator for splitting."""

    LEFTMOST = auto()
    RIGHTMOST = auto()


class RewriteKind(Enum):
    """Kind of write operation to apply."""

    STRIP = auto()
    REMOVECHARS = auto()
    STR_SUBSTITUTE = auto()
    REGEX_SUBSTITUTE = auto()


class CaseSensitivity(Enum):
    """Shall matching be case sensitive or not."""

    MATCH_CASE = auto()
    IGNORE_CASE = auto()


class ColumnRef(Enum):
    """Are columns referenced by number or name."""

    BY_NUMBER = auto()
    BY_NAME = auto()


class EmptyParts(Enum):
    """Which empty parts to drop when splitting a row."""

    KEEP_ALL = auto()
    DROP_ALL = auto()
    DROP_TRAILING = auto()


class EndSeparators(Enum):
    """Which ends of a value to strip of separators before column split."""

    KEEP = auto()
    STRIP_LEADING = auto()
    STRIP_TRAILING = auto()
    STRIP_BOTH_ENDS = auto()
