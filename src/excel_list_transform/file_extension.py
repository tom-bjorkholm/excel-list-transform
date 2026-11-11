#! /usr/local/bin/python3
"""Functions to add and remove file extensions."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


from os import path
from typing import Optional


def _ends_with_ext(filename: str, ext: str) -> bool:
    """Check (case insensitive) if filename ends with non-empty ext."""
    return ext != '' and filename.lower().endswith(ext.lower())


def fix_file_extension(filename: str, ext_to_add: str,
                       ext_to_remove: Optional[str] = None,
                       for_reading: bool = False) -> str:
    """Add and remove file extensions as needed."""
    if for_reading and path.exists(path=filename):
        return filename
    ret = filename
    if ext_to_remove is not None and _ends_with_ext(ret, ext_to_remove):
        ret = ret[:-len(ext_to_remove)]
    if not _ends_with_ext(ret, ext_to_add):
        ret += ext_to_add
    return ret
