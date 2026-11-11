#! /usr/local/bin/python3
"""Test the excel_list_transform file extension utility functionality."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from copy import deepcopy
from pathlib import Path
from typing import Optional
import pytest
from pytest import CaptureFixture
from excel_list_transform.file_extension import fix_file_extension


@pytest.mark.parametrize('fil, add, rem, readp, res',
                         [('ls', '.bo', '.fo', True, 'ls'),
                          ('ls', '.bo', '.fo', False, 'ls.bo'),
                          ('ls.bo', '.bo', '.fo', True, 'ls.bo'),
                          ('ls.bo', '.bo', '.fo', False, 'ls.bo'),
                          ('ls.fo', '.bo', '.fo', True, 'ls.bo'),
                          ('ls.fo', '.bo', '.fo', False, 'ls.bo'),
                          ('ls/b', '.bo', '.fo', True, 'ls/b.bo'),
                          ('ls/b', '.bo', '.fo', False, 'ls/b.bo'),
                          ('ls.FO', '.bo', '.fo', False, 'ls.bo'),
                          ('ls.BO', '.bo', '.fo', False, 'ls.BO'),
                          ('ls.bo', '.bo', '.bo', False, 'ls.bo'),
                          ('ls.bo', '.bo', '', False, 'ls.bo'),
                          ('ls.bo', '.bo', None, False, 'ls.bo'),
                          ('ls.fo', '.bo', None, False, 'ls.fo.bo')])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_file_extension_1(capsys: CaptureFixture[str], tmp_path: Path,
                          fil: str, add: str, rem: Optional[str], readp: bool,
                          res: str) -> None:
    """Test fix_file_extension. Only file ls exists in tmp_path."""
    (tmp_path / 'ls').touch()
    fil = str(tmp_path / fil)
    res = str(tmp_path / res)
    filn = deepcopy(fil)
    addn = deepcopy(add)
    remn = deepcopy(rem)
    readn = deepcopy(readp)
    ret = fix_file_extension(filename=filn, ext_to_add=addn,
                             ext_to_remove=remn, for_reading=readn)
    out, err = capsys.readouterr()
    assert filn == fil
    assert addn == add
    assert remn == rem
    assert readn == readp
    assert ret == res
    assert '' == out
    assert '' == err
