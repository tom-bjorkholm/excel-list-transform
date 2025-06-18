#! /usr/local/bin/python3
"""The test cases for generate_txt."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


from tempfile import TemporaryDirectory
import pytest
from excel_list_transform.generate_txt import generate_syntax_txt


@pytest.mark.parametrize('txt_file, edescr, cfg_file',
                         [('a.txt',
                           'abc is not xyz', 'a.cfg'),
                          ('x.txt',
                           'some other description of example',
                           'x.cfg')])
def test_generate_syntax_txt(capsys, txt_file, edescr, cfg_file):
    """Test generate_syntax_txt."""
    content = None
    with TemporaryDirectory() as dname:
        fname = dname + '/' + txt_file
        cfname = dname + '/' + cfg_file
        generate_syntax_txt(filename=fname,
                            example_description=edescr,
                            cfgfilename=cfname)
        with open(file=fname, mode='r', encoding='utf-8') as file:
            content = file.read()
    out, err = capsys.readouterr()
    assert edescr in content
    assert 'Explanation for example configuration file' in content
    assert cfg_file in content
    assert 's10_column_order' in content
    assert '' == out
    assert '' == err
