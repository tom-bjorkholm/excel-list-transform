#! /usr/local/bin/python3
"""Test the Config class (part 2 of tests)."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from typing import Any, cast
# pylint: disable-next=unused-import,ungrouped-imports
from typing import Optional
from copy import deepcopy
import pytest
from pytest import CaptureFixture
from excel_list_transform.config import Config, BackwardCompatible
from excel_list_transform.commontypes import JsonType


@pytest.mark.parametrize('enc, is_ok',
                         [('utf-8', True),
                          ('abc123', False)])
def test_cfg_valid_chr_enc_ok(capsys: CaptureFixture[str], enc: str,
                              is_ok: bool) -> None:
    """Test OK cases of valid_char_encoding."""
    ret = Config.valid_char_encoding(enc)
    out, err = capsys.readouterr()
    assert ret == is_ok
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('enc', [8, True])
def test_cfg_val_chr_enc_nok(capsys: CaptureFixture[str], enc: object) -> None:
    """Test not OK cases of valid_char_encoding."""
    with pytest.raises(Exception) as exc:
        _ = Config.valid_char_encoding(cast(str, enc))
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert 'must be str' in str(exc)


@pytest.mark.parametrize('enc', ['utf-8', 'iso8859-1'])
def test_cfg_check_chr_enc_ok(capsys: CaptureFixture[str], enc: str) -> None:
    """Test OK cases of check_char_encoding."""
    Config.check_char_encoding(enc)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('enc', ['utf-88', 'abc123'])
def test_cfg_chk_chr_enc_nok(capsys: CaptureFixture[str], enc: str) -> None:
    """Test not OK cases of check_char_encoding."""
    with pytest.raises(SystemExit):
        Config.check_char_encoding(enc)
    out, err = capsys.readouterr()
    assert '' == out
    assert f'{enc} is not a recognized encoding' in err


class AbcConfig(Config):
    """Class to test defualt values."""

    def __init__(self, from_json_data_text: Optional[str] = None,
                 from_json_filename: Optional[str] = None) -> None:
        """Construct test config object."""
        self.ab = 'a1b2'
        self.cd = 'c3d4'
        self.ef = 'e5f6'
        super().__init__(from_json_data_text=from_json_data_text,
                         from_json_filename=from_json_filename)

    def _def_vals_for_optional(self) -> dict[str, JsonType]:
        """Return default values for optional parameters."""
        return {'cd': 'cd99', 'ef': 'ef99'}


def test_cfg_abc_dump_ok(capsys: CaptureFixture[str]) -> None:
    """Test dump of default constructed AbcConfig."""
    abc = AbcConfig()
    jstext = abc.as_json_string()
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    jstext = jstext.replace('\n', ' ')
    for _ in range(10):
        jstext = jstext.replace('  ', ' ')
    assert jstext == '{ "ab": "a1b2", "cd": "c3d4", "ef": "e5f6" }'


@pytest.mark.parametrize('jstext, aval, cval, fval',
                         [('{ "ab": "a1b2", "cd": "c3d4", "ef": "e5f6" }',
                           'a1b2', 'c3d4', 'e5f6'),
                          ('{ "ab": "donald", "cd": "duck", "ef": "mouse" }',
                           'donald', 'duck', 'mouse'),
                          ('{ "ab": "aaa", "cd": "mickey" }',
                           'aaa', 'mickey', 'ef99'),
                          ('{ "ab": "donald", "ef": "duck" }',
                           'donald', 'cd99', 'duck'),
                          ('{ "ab": "duck"}',
                           'duck', 'cd99', 'ef99')])
def test_cfg_def_val_json_ok(capsys: CaptureFixture[str], jstext: str,
                             aval: str, cval: str, fval: str) -> None:
    """Test construction of cfg from json with default values."""
    abc = AbcConfig(from_json_data_text=jstext, from_json_filename=None)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err
    assert abc.ab == aval
    assert abc.cd == cval
    assert abc.ef == fval


@pytest.mark.parametrize('jstext',
                         ['{ "cd": "c3d4", "ef": "e5f6" }',
                          '{ "cd": "duck", "ef": "mouse" }',
                          '{ "cd": "mickey" }',
                          '{ "ef": "duck" }', '{}'])
def test_cfg_def_val_json_nok(capsys: CaptureFixture[str],
                              jstext: str) -> None:
    """Test construction of cfg from json with default values."""
    with pytest.raises(KeyError):
        _ = AbcConfig(from_json_data_text=jstext, from_json_filename=None)
    out, err = capsys.readouterr()
    assert '' == out
    assert 'No value for ab in' in err


@pytest.mark.parametrize('ind, outd, ren, errtxt',
                         [({'foo': 12, 'bar': 'data'},
                           {'foo': 12, 'bar': 'data'},
                           BackwardCompatible(old='star', new='sun'), ''),
                          ({'foo': 12, 'bar': 'data'},
                           {'sun': 12, 'bar': 'data'},
                           BackwardCompatible(old='foo', new='sun'), ''),
                          ({'foo': 12, 'bar': 'data'},
                           {'foo': 12, 'sun': 'data'},
                           BackwardCompatible(old='bar', new='sun'), ''),
                          ({'foo': 12, 'bar': 'data'},
                           {'foo': 12},
                           BackwardCompatible(old='bar', new='foo'),
                           'Inconsistent configuration:\n' +
                           'Both new config parameter foo and old bar ' +
                           'present.\nIgnoring old parameter bar\n'),
                          ({'foo': {'a': 'b', 'bar': 'c'}, 'bar': 'data'},
                           {'foo': {'a': 'b', 'sun': 'c'}, 'sun': 'data'},
                           BackwardCompatible(old='bar', new='sun'), ''),
                          ({'foo': {'a': 'b', 'foo': 'c'}, 'bar': 'data'},
                           {'sun': {'a': 'b', 'sun': 'c'}, 'bar': 'data'},
                           BackwardCompatible(old='foo', new='sun'), ''),
                          ({'foo': {'foo': 'b', 'bar': 'c'}, 'bar': 'data'},
                           {'sun': {'sun': 'b', 'bar': 'c'}, 'bar': 'data'},
                           BackwardCompatible(old='foo', new='sun'), ''),
                          ({'a': [{'a': 1, 'b': 2}, {'a': 3, 'b': 4}], 'b': 5},
                           {'c': [{'c': 1, 'b': 2}, {'c': 3, 'b': 4}], 'b': 5},
                           BackwardCompatible(old='a', new='c'), '')])
def test_bw_compat_single1(capsys: CaptureFixture[str], ind: Any, outd: Any,
                           ren: BackwardCompatible, errtxt: str) -> None:
    """Test Config._bwcompat_single for case 1."""
    data = deepcopy(ind)
    # pylint: disable-next=protected-access
    Config._bwcompat_single(rename=ren, json_data=data)
    out, err = capsys.readouterr()
    assert '' == out
    assert err == errtxt
    assert data == outd


@pytest.mark.parametrize('ren',
                         [BackwardCompatible(old=cast(str, None), new='sun'),
                          BackwardCompatible(old='foo', new=cast(str, None)),
                          BackwardCompatible(old='foo', new='foo')])
def test_bw_compat_single2(capsys: CaptureFixture[str],
                           ren: BackwardCompatible) -> None:
    """Test Config._bwcompat_single for not OK case."""
    with pytest.raises(AssertionError):
        # pylint: disable-next=protected-access
        Config._bwcompat_single(rename=ren, json_data={'a': 'b'})
    out, _ = capsys.readouterr()
    assert '' == out


@pytest.mark.parametrize('ind,outd,ren,errtxt',
                         [([{'a': 2, 'b': 'c'},
                            {'a': 4, 'b': 'd'}],
                           [{'e': 2, 'b': 'c'},
                            {'e': 4, 'b': 'd'}],
                           BackwardCompatible(old='a', new='e'), ''),
                          ([{'foo': 12, 'bar': 'data'},
                            {'fff': 14, 'bar': 'other'}],
                           [{'foo': 12}, {'fff': 14, 'foo': 'other'}],
                           BackwardCompatible(old='bar', new='foo'),
                           'Inconsistent configuration:\n' +
                           'Both new config parameter foo and old bar ' +
                           'present.\nIgnoring old parameter bar\n'),
                          ([[{'a': 1, 'b': 2}, {'a': 3, 'b': 4}]],
                           [[{'a': 1, 'c': 2}, {'a': 3, 'c': 4}]],
                           BackwardCompatible(old='b', new='c'), '')])
def test_bwcompat_single_lst1(capsys: CaptureFixture[str], ind: Any, outd: Any,
                              ren: BackwardCompatible, errtxt: str) -> None:
    """Test Config._bwcompat_single_lst for case 1."""
    data = deepcopy(ind)
    # pylint: disable-next=protected-access
    Config._bwcompat_single_lst(rename=ren, json_data=data)
    out, err = capsys.readouterr()
    assert '' == out
    assert err == errtxt
    assert data == outd


class DummyCfg(Config):
    """Dummy Config for testing only."""

    def __init__(self) -> None:
        """Create a DummyCfg object."""
        self.aa = 'text'
        super().__init__(from_json_data_text='{ "aa": "text" }',
                         from_json_filename=None)

    def _backward_compatible(self) -> list[BackwardCompatible]:
        return [
            BackwardCompatible(old='a', new='x'),
            BackwardCompatible(old='b', new='y'),
            BackwardCompatible(old='c', new='z')
        ]


@pytest.mark.parametrize('ind,outd,errtxt',
                         [({'p': [{'a': 2, 'b': 'c'}, {'a': 4, 'b': 'd'}]},
                           {'p': [{'x': 2, 'y': 'c'}, {'x': 4, 'y': 'd'}]},
                           '')])
def test_ren_bak_compat(capsys: CaptureFixture[str], ind: Any, outd: Any,
                        errtxt: str) -> None:
    """Test Config._rename_backward_compatible."""
    data = deepcopy(ind)
    cfg = DummyCfg()
    # pylint: disable-next=protected-access
    cfg._rename_backward_compatible(json_data=data)
    out, err = capsys.readouterr()
    assert '' == out
    assert err == errtxt
    assert data == outd


@pytest.mark.parametrize('par, inp, key, opt, vtype, mls',
                         [('parr', [{'bar': 3, 'foo': 2},
                                    {'bar': 'b', 'foo': 3}],
                           'foo', False, int, 1),
                          ('par2', [{'bar': 'b', 'foo': 2},
                                    {'bar': 'b', 'foo': 3}],
                           'bar', False, str, 2),
                          ('par3', [{'bar': 'b', 'foo': 2},
                                    {'bar': 'b', 'foo': 3}],
                           'foobar', True, str, 0)])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_check_list_dict_ok(capsys: CaptureFixture[str], par: str, inp: Any,
                            key: str, opt: bool, vtype: type,
                            mls: int) -> None:
    """Test OK cases of Config.check_lst_dict."""
    Config.check_lst_dict(paramname=par, inp=inp, key=key, key_optional=opt,
                          valtype=vtype, min_size_list=mls)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('par, inp, key, opt, vtype, mls, msgs',
                         [('parr', {'bar': 3, 'foo': 2},
                           'foo', False, int, 0,
                           ['Error in parameter parr.',
                            'Expected list but found dict',
                            "{'bar': 3, 'foo': 2}"]),
                          ('par2', ['bar', 'b', 'foo', '2'],
                           'bar', False, int, 0,
                           ['Error in parameter par2.',
                            'Expected dict in list but found str',
                            'bar']),
                          ('par3', [{'bar': 'b', 'foo': 2},
                                    {'bar': 'b', 'foo': 3}],
                           'foobar', False, str, 0,
                           ['Error in parameter par3.',
                            'Expected key foobar not in dict in list',
                            "{'bar': 'b', 'foo': 2}"]),
                          ('par4', [{'bar': 'b', 'foo': 2},
                                    {'bar': 'b', 'foo': 3}],
                           'bar', False, int, 0,
                           ['Error in parameter par4.',
                            'Value for key bar expected to be of type',
                            'of type int but is of type str',
                            '\nb\n']),
                          ('parr', [{'bar': 3, 'foo': 2},
                                    {'bar': 'b', 'foo': 3}],
                           'foo', False, int, 3,
                           ['Error in parameter parr.',
                            'Minimum 3 elements needed in list but ' +
                            'only 2 found'])])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_check_list_dict_nok(capsys: CaptureFixture[str], par: str, inp: Any,
                             key: str, opt: bool, vtype: type, mls: int,
                             msgs: list[str]) -> None:
    """Test not OK cases of Config.check_lst_dict."""
    with pytest.raises(SystemExit):
        Config.check_lst_dict(paramname=par, inp=inp, key=key,
                              key_optional=opt, valtype=vtype,
                              min_size_list=mls)
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err


@pytest.mark.parametrize('par, inp, key, opt, vtype, mlso, mlsi, ',
                         [('parr', [{'bar': 3, 'foo': [2, 3]},
                                    {'bar': 'b', 'foo': [3, 5]}],
                           'foo', False, int, 1, 2),
                          ('par2', [{'bar': ['b', 'c'], 'foo': 2},
                                    {'bar': ['b', 'e'], 'foo': 3}],
                           'bar', False, str, 1, 2),
                          ('par3', [{'bar': 'b', 'foo': 2},
                                    {'bar': 'b', 'foo': 3}],
                           'foobar', True, str, 0, 0)])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_chk_lst_dct_lst_ok(capsys: CaptureFixture[str], par: str, inp: Any,
                            key: str, opt: bool, mlso: int, mlsi: int,
                            vtype: type) -> None:
    """Test OK cases of Config.check_lst_dict_lst."""
    Config.check_lst_dict_lst(paramname=par, inp=inp, key=key,
                              key_optional=opt, valtype=vtype,
                              min_size_outer_list=mlso,
                              min_size_inner_list=mlsi)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('par, inp, key, opt, vtype, mlso, mlsi, msgs',
                         [('parr', [{'bar': 3, 'foo': [2, 'a']},
                                    {'bar': 'b', 'foo': [3, 5]}],
                           'foo', False, int, 0, 0,
                           ['Error in parameter parr.',
                            'Value for key foo expected to be list of int',
                            'But element in list is str', "\n[2, 'a']\n"]),
                          ('par2', [{'bar': ['a', 'c'], 'foo': 2},
                                    {'bar': [1, 'xx'], 'foo': 'c'}],
                           'bar', False, str, 0, 0,
                           ['Error in parameter par2.',
                            'Value for key bar expected to be list of str',
                            'But element in list is int', "\n[1, 'xx']\n"]),
                          ('par3', [{'bar': 'b', 'foo': 2},
                                    {'bar': 'b', 'foo': 3}],
                           'foobar', False, str, 0, 0,
                           ['Error in parameter par3.',
                            'Expected key foobar not in dict in list']),
                          ('par2', [{'bar': ['b', 'c'], 'foo': 2},
                                    {'bar': ['b', 'e'], 'foo': 3}],
                           'bar', False, str, 3, 1,
                           ['Error in parameter par2.',
                            'Minimum 3 elements needed in list but ' +
                            'only 2 found.']),
                          ('par2', [{'bar': ['b', 'c'], 'foo': 2},
                                    {'bar': ['b', 'e'], 'foo': 3}],
                           'bar', False, str, 1, 3,
                           ['Error in parameter par2.',
                            'List for key bar shall be minimum 3 elements.',
                            'But it is 2 elements only.'])])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_chk_lst_dct_lst_nok(capsys: CaptureFixture[str], par: str, inp: Any,
                             key: str, opt: bool, vtype: type, mlso: int,
                             mlsi: int, msgs: list[str]) -> None:
    """Test not OK cases of Config.check_lst_dict_lst."""
    with pytest.raises(SystemExit):
        Config.check_lst_dict_lst(paramname=par, inp=inp, key=key,
                                  key_optional=opt, valtype=vtype,
                                  min_size_outer_list=mlso,
                                  min_size_inner_list=mlsi)
    out, err = capsys.readouterr()
    assert '' == out
    for msg in msgs:
        assert msg in err
