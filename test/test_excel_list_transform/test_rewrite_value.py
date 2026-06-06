#! /usr/local/bin/python3
"""Test the rewriting of a single value."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from typing import Optional, cast
import pytest
from pytest import CaptureFixture
from excel_list_transform.config_enums import CaseSensitivity, RewriteKind
from excel_list_transform.rewrite_value import rewrite_value, \
    strip_value, str_replace_value, remove_from_value, regex_replace_value
from excel_list_transform.config_excel_list_transform import \
    SingleRuleRewrite


@pytest.mark.parametrize('ind, chars, caseh, outd',
                         [(' some text \t\n', '', CaseSensitivity.IGNORE_CASE,
                           'some text'),
                          (' some text \t\n', '', CaseSensitivity.MATCH_CASE,
                           'some text'),
                          (' some text \t\n', None,
                           CaseSensitivity.IGNORE_CASE, 'some text'),
                          (' some text \t\n', None,
                           CaseSensitivity.MATCH_CASE, 'some text'),
                          (' some text \t\n', ' \n',
                           CaseSensitivity.IGNORE_CASE, 'some text \t'),
                          (' some text \t\n', ' \n',
                           CaseSensitivity.MATCH_CASE, 'some text \t'),
                          ('abc', '', CaseSensitivity.IGNORE_CASE, 'abc'),
                          ('abc', '', CaseSensitivity.MATCH_CASE, 'abc'),
                          ('abc', 'a', CaseSensitivity.IGNORE_CASE, 'bc'),
                          ('abc', 'a', CaseSensitivity.MATCH_CASE, 'bc'),
                          ('abc', 'A', CaseSensitivity.IGNORE_CASE, 'bc'),
                          ('abc', 'A', CaseSensitivity.MATCH_CASE, 'abc'),
                          ('abc', 'c', CaseSensitivity.IGNORE_CASE, 'ab'),
                          ('abc', 'c', CaseSensitivity.MATCH_CASE, 'ab'),
                          ('abc', 'C', CaseSensitivity.IGNORE_CASE, 'ab'),
                          ('abc', 'C', CaseSensitivity.MATCH_CASE, 'abc'),
                          ('abc', 'b', CaseSensitivity.IGNORE_CASE, 'abc'),
                          ('abc', 'b', CaseSensitivity.MATCH_CASE, 'abc'),
                          ('abc', 'B', CaseSensitivity.IGNORE_CASE, 'abc'),
                          ('abc', 'B', CaseSensitivity.MATCH_CASE, 'abc')])
def test_strip_value(capsys: CaptureFixture[str], ind: str,
                     chars: Optional[str], caseh: CaseSensitivity,
                     outd: str) -> None:
    """Test strip of value."""
    ret = strip_value(value=ind, chars=chars, casehandle=caseh)
    out, err = capsys.readouterr()
    assert ret == outd
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, chars, caseh, outd',
                         [('abcabc', [],
                           CaseSensitivity.IGNORE_CASE, 'abcabc'),
                          ('abcabc', [],
                           CaseSensitivity.MATCH_CASE, 'abcabc'),
                          ('abcabc', ['b'],
                           CaseSensitivity.IGNORE_CASE, 'acac'),
                          ('abcabc', ['b'],
                           CaseSensitivity.MATCH_CASE, 'acac'),
                          ('abcabc', ['B'],
                           CaseSensitivity.IGNORE_CASE, 'acac'),
                          ('abcabc', ['B'],
                           CaseSensitivity.MATCH_CASE, 'abcabc'),
                          ('abcabc', ['a', 'c'],
                           CaseSensitivity.IGNORE_CASE, 'bb'),
                          ('abcabc', ['a', 'c'],
                           CaseSensitivity.MATCH_CASE, 'bb'),
                          ('abcabc', ['A', 'C'],
                           CaseSensitivity.IGNORE_CASE, 'bb'),
                          ('abcabc', ['A', 'C'],
                           CaseSensitivity.MATCH_CASE, 'abcabc')])
def test_remove_from_value(capsys: CaptureFixture[str], ind: str,
                           chars: list[str], caseh: CaseSensitivity,
                           outd: str) -> None:
    """Test remove chars from value."""
    ret = remove_from_value(value=ind, chars=chars, casehandle=caseh)
    out, err = capsys.readouterr()
    assert ret == outd
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, fro, to, caseh, outd',
                         [('tahaba', 'aha', 'c',
                           CaseSensitivity.IGNORE_CASE, 'tcba'),
                          ('tahaba', 'aha', 'c',
                           CaseSensitivity.MATCH_CASE, 'tcba'),
                          ('tahaba', 'AHA', 'c',
                           CaseSensitivity.IGNORE_CASE, 'tcba'),
                          ('tahaba', 'AHA', 'c',
                           CaseSensitivity.MATCH_CASE, 'tahaba'),
                          ('ahaba', 'aha', 'c',
                           CaseSensitivity.IGNORE_CASE, 'cba'),
                          ('ahaba', 'aha', 'c',
                           CaseSensitivity.MATCH_CASE, 'cba'),
                          ('ahaba', 'AHA', 'c',
                           CaseSensitivity.IGNORE_CASE, 'cba'),
                          ('ahaba', 'AHA', 'c',
                           CaseSensitivity.MATCH_CASE, 'ahaba'),
                          ('tahaha', 'aha', 'c',
                           CaseSensitivity.IGNORE_CASE, 'tcha'),
                          ('tahaba', 'aha', 'c',
                           CaseSensitivity.MATCH_CASE, 'tcba'),
                          ('tahaha', 'AHA', 'a',
                           CaseSensitivity.IGNORE_CASE, 'taha'),
                          ('tahaha', 'AHA', 'a',
                           CaseSensitivity.MATCH_CASE, 'tahaha'),
                          ('tahaha', 'aha', 'a',
                           CaseSensitivity.MATCH_CASE, 'taha')])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_str_replace_value(capsys: CaptureFixture[str], ind: str, fro: str,
                           to: str, caseh: CaseSensitivity, outd: str) -> None:
    """Test replace substring in value."""
    ret = str_replace_value(value=ind, fro=fro, to=to, casehandle=caseh)
    out, err = capsys.readouterr()
    assert ret == outd
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('ind, fro, to, caseh, outd',
                         [('ahahah', '^ah', 'b',
                           CaseSensitivity.IGNORE_CASE, 'bahah'),
                          ('ahahah', '^AH', 'b',
                           CaseSensitivity.IGNORE_CASE, 'bahah'),
                          ('ahahah', '^ah', 'b',
                           CaseSensitivity.MATCH_CASE, 'bahah'),
                          ('ahahah', '^AH', 'b',
                           CaseSensitivity.MATCH_CASE, 'ahahah')])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_reg_replace_value(capsys: CaptureFixture[str], ind: str, fro: str,
                           to: str, caseh: CaseSensitivity, outd: str) -> None:
    """Test replace regex in value for OK cases."""
    ret = regex_replace_value(value=ind, fro=fro, to=to, casehandle=caseh)
    out, err = capsys.readouterr()
    assert ret == outd
    assert '' == err
    assert '' == out


def test_regex_repl_val_nok(capsys: CaptureFixture[str]) -> None:
    """Test replace regex in value for not OK case."""
    with pytest.raises(SystemExit):
        _ = regex_replace_value(value='abc', fro='^+4607', to='+467',
                                casehandle=CaseSensitivity.IGNORE_CASE)
    out, err = capsys.readouterr()
    assert 'invalid regular expression (regex)' in err
    assert '"^+4607"' in err
    assert '' == out


@pytest.mark.parametrize('ind, spec, outd',
                         [('adcd', {'kind': RewriteKind.STRIP,
                                    'chars': 'd',
                                    'case': CaseSensitivity.IGNORE_CASE},
                           'adc'),
                          ('adcd', {'kind': RewriteKind.REMOVECHARS,
                                    'chars': ['d'],
                                    'case': CaseSensitivity.IGNORE_CASE},
                           'ac'),
                          ('adcd', {'kind': RewriteKind.STR_SUBSTITUTE,
                                    'from': 'd', 'to': 'bx',
                                    'case': CaseSensitivity.IGNORE_CASE},
                           'abxcbx'),
                          ('adad', {'kind': RewriteKind.REGEX_SUBSTITUTE,
                                    'from': '^a', 'to': 'bx',
                                    'case': CaseSensitivity.IGNORE_CASE},
                           'bxdad'),
                          (None, {'kind': RewriteKind.REGEX_SUBSTITUTE,
                                  'from': '^a', 'to': 'bx',
                                  'case': CaseSensitivity.IGNORE_CASE},
                           None)])
def test_rewrite_value_ok(capsys: CaptureFixture[str], ind: Optional[str],
                          spec: SingleRuleRewrite[int],
                          outd: Optional[str]) -> None:
    """Test rewrite_value for OK cases."""
    ret = rewrite_value(value=ind, spec=spec, tinfo=2)
    out, err = capsys.readouterr()
    assert ret == outd
    assert '' == err
    assert '' == out


def test_rewrite_value_nok(capsys: CaptureFixture[str]) -> None:
    """Test rewrite_value for not OK case."""
    ind = 'abc'
    spec = {'kind': 'foo', 'chars': 'a',
            'case': CaseSensitivity.IGNORE_CASE}
    with pytest.raises(AssertionError):
        _ = rewrite_value(value=ind, spec=cast(SingleRuleRewrite[int], spec),
                          tinfo=2)
    out, _ = capsys.readouterr()
    assert '' == out
