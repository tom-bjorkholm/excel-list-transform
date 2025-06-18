#! /usr/local/bin/python3
"""Functions for manipulating and rewriting a value."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


from copy import deepcopy
import sys
from typing import Optional, assert_never
from re import compile as re_compile
from re import Pattern
from re import IGNORECASE as re_IGNORECASE
from re import error as re_error
from excel_list_transform.commontypes import get_checked_type
from excel_list_transform.config_enums import RewriteKind, CaseSensitivity
from excel_list_transform.config_excel_list_transform import \
    Column, SingleRuleRewrite


def strip_value(value: str, chars: str | None,
                casehandle: CaseSensitivity) -> str:
    """Strip off characters from beginning/end of value."""
    if chars is None or chars == '':
        return value.strip()
    cha = chars
    if casehandle == CaseSensitivity.IGNORE_CASE and len(chars) > 0:
        cha = chars.lower() + chars.upper()
    return value.strip(cha)


def str_replace_value(value: str, fro: str, to: str,
                      casehandle: CaseSensitivity) -> str:
    """Replace sub-strings in value."""
    if casehandle == CaseSensitivity.MATCH_CASE:
        return value.replace(fro, to)
    fromstr = deepcopy(fro)
    pos = -1
    while True:
        ret = deepcopy(value)
        retl = ret.lower()
        froml = fromstr.lower()
        pos = retl.find(froml, pos+1)
        if pos < 0:  # from not found any more
            break
        fro_cased = value[pos:pos+len(fro)]
        value = value.replace(fro_cased, to)
    return value


def remove_from_value(value: str, chars: list[str],
                      casehandle: CaseSensitivity) -> str:
    """Remove sub-string from value."""
    if len(chars) == 0:
        return value
    for cha in chars:
        value = str_replace_value(value=value, fro=cha, to='',
                                  casehandle=casehandle)
    return value


def regex_replace_value(value: str, fro: str, to: str,
                        casehandle: CaseSensitivity) -> str:
    """Replace regexp match in value."""
    flags = re_IGNORECASE if casehandle == CaseSensitivity.IGNORE_CASE else 0
    if fro in regex_replace_value_cache:
        cached: dict[int, Pattern[str]] = regex_replace_value_cache[fro]
        if flags in cached:
            re_obj: Pattern[str] = cached[flags]
            assert re_obj is not None
            return re_obj.sub(repl=to, string=value)
    try:
        re_obj2: Pattern[str] = re_compile(pattern=fro, flags=flags)
    except re_error as exc:
        msg = 's09_rewrite_columns: invalid regular expression (regex): '
        msg += f'"{fro}"\n'
        msg += str(exc)
        print(msg, file=sys.stderr)
        sys.exit(1)
    assert re_obj2 is not None
    if fro in regex_replace_value_cache:
        regex_replace_value_cache[fro][flags] = re_obj2
    else:
        regex_replace_value_cache[fro] = {flags: re_obj2}
    return re_obj2.sub(repl=to, string=value)


regex_replace_value_cache: dict[str, dict[int, Pattern[str]]] = {}


def rewrite_value(value: Optional[str],
                  spec: SingleRuleRewrite[Column],
                  tinfo: Column) -> Optional[str]:
    """Return a value rewritten according to spec."""
    if value is None:
        return None
    assert value is not None
    assert tinfo is not None  # tinfo needed for mypy, avoid pylint unused
    ret: str = deepcopy(value)
    kind: RewriteKind = get_checked_type(spec['kind'], RewriteKind)
    casehandle: CaseSensitivity = get_checked_type(spec['case'],
                                                   CaseSensitivity)
    if kind == RewriteKind.STRIP:
        chars1 = spec['chars']
        if chars1 is not None:
            assert isinstance(chars1, str)
        return strip_value(value=ret, chars=chars1,
                           casehandle=casehandle)
    if kind == RewriteKind.REMOVECHARS:
        chars: list[str] = get_checked_type(spec['chars'], list)
        return remove_from_value(value=ret, chars=chars,
                                 casehandle=casehandle)
    fro: str = get_checked_type(spec['from'], str)
    to: str = get_checked_type(spec['to'], str)
    if kind == RewriteKind.STR_SUBSTITUTE:
        return str_replace_value(value=ret, fro=fro, to=to,
                                 casehandle=casehandle)
    if kind == RewriteKind.REGEX_SUBSTITUTE:
        return regex_replace_value(value=ret, fro=fro, to=to,
                                   casehandle=casehandle)
    assert_never(f'unexpected value kind={kind.name} in ' +
                 'rewrite_value')  # pylint: disable=unreachable,line-too-long  # pragma: no cover # noqa: E501
    assert False  # pylint: disable=unreachable,line-too-long  # pragma: no cover # noqa: E501
