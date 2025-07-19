#! /usr/local/bin/python3
"""Test recode command and recode function."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


from datetime import date
from packaging.version import Version
import pytest
from excel_list_transform.version_information import VersionInformation, \
    VersionInfo


@pytest.mark.parametrize('data, texts',
                         [({'Ab': '1.2.3', 'longer': '4.5', 'short': '6.7.8'},
                           ['Ab .... 1.2.3', 'longer  4.5', 'short . 6.7.8']),
                          ({'really_long': '9.8.7.123', 'another': '1.0'},
                           ['really_long  9.8.7.123', 'another .... 1.0'])
                          ])
@pytest.mark.parametrize('asarg, usearg', [[True, False], [True, False]])
def test_version_print1(capsys, data, texts, asarg, usearg):
    """Test VersionInformation.print."""
    class Derived(VersionInformation):
        """Derived class for fake get method."""

        def get(self) -> VersionInfo:
            """Get fake version info."""
            if asarg:
                return {key: Version(val) for key, val in data.items()}
            return {'Foo': Version('3.1415')}

    vers = Derived()
    vers.print(data if usearg else None)
    out, err = capsys.readouterr()
    assert err == ''
    if not usearg and not asarg:
        assert out == 'Foo  3.1415\n'
    else:
        for i in texts:
            assert i in out
        rows = out.split('\n')
        for row in rows:
            print(row)
            words = row.split()
            if len(words) > 1:
                assert data[words[0]] == words[-1]


@pytest.mark.parametrize('calls', [0, 1, 2, 8])
def test_version_print2(capsys, calls):
    """Test VersionInformation.print loops."""
    num = 0

    class Derived(VersionInformation):
        """Derived class for fake get method."""

        def get(self) -> VersionInfo:
            """Get fake version info."""
            nonlocal num
            num += 1
            return {'Foo': Version('3.1415')}

    vers = Derived()
    for _ in range(calls):
        vers.print()
    out, err = capsys.readouterr()
    assert err == ''
    assert out == 'Foo  3.1415\n'*calls
    assert num == min(calls, 1)


@pytest.mark.parametrize('vers', [['1', '2.3', '4.5.6', '7b'],
                                  ['4.5', '2.7', '4.5', '1.9.0']])
@pytest.mark.parametrize('pyt', [[3, 10, 7], [3, 13, 1]])
def test_version_get1(monkeypatch, capsys, vers, pyt):
    """Test VersionInformation.get with patched getters."""
    metaversnum = 0

    def metavers(modulename: str) -> str:
        """Patched metadata.version."""
        assert modulename is not None
        nonlocal metaversnum
        ret = vers[metaversnum]
        metaversnum += 1
        return ret

    mod = 'excel_list_transform.version_information.'
    monkeypatch.setattr(mod + 'sys.version_info', pyt)
    monkeypatch.setattr(mod + 'metadata_version', metavers)
    versionobj = VersionInformation()
    ret = versionobj.get()
    mods = versionobj.module_names()
    out, err = capsys.readouterr()
    assert '' == err
    assert '' == out
    for i, mod in enumerate(mods):
        assert Version(vers[i]) == ret[mod]


@pytest.mark.parametrize('numcalls', [0, 1, 2, 17])
def test_version_get2(monkeypatch, capsys, numcalls):
    """Test VersionInformation.get with patched getters."""
    calls = 0

    def metavers(modulename: str) -> str:
        """Patched metadata.version."""
        assert modulename is not None
        nonlocal calls
        calls += 1
        return '1.8.2'

    mod = 'excel_list_transform.version_information.'
    monkeypatch.setattr(mod + 'metadata_version', metavers)
    versionobj = VersionInformation()
    for _ in range(numcalls):
        ret = versionobj.get()
        mods = versionobj.module_names()
        for mod in mods:
            assert ret[mod] == Version('1.8.2')
        assert VersionInformation.python_version() == ret['Python']
        assert len(mods) == calls
    if 0 == numcalls:
        assert 0 == calls
    out, err = capsys.readouterr()
    assert '' == err
    assert '' == out


@pytest.mark.parametrize('os, pip',
                         [('Windows', 'pip'),
                          ('windows', 'pip'),
                          ('linux', 'pip3'),
                          ('darwin', 'pip3')])
def test_print_upgrade(capsys, monkeypatch, os, pip):
    """Test _print_upgrade_instructions."""
    mod = 'excel_list_transform.version_information.sys.platform'
    monkeypatch.setattr(mod, os)
    VersionInformation()._print_upgrade_instruction({'aha': '0.1'})  # pylint: disable=protected-access # noqa: E501
    out, err = capsys.readouterr()
    assert '' == err
    assert pip + ' install --upgrade aha' in out
    assert pip + ' install --upgrade excel-list-transform' in out


def test_version_print_old_p(capsys, monkeypatch):
    """Test version print with old Python."""
    mod = 'excel_list_transform.version_information.sys.version_info'
    monkeypatch.setattr(mod, [3, 11, 1])
    VersionInformation().print()
    out, err = capsys.readouterr()
    assert '' == err
    assert 'excel_list_transform  ' in out
    assert 'openpyxl ............ ' in out
    assert 'pylightxl ........... ' in out
    assert 'XlsxWriter .......... ' in out
    assert 'Python .............. 3.11.1' in out
    assert 'You are running an old Python version.' in out
    assert 'Upgrade Python to a new version.' in out
    assert '(Download Python from https://www.python.org/downloads/ )' in out
    assert 'After installing new Python, upgrade application with' in out
    assert ' install --upgrade ' in out


@pytest.mark.parametrize('ver, dat, errprint',
                         [((3, 11, 1, 0, 0),
                           date(year=2026, month=12, day=25), True),
                          ((3, 11, 1, 0, 0),
                           date(year=2024, month=12, day=25), False),
                          ((3, 10, 11, 75, 0),
                           date(year=2027, month=12, day=25), True)])
def test_version_check_if_u(capsys, monkeypatch, ver, dat, errprint):
    """Test version check if unsupported python widh old Python."""
    mod = 'excel_list_transform.version_information.sys.version_info'
    monkeypatch.setattr(mod, ver)

    class MockVersion1(VersionInformation):
        """Version with mocked today."""

        def _today(self):
            """Mock today."""
            return dat

    MockVersion1().check_if_unsupported_python()
    out, err = capsys.readouterr()
    assert '' == err
    if errprint:
        assert 'You are running an old version of Python:' in out
        assert 'This application no longer releases bug fixes ' in out
        assert 'for this old Python version.' in out
        assert 'Upgrade Python to a new version.' in out
        assert '(Download Python from https://www.python.org/downloads' in out
        assert 'After installing new Python, upgrade application with' in out
        assert ' install --upgrade ' in out
    else:
        assert '' == out
