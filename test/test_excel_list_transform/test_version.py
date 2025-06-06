#! /usr/local/bin/python3
"""Test recode command and recode function."""

import sys
import pytest
from excel_list_transform.version import Version, VersionInfo


@pytest.mark.parametrize('data, texts',
                         [({'Ab': '1.2.3', 'longer': '4.5', 'short': '6.7.8'},
                           ['Ab .... 1.2.3', 'longer  4.5', 'short . 6.7.8']),
                          ({'really_long': '9.8.7.123', 'another': '1a'},
                           ['really_long  9.8.7.123', 'another .... 1a'])
                          ])
@pytest.mark.parametrize('asarg, usearg', [[True, False], [True, False]])
def test_version_print1(capsys, data, texts, asarg, usearg):
    """Test Version.print."""
    class Derived(Version):
        """Derived class for fake get method."""

        def get(self) -> VersionInfo:
            """Get fake version info."""
            if asarg:
                return data
            return {'Foo': '3.1415'}

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
            words = row.split()
            if len(words) > 1:
                assert data[words[0]] == words[-1]


@pytest.mark.parametrize('calls', [0, 1, 2, 8])
def test_version_print2(capsys, calls):
    """Test Version.print loops."""
    num = 0

    class Derived(Version):
        """Derived class for fake get method."""

        def get(self) -> VersionInfo:
            """Get fake version info."""
            nonlocal num
            num += 1
            return {'Foo': '3.1415'}

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
    """Test Version.get with patched getters."""
    metaversnum = 0

    def metavers(modulename: str) -> str:
        """Patched metadata.version."""
        assert modulename is not None
        nonlocal metaversnum
        ret = vers[metaversnum]
        metaversnum += 1
        return ret

    mod = 'excel_list_transform.version.'
    monkeypatch.setattr(mod + 'sys.version_info', pyt)
    monkeypatch.setattr(mod + 'metadata_version', metavers)
    versionobj = Version()
    ret = versionobj.get()
    mods = versionobj.module_names()
    out, err = capsys.readouterr()
    assert '' == err
    assert '' == out
    for i, mod in enumerate(mods):
        assert vers[i] == ret[mod]


@pytest.mark.parametrize('numcalls', [0, 1, 2, 17])
def test_version_get2(monkeypatch, capsys, numcalls):
    """Test Version.get with patched getters."""
    calls = 0

    def metavers(modulename: str) -> str:
        """Patched metadata.version."""
        assert modulename is not None
        nonlocal calls
        calls += 1
        return '1.8.2'

    mod = 'excel_list_transform.version.'
    monkeypatch.setattr(mod + 'metadata_version', metavers)
    versionobj = Version()
    for _ in range(numcalls):
        ret = versionobj.get()
        mods = versionobj.module_names()
        for mod in mods:
            assert ret[mod] == '1.8.2'
        assert '.'.join(map(str, sys.version_info)) == ret['Python']
        assert len(mods) == calls
    if 0 == numcalls:
        assert 0 == calls
    out, err = capsys.readouterr()
    assert '' == err
    assert '' == out
