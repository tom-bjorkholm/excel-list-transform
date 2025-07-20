#! /usr/local/bin/python3
"""Test recode command and recode function."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


from datetime import date
from typing import NamedTuple, Optional
from packaging.version import Version
from pypi_simple import ProjectPage, DistributionPackage, \
    NoSuchProjectError
import pytest
from excel_list_transform.version_information import VersionInformation, \
    VersionInfo, AvailableVersion


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
    monkeypatch.setattr(mod, (3, 11, 1, 0, 0))
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
                           date(year=2025, month=12, day=25), True),
                          ((3, 11, 1, 0, 0),
                           date(year=2025, month=9, day=25), False),
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


@pytest.mark.parametrize('ver, res',
                         [((3, 11, 2, 0, 0), '3.11.2'),
                          ((3, 13, 3, 0, 0), '3.13.3'),
                          ((3, 14, 4, 0, 0), '3.14.4'),
                          ((3, 10, 5, 0, 0), '3.10.5')])
def test_python_version(capsys, monkeypatch, ver, res):
    """Test VersionsInformation.python_version."""
    mod = 'excel_list_transform.version_information.sys.version_info'
    monkeypatch.setattr(mod, ver)
    ret = VersionInformation.python_version()
    out, err = capsys.readouterr()
    assert Version(res) == ret
    assert res == str(ret)
    assert '' == out
    assert '' == err


class MockReturn(NamedTuple):
    """What should packages in mocked page be like."""

    version: str
    is_yanked: bool
    req_python: Optional[str]


type PageReturn = list[MockReturn]
type PageReturns = list[PageReturn]

pagereturns: PageReturns = []


def mocked_get_project_page(_, project: str,
                            timeout: float | tuple[float, float] | None
                            = None, accept: Optional[str] = None,
                            headers: Optional[dict[str, str]]
                            = None) -> ProjectPage:
    """Mock PyPISimple.get_project_page."""
    packages: list[DistributionPackage] = []
    assert timeout is not None
    assert accept is None
    assert headers is None
    if not pagereturns or pagereturns[-1] is None:
        raise NoSuchProjectError(project=project, url='')
    mockvals: PageReturn = pagereturns.pop()
    for mockret in mockvals:
        packages.append(DistributionPackage(filename=f'{project}.whl',
                                            url='',
                                            project=project,
                                            package_type='wheel',
                                            version=mockret.version,
                                            digests={},
                                            requires_python=mockret.req_python,
                                            has_sig=False,
                                            is_yanked=mockret.is_yanked))
    return ProjectPage(project=project, packages=packages,
                       repository_version=None, last_serial=None)


@pytest.mark.parametrize('page,name,ver,pyver,res',
                         [([MockReturn(version='1.0', is_yanked=False,
                                       req_python=''),
                            MockReturn(version='10.1', is_yanked=False,
                                       req_python='>= 3.0')],
                           'abc', '2.0', '3.10',
                           AvailableVersion(is_better_for_pyver=True,
                                            is_best_for_new_py=False,
                                            best_ver=Version('10.1'),
                                            better_ver=Version('10.1'),
                                            pkgname='abc')),
                          ([MockReturn(version='3.0', is_yanked=False,
                                       req_python=''),
                            MockReturn(version='10.1', is_yanked=False,
                                       req_python='>= 3.12')],
                           'ghi', '2.0', '3.10',
                           AvailableVersion(is_better_for_pyver=True,
                                            is_best_for_new_py=True,
                                            best_ver=Version('10.1'),
                                            better_ver=Version('3.0'),
                                            pkgname='ghi')),
                          ([MockReturn(version='3.0', is_yanked=False,
                                       req_python=''),
                            MockReturn(version='10.1', is_yanked=True,
                                       req_python='>= 3.12')],
                           'jkl', '2.0', '3.10',
                           AvailableVersion(is_better_for_pyver=True,
                                            is_best_for_new_py=False,
                                            best_ver=Version('3.0'),
                                            better_ver=Version('3.0'),
                                            pkgname='jkl')),
                          ([MockReturn(version='3.0', is_yanked=False,
                                       req_python='>= 3.12'),
                            MockReturn(version='10.1', is_yanked=False,
                                       req_python='>= 3.12')],
                           'mno', '4.0', '3.10',
                           AvailableVersion(is_better_for_pyver=False,
                                            is_best_for_new_py=True,
                                            best_ver=Version('10.1'),
                                            better_ver=Version('0.0.1'),
                                            pkgname='mno')),
                          ([MockReturn(version='1.0', is_yanked=False,
                                       req_python='>= 3.12'),
                            MockReturn(version='3.1', is_yanked=False,
                                       req_python='>= 3.12')],
                           'pqr', '4.0', '3.10',
                           AvailableVersion(is_better_for_pyver=False,
                                            is_best_for_new_py=False,
                                            best_ver=Version('3.1'),
                                            better_ver=Version('0.0.1'),
                                            pkgname='pqr')),
                          ([MockReturn(version='3.0', is_yanked=False,
                                       req_python=None),
                            MockReturn(version='11.1', is_yanked=False,
                                       req_python=None)],
                           'stu', '4.0', '3.10',
                           AvailableVersion(is_better_for_pyver=True,
                                            is_best_for_new_py=False,
                                            best_ver=Version('11.1'),
                                            better_ver=Version('11.1'),
                                            pkgname='stu')),
                          (None, 'def', '2.0', '3.12',
                           AvailableVersion(is_better_for_pyver=False,
                                            is_best_for_new_py=False,
                                            best_ver=Version('0.0.1'),
                                            better_ver=Version('0.0.1'),
                                            pkgname='def'))])
def test_get_avail_version(capsys,  # pylint: disable=too-many-arguments,too-many-positional-arguments) # noqa: E501
                           monkeypatch, page, name, ver, pyver, res):
    """Test VersionInformation.get_available_version."""
    global pagereturns  # pylint: disable=global-statement
    pagereturns = [page]
    mod = 'excel_list_transform.version_information.PyPISimple.'
    monkeypatch.setattr(mod + 'get_project_page', mocked_get_project_page)
    pkgversion = Version(ver)
    pyversion = Version(pyver)
    ret = VersionInformation.get_available_version(pkgname=name,
                                                   pkgversion=pkgversion,
                                                   python_version=pyversion)
    out, err = capsys.readouterr()
    assert res == ret
    assert '' == out
    assert '' == err
