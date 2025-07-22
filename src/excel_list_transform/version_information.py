#! /usr/local/bin/python3
"""Get and print version information."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from importlib.metadata import version as metadata_version
from typing import Optional, NamedTuple, TypeAlias
from datetime import date
from pypi_simple import PyPISimple
from pypi_simple.errors import NoSuchProjectError, \
    UnsupportedContentTypeError, UnsupportedRepoVersionError
from requests import HTTPError
from packaging.version import Version
from packaging.requirements import Requirement

VersionInfo: TypeAlias = dict[str, Version]
SupportExpires: TypeAlias = dict[date, str]


class AvailableVersion(NamedTuple):
    """Information on vailable versions of a package."""

    is_better_for_pyver: bool
    is_best_for_new_py: bool
    best_ver: Version
    better_ver: Version
    pkgname: str


AvailableVersions: TypeAlias = list[AvailableVersion]


class VersionInformation():
    """Get and print version information."""

    def __init__(self) -> None:
        """Create version information object."""
        self.versions: Optional[VersionInfo] = None

    def print(self, versions: Optional[VersionInfo] = None) -> None:
        """Print version information."""
        if versions is None:
            if self.versions is None:
                self.versions = self.get()
            versions = self.versions
        maxlen: int = max(map(len, versions.keys()))
        for key, value in versions.items():
            pad = '.' * (maxlen - len(key))
            print(f'{key} {pad} {str(value)}')
        self.print_info_on_new_pkgs(versions=versions)
        if 'Python' in versions and versions['Python'] < Version('3.13'):
            self._print_upgrade_instruction(versions=versions)

    def check_if_unsupported_python(self) -> None:
        """Check if Python version is no longer supported by app."""
        running_versions = self.get()
        sup_ends = self.get_app_support_expires()
        now = self._today()
        pyvers = running_versions['Python']
        txt = '\n\nYou are running an old version of Python: ' + str(pyvers)
        txt += '\nThis application no longer releases bug fixes for this old'
        txt += ' Python version.'
        txt += '\nThis application no longer releases new features for this'
        txt += ' old Python version.\n'
        for day, version in sup_ends.items():
            if now >= day:
                cmpvers = Version('.'.join(list(str(s) for s in
                                                pyvers.release[0:2])))
                if cmpvers <= Version(version):
                    print(txt)
                    self._print_upgrade_instruction(running_versions)
                    break

    @staticmethod
    def _print_upgrade_instruction(versions: VersionInfo) -> None:
        """Print instructions for how to upgrade Python and app."""
        pip = 'pip3'
        applic = list(versions.keys())[0].replace('_', '-')
        if sys.platform.lower().startswith('win') \
                or sys.platform.lower().startswith('nt'):
            pip = 'pip'
        txt = '\nYou are running an old Python version.\n'
        txt += 'Upgrade Python to a new version.\n'
        txt += '(Download Python from https://www.python.org/downloads/ )\n'
        txt += 'After installing new Python, upgrade application with\n'
        txt += pip + ' install --upgrade excel-list-transform\n'
        if applic != 'excel-list-transform':
            txt += pip + ' install --upgrade ' + applic + '\n'
        txt += '\n'
        print(txt)

    @staticmethod
    def python_version() -> Version:
        """Get running python version."""
        py_ver = '.'.join([str(s) for s in sys.version_info[0:3]])
        return Version(py_ver)

    def get(self) -> VersionInfo:
        """Get version information for module and dependencies."""
        if self.versions is not None:
            return self.versions
        versions: VersionInfo = {i: Version(metadata_version(i))
                                 for i in self.module_names()}
        versions['Python'] = self.python_version()
        self.versions = versions
        return versions

    def module_names(self) -> list[str]:
        """Get list of modules to get versions for."""
        return ['excel_list_transform', 'openpyxl', 'pylightxl',
                'XlsxWriter']

    def get_app_support_expires(self) -> SupportExpires:
        """Get which python version is no loger supported at date."""
        support_end = {date(year=2025, month=10, day=1): '3.10',
                       date(year=2025, month=12, day=1): '3.11',
                       date(year=2026, month=3, day=1): '3.12'}
        return support_end

    def _today(self) -> date:
        """Get date of today."""
        # Mockable to enable testing.
        return date.today()

    @staticmethod
    def get_available_version(pkgname: str,
                              pkgversion: Version,
                              python_version: Version) -> AvailableVersion:
        """Get information on available version in PyPi."""
        bestver = Version('0.0.1')
        betterver = Version('0.0.1')
        with PyPISimple() as client:
            try:
                page = client.get_project_page(pkgname, timeout=15.0)
            except (NoSuchProjectError, UnsupportedContentTypeError,
                    UnsupportedRepoVersionError, HTTPError):
                return AvailableVersion(False, False, betterver, betterver,
                                        pkgname=pkgname)
            for pkg in page.packages:
                if pkg.is_yanked:
                    continue
                ver = Version(pkg.version if pkg.version is not None
                              else '0.0.1')
                bestver = max(bestver, ver)
                if ver > betterver:
                    if pkg.requires_python is not None:
                        req = Requirement('python ' + pkg.requires_python)
                        if req.specifier.contains(python_version):
                            betterver = ver
                    else:
                        betterver = ver
        better_for_pyver = betterver > pkgversion
        better_for_new_py = bestver > pkgversion and bestver > betterver
        return AvailableVersion(is_better_for_pyver=better_for_pyver,
                                is_best_for_new_py=better_for_new_py,
                                best_ver=bestver, better_ver=betterver,
                                pkgname=pkgname)

    @staticmethod
    def get_available_versions(versions: VersionInfo) -> AvailableVersions:
        """Get information on available versions in PyPi."""
        py_ver = versions['Python'] if 'Python' in versions else \
            VersionInformation().get()['Python']
        ret: AvailableVersions = []
        for pkgname, pkgversion in versions.items():
            if pkgname == 'Python':
                continue
            avail: AvailableVersion = \
                VersionInformation.get_available_version(pkgname=pkgname,
                                                         pkgversion=pkgversion,
                                                         python_version=py_ver)
            ret.append(avail)
        return ret

    @staticmethod
    def print_if_better_versions(vers: AvailableVersions) -> None:
        """Print information if better versions are available."""
        better_for_this: AvailableVersions = []
        better_with_new_py: AvailableVersions = []
        for item in vers:
            if item.is_better_for_pyver:
                better_for_this.append(item)
            if item.is_best_for_new_py and item.better_ver != item.best_ver:
                better_with_new_py.append(item)
        if not (better_for_this or better_with_new_py):
            return
        maxlen: int = max(list(len(x.pkgname)
                               for x in better_for_this + better_with_new_py))
        if better_for_this:
            print('Upgraded packages are available for this python version:')
            for item in better_for_this:
                pad = '.'*(maxlen - len(item.pkgname))
                print(f'{item.pkgname} {pad} {str(item.better_ver)}')
        if better_with_new_py:
            print('Even newer packages are available if upgrading python:')
            for item in better_with_new_py:
                pad = '.'*(maxlen - len(item.pkgname))
                print(f'{item.pkgname} {pad} {str(item.best_ver)}')

    def print_info_on_new_pkgs(self,
                               versions: Optional[VersionInfo] = None) -> None:
        """Print information if better versions are available."""
        if versions is None:
            if self.versions is None:
                self.versions = self.get()
            versions = self.versions
        self.print_if_better_versions(self.get_available_versions(versions))
