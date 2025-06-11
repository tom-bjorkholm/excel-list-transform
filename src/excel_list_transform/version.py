#! /usr/local/bin/python3
"""Get and print version information."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from importlib.metadata import version as metadata_version
from typing import Optional
from datetime import date

type VersionInfo = dict[str, str]
type SupportExpires = dict[date, str]


class Version():
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
            print(f'{key} {pad} {value}')
        if 'Python' in versions and versions['Python'] < '3.13':
            self._print_upgrade_instruction(versions=versions)

    def check_if_unsupported_python(self) -> None:
        """Check if Python version is no longer supported by app."""
        running_versions = self.get()
        sup_ends = self.get_app_support_expires()
        now = self._today()
        pyvers = running_versions['Python']
        txt = '\n\nYou are running an old version of Python: ' + pyvers
        txt += '\nThis application no longer releases bug fixes for this old'
        txt += ' Python version.'
        txt += '\nThis application no longer releases new features for this'
        txt += ' old Python version.\n'
        for day, version in sup_ends.items():
            if now >= day:
                cmpvers = pyvers[0:len(version)]
                if cmpvers <= version:
                    print(txt)
                    self._print_upgrade_instruction(running_versions)
                    break

    @staticmethod
    def _print_upgrade_instruction(versions: VersionInfo) -> None:
        """Print instructions for how to upgrade Python and app."""
        pip = 'pip3'
        applic = list(versions.keys())[0].replace('_', '-')
        if sys.platform.startswith('win') \
                or sys.platform.startswith('Win'):
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

    def get(self) -> VersionInfo:
        """Get version information for module and dependencies."""
        if self.versions is not None:
            return self.versions
        versions: VersionInfo = {i: metadata_version(i)
                                 for i in self.module_names()}
        versions['Python'] = '.'.join([str(i) for i in sys.version_info])
        self.versions = versions
        return versions

    def module_names(self) -> list[str]:
        """Get list of modules to get versions for."""
        return ['excel_list_transform', 'openpyxl', 'pylightxl',
                'XlsxWriter']

    def get_app_support_expires(self) -> SupportExpires:
        """Get which python version is no loger supported at date."""
        support_end = {date(year=2025, month=10, day=1): '3.10',
                       date(year=2025, month=12, day=1): '3.11'}
        return support_end

    def _today(self) -> date:
        """Get date of today."""
        # Mockable to enable testing.
        return date.today()
