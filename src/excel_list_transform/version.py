#! /usr/local/bin/python3
"""Get and print version information."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License


import sys
from importlib.metadata import version as metadata_version
from typing import Optional, TypeAlias

VersionInfo: TypeAlias = dict[str, str]


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
            pip = 'pip3'
            applic = list(versions.keys())[0].replace('_', '-')
            if sys.platform.startswith('win') \
                    or sys.platform.startswith('Win'):
                pip = 'pip'
            txt = '\nYou are running an old Python version.\n'
            txt += 'Upgrade Python to a new version.\n'
            txt += '(Download Python from https://www.python.org/downloads/)\n'
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
