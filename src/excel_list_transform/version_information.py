#! /usr/local/bin/python3
"""Report application and dependency version information."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


from datetime import date
from packaging.version import Version
from versionreporter import VersionReporter


class ExcelListVersionReporter(VersionReporter):
    """Report version information for excel-list-transform."""

    def package_names(self) -> list[str]:
        """Return package names included in the version report."""
        return ['excel_list_transform', 'tableio',
                'tableio-cfg-json', 'config-as-json',
                'versionreporter', 'edit-cfg-json',
                'edit-cfg-json-tk']

    def get_app_support_expires(self) -> dict[date, str]:
        """Return Python-version support cutoffs for this application."""
        return {date(year=2026, month=3, day=1): '3.12',
                date(year=2027, month=11, day=1): '3.13',
                date(year=2028, month=11, day=1): '3.14'}

    @classmethod
    def recommended_python(cls) -> Version:
        """Return the recommended Python version for this application."""
        return Version('3.15')

    @classmethod
    def get_main_package_name(cls) -> str:
        """Return the PyPI package name of this application."""
        return 'excel-list-transform'
