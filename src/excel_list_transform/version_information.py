#! /usr/local/bin/python3
"""Report application and dependency version information."""

# Copyright (c) 2024-2026 Tom Björkholm
# MIT License


from datetime import date

from versionreporter import VersionReporter


class ExcelListVersionReporter(VersionReporter):
    """Report version information for excel-list-transform."""

    def package_names(self) -> list[str]:
        """Return package names included in the version report."""
        return ['excel_list_transform', 'tableio',
                'tableio-cfg-json', 'config-as-json',
                'versionreporter']

    def get_app_support_expires(self) -> dict[date, str]:
        """Return Python-version support cutoffs for this application."""
        return {date(year=2025, month=10, day=1): '3.10',
                date(year=2025, month=12, day=1): '3.11',
                date(year=2026, month=3, day=1): '3.12',
                date(year=2027, month=10, day=1): '3.13'}

    @classmethod
    def get_main_package_name(cls) -> str:
        """Return the PyPI package name of this application."""
        return 'excel-list-transform'
