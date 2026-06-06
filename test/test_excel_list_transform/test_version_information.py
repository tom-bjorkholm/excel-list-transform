#! /usr/local/bin/python3
"""Test application version reporting configuration and output."""

# Copyright (c) 2024 - 2026 Tom Björkholm
# MIT License


from datetime import date

from packaging.version import Version
import pytest
from pytest import CaptureFixture, MonkeyPatch

from excel_list_transform.version_information import ExcelListVersionReporter


def test_reporter_config() -> None:
    """Test app-specific configuration for the version reporter."""
    reporter = ExcelListVersionReporter()
    for pkg in ['excel_list_transform', 'tableio',
                'tableio-cfg-json', 'config-as-json',
                'versionreporter']:
        assert pkg in reporter.package_names()
    assert reporter.get_app_support_expires() == {
        date(year=2025, month=10, day=1): '3.10',
        date(year=2025, month=12, day=1): '3.11',
        date(year=2026, month=3, day=1): '3.12',
        date(year=2027, month=10, day=1): '3.13'}
    assert ExcelListVersionReporter.get_main_package_name() == \
        'excel-list-transform'
    assert ExcelListVersionReporter.get_recommended_python_version() == \
        Version('3.13')


def test_version_table(capsys: CaptureFixture[str],
                       monkeypatch: MonkeyPatch) -> None:
    """Test version table output for this application's packages."""
    version_data = {'excel_list_transform': Version('1.2.3'),
                    'openpyxl': Version('4.5.6'),
                    'pylightxl': Version('7.8.9'),
                    'XlsxWriter': Version('10.11.12'),
                    'Python': Version('3.13.1')}

    def no_new_packages(_: ExcelListVersionReporter,
                        versions: dict[str, Version]) -> None:
        """Skip PyPI lookup in this output-focused test."""
        assert versions == version_data

    monkeypatch.setattr(ExcelListVersionReporter, '_print_info_on_new_pkgs',
                        no_new_packages)
    ExcelListVersionReporter().print(version_data)
    out, err = capsys.readouterr()
    assert err == ''
    assert 'excel_list_transform  1.2.3' in out
    assert 'openpyxl ............ 4.5.6' in out
    assert 'pylightxl ........... 7.8.9' in out
    assert 'XlsxWriter .......... 10.11.12' in out
    assert 'Python .............. 3.13.1' in out
    assert 'You are running an old Python version.' not in out


@pytest.mark.parametrize('platform_name, pip_cmd',
                         [('Windows', 'pip'),
                          ('windows', 'pip'),
                          ('linux', 'pip3'),
                          ('darwin', 'pip3')])
def test_print_upgrade_text(capsys: CaptureFixture[str],
                            monkeypatch: MonkeyPatch, platform_name: str,
                            pip_cmd: str) -> None:
    """Test old-Python report prints the application upgrade command."""
    version_data = {'excel_list_transform': Version('1.2.3'),
                    'Python': Version('3.12.0')}

    def no_new_packages(_: ExcelListVersionReporter,
                        versions: dict[str, Version]) -> None:
        """Skip PyPI lookup in this output-focused test."""
        assert versions == version_data

    monkeypatch.setattr(ExcelListVersionReporter, '_print_info_on_new_pkgs',
                        no_new_packages)
    monkeypatch.setattr('versionreporter.versionreporter.sys.platform',
                        platform_name)
    ExcelListVersionReporter().print(version_data)
    out, err = capsys.readouterr()
    assert err == ''
    assert 'You are running an old Python version.' in out
    assert 'Upgrade Python to a new version.' in out
    assert '(Download Python from https://www.python.org/downloads/ )' in out
    assert 'After installing new Python, upgrade application with' in out
    assert f'{pip_cmd} install --upgrade excel-list-transform' in out


@pytest.mark.parametrize('py_version, today, should_print',
                         [(Version('3.11.1'),
                           date(year=2025, month=12, day=25), True),
                          (Version('3.11.1'),
                           date(year=2025, month=9, day=25), False),
                          (Version('3.10.11'),
                           date(year=2027, month=12, day=25), True)])
def test_support_check(capsys: CaptureFixture[str], py_version: Version,
                       today: date, should_print: bool) -> None:
    """Test unsupported Python warning uses this app's support dates."""
    class MockReporter(ExcelListVersionReporter):
        """Version reporter with deterministic Python version and date."""

        def _get(self) -> dict[str, Version]:
            """Return deterministic version data for this test."""
            return {'excel_list_transform': Version('1.2.3'),
                    'Python': py_version}

        def _today(self) -> date:
            """Return deterministic current date for this test."""
            return today

    MockReporter().check_if_unsupported_python()
    out, err = capsys.readouterr()
    assert err == ''
    if should_print:
        assert 'You are running an old version of Python:' in out
        assert 'This application no longer releases bug fixes ' in out
        assert 'for this old Python version.' in out
        assert 'Upgrade Python to a new version.' in out
        assert 'After installing new Python, upgrade application with' in out
        assert ' install --upgrade excel-list-transform' in out
    else:
        assert out == ''
