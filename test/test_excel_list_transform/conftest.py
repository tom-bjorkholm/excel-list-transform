#! /usr/local/bin/python3
"""Shared pytest fixtures for the excel-list-transform tests."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

import tkinter
from contextlib import suppress
from typing import Iterator
import pytest


def tk_description() -> str:
    """Describe the Tk that the tests run on, or say there is no display.

    The exact Tk version and the windowing system are only known once
    Tk has opened a display, so without one only its major and minor
    version are told.
    """
    try:
        root = tkinter.Tk()
    except tkinter.TclError:
        return f'Tk {tkinter.TkVersion} (no display)'
    root.withdraw()
    patch_level = root.getvar('tk_patchLevel')
    system = root.call('tk', 'windowingsystem')
    root.destroy()
    return f'Tk {patch_level}, windowing system {system}'


def pytest_report_header() -> str:
    """Tell which Tcl/Tk the tests run on, to help explain Tk failures."""
    tcl_level = tkinter.Tcl().call('info', 'patchlevel')
    return f'tcl/tk: Tcl {tcl_level}, {tk_description()}'


@pytest.fixture(name='tk_root', scope='session')
def fixture_tk_root() -> Iterator[tkinter.Tk]:
    """Provide one hidden Tk root shared by the whole test session."""
    try:
        root = tkinter.Tk()
    except tkinter.TclError as error:
        pytest.skip(f'Tk has no usable display: {error}')
    root.withdraw()
    yield root
    with suppress(tkinter.TclError):
        root.destroy()
