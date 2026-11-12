#! /usr/local/bin/python3
"""Test that the editcfg sub-command opens the Tk configuration editor.

The editor itself is tested in its own package, so these tests only
check the integration: that editcfg reaches the Tk window, and that the
values of the file and the descriptions of this application are shown
in it. The event loop of the window is replaced by a function that
withdraws the window before Tk has drawn it, records what it holds and
destroys it, so no window is ever on the screen. On a machine without
a display (such as a build server) the tests are skipped.
"""

# Copyright (c) 2026 Tom Björkholm
# MIT License

import json
import tkinter
from collections.abc import Callable, Iterator
from dataclasses import dataclass, field
from pathlib import Path
import pytest
from pytest import CaptureFixture, MonkeyPatch
from edit_cfg_json_tk.tk_ui import tk_can_run
from test_excel_list_transform.test_edit_cfg import example_cfg
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.edit_cfg import NOT_SAVED
from excel_list_transform.edit_descriptions import cfg_descriptions
from excel_list_transform.transform_cmd import transform_cmd

pytestmark = pytest.mark.skipif(not tk_can_run(),
                                reason='Tk cannot open a window here')
"""Skip every test of this file where Tk has no display to open on."""

MARKER = 4711
"""Value written to the file, to be recognized in the editor window."""


@dataclass
class Shown:
    """Texts the editor window showed before the test destroyed it."""

    windows: int = 0
    """Number of times the event loop of a Tk window was started."""
    values: list[str] = field(default_factory=list)
    """Text of every entry field in the layout."""
    labels: list[str] = field(default_factory=list)
    """Text of every label in the layout."""


def all_widgets(widget: tkinter.Misc) -> Iterator[tkinter.Misc]:
    """Return a widget and every widget below it."""
    yield widget
    for child in widget.winfo_children():
        yield from all_widgets(child)


def in_layout(widget: tkinter.Misc) -> bool:
    """Return whether the widget and all its parents are in the layout.

    The editor hides a widget by taking it out of the layout, so a
    widget the user could see has a geometry manager, and so have all
    its parents up to (but not including) the window.
    """
    while widget.winfo_parent():
        if not widget.winfo_manager():
            return False
        widget = widget.nametowidget(widget.winfo_parent())
    return True


def record(window: tkinter.Tk, shown: Shown) -> None:
    """Record the texts of the entries and labels shown in a window."""
    shown.windows += 1
    for widget in filter(in_layout, all_widgets(window)):
        if isinstance(widget, tkinter.Entry):
            shown.values.append(widget.get())
        elif isinstance(widget, tkinter.Label):
            shown.labels.append(str(widget.cget('text')))


def fake_mainloop(shown: Shown) -> Callable[[tkinter.Tk, int], None]:
    """Return an event loop that records a window and then closes it.

    The window is withdrawn at once, before Tk has had any idle time to
    draw it, and is always destroyed, also when recording fails, so a
    failing test leaves no Tk behind for the next test.
    """
    def mainloop(window: tkinter.Tk, n: int = 0) -> None:
        """Record what the window shows instead of handing it to a user."""
        assert n == 0
        window.withdraw()
        try:
            record(window, shown)
        finally:
            window.destroy()
    return mainloop


def marked_cfg(folder: Path, ref: ColumnRef) -> Path:
    """Generate an example configuration file holding the marker value."""
    cfg_file = example_cfg(folder, ref)
    data = json.loads(cfg_file.read_text(encoding='utf-8'))
    data['max_column_read'] = MARKER
    cfg_file.write_text(json.dumps(data), encoding='utf-8')
    return cfg_file


@pytest.mark.parametrize('ref', list(ColumnRef))
@pytest.mark.parametrize('ui_args', [[], ['--ui', 'tk']])
def test_editcfg_tk_shows(capsys: CaptureFixture[str],
                          monkeypatch: MonkeyPatch, tmp_path: Path,
                          ref: ColumnRef, ui_args: list[str]) -> None:
    """Test that editcfg shows a file value and a description in Tk.

    Without --ui the editor with the highest priority is chosen, which
    is Tk on a machine with a display. The home folder is moved to an
    empty folder, so that settings files of the user running the tests
    cannot choose another editor or hide the descriptions.
    """
    monkeypatch.setenv('HOME', str(tmp_path))
    monkeypatch.setenv('USERPROFILE', str(tmp_path))
    cfg_file = marked_cfg(tmp_path, ref)
    capsys.readouterr()
    shown = Shown()
    monkeypatch.setattr(tkinter.Tk, 'mainloop', fake_mainloop(shown))
    transform_cmd(['editcfg', '-i', str(cfg_file)] + ui_args)
    out, err = capsys.readouterr()
    assert '' == err
    assert NOT_SAVED + '\n' == out
    assert shown.windows == 1
    assert str(MARKER) in shown.values
    description = cfg_descriptions(ref)[('column_ref',)]
    assert any(description in label for label in shown.labels)
