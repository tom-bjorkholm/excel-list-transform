#! /usr/local/bin/python3
"""Class that warns user to migrate configuration file."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

import sys
from copy import deepcopy
from typing import Optional, TextIO
from config_as_json import MigrateCfgWarnHook as BaseMigrateCfgWarnHook


class MigrateCfgWarnHook(BaseMigrateCfgWarnHook):
    """Class that warns user to migrate configuration file."""

    @classmethod
    def migrate_instructions(cls) -> str:
        """Return application-specific migration instructions."""
        _ = cls
        txt = ''
        txt += 'Use "migrate-cfg" sub-command to migrate configuration '
        txt += 'to new format.\n\n'
        return deepcopy(txt)  # copy to make sure original is not manipulated.

    def default_value_provided(self, def_val_key: str) -> None:
        """Record old-style default insertion as a ROCF value."""
        self.rocf_missing_value_provided(def_val_key)

    def all_autochanges_done(self,
                             stderr_file: Optional[TextIO] = None) -> None:
        """Report collected automatic changes using current stderr."""
        err_file = sys.stderr if stderr_file is None else stderr_file
        super().all_autochanges_done(stderr_file=err_file)
