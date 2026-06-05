#! /usr/local/bin/python3
"""Class that warns user to migrate configuration file."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

from copy import deepcopy
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
