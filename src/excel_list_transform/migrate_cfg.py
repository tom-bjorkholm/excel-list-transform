#! /usr/local/bin/python3
"""Migrate configuration files to the current JSON format."""

# Copyright (c) 2024-2025 Tom Björkholm
# MIT License

import sys
from config_as_json import migrate_cfg as config_as_json_migrate_cfg
from excel_list_transform.config_match import MATCH_CONFIGS


def migrate_cfg(infile: str, outfile: str) -> int:
    """Migrate configuration to newest format.

    Read in input configuration file using backward compatibility.
    Write out configuration using newest format.
    @param infile  Name of input configuration file.
    @param outfile Name of output configuration file.
    """
    return config_as_json_migrate_cfg(infile=infile, outfile=outfile,
                                      config_class=MATCH_CONFIGS,
                                      stderr_file=sys.stderr)
