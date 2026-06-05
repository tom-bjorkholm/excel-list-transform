#! /usr/local/bin/python3
"""Config-as-json match rules for selecting application config classes."""

# Copyright (c) 2026 Tom Björkholm
# MIT License

from config_as_json import JsonValueMatcher, MatchConfig, MatchConfigSeq
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.config_xls_list_transf_name import \
    ConfigXlsListTransfName
from excel_list_transform.config_xls_list_transf_num import \
    ConfigXlsListTransfNum


MATCH_CONFIGS: MatchConfigSeq = [
    MatchConfig(match_func=JsonValueMatcher('column_ref',
                                            ColumnRef.BY_NAME.name),
                config_class=ConfigXlsListTransfName),
    MatchConfig(match_func=JsonValueMatcher('column_ref',
                                            ColumnRef.BY_NUMBER.name),
                config_class=ConfigXlsListTransfNum)]
"""Configuration class selection rules for JSON configuration files."""
