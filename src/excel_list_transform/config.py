#! /usr/local/bin/python3
"""Base class to read and write configurations."""

# Copyright (c) 2024 Tom Björkholm
# MIT License


from copy import deepcopy
import json
import sys
import csv
from typing import Any, Optional, Type, TypeVar, Mapping, NamedTuple, Callable
from enum import Enum, IntEnum
from excel_list_transform.str_to_enum import string_to_enum_best_match
from excel_list_transform.config_enums import RewriteKind
from excel_list_transform.file_must_exist import file_must_exist


Keya = TypeVar('Keya', str, Enum, RewriteKind)


class ConfigEncoder(json.JSONEncoder):
    """Encoder for enumerations in config."""

    def default(self, o: object) -> object:
        """Define default encoding of enumerations."""
        if isinstance(o, (Enum, IntEnum)):
            return str(o.name)
        return super().default(o)


class ConfigBadJson(json.JSONDecodeError):
    """Exception class for bad JSON in configuration."""


def over_ride_needed(stri: str) -> Any:
    """Tell programmer override is needed."""
    if stri is not None:
        msg = 'Override of Config.parse_converters needed.'
        raise NotImplementedError(msg)
    return 42


ParseConverter = NamedTuple('ParseConverter', [('result_type', type),
                                               ('func', Callable[..., Any]),
                                               ('args', dict[str, Any])])


class Config():
    """Base class for classes that read and write configurations.

    Derive from this class to and have the configuration parameters
    as instance member variables of the derived class.
    """

    def __init__(self, from_json_data_text: Optional[str],
                 from_json_filename: Optional[str]) -> None:
        """Construct configuration base class.

        Derived class __init__ must create all object variables
        before calling super().__init__()
        Derived class __init__ usually does checking after
        call to super().__init__()
        """
        self_keys = [i for i in vars(self).keys() if not
                     callable(getattr(self, i)) and not i.startswith('_')]
        if not self_keys:
            msg = 'No object variables in object of class derived from '
            msg += 'Config. (Create object variables in __init__ before '
            msg += 'calling super().__init__().)'
            raise AttributeError(msg)
        unchecked = getattr(self, '_unchecked_dicts', None)
        if unchecked is None:
            self._unchecked_dicts: list[str] = []
        elif not isinstance(unchecked, list):
            msg = '_unchecked_dicts must be a list'
            raise TypeError(msg)
        self._hook_dict = self.parse_converters()
        if from_json_data_text is not None:
            self.parse_json(from_json_data_text)
        elif from_json_filename is not None:
            self.read(from_json_filename)

    def parse_converters(self) -> dict[str, ParseConverter]:
        """Get converters for use when parsing JSON.

        Override in derived class.
        Return None if no conversions.
        Return dict of dict for use in json decoder hook.
        Structure of return value shall be:
        {key: {'result type': res_type, 'func': function,
        'args': {arg_name: arg_value}}}.
        """
        return {'in_type': ParseConverter(result_type=int,
                                          func=over_ride_needed,
                                          args={})}

    @staticmethod
    def check_key_match(expected_keys: list[str],
                        j_keys: list[str],
                        ok_to_use_defaults: bool) -> None:
        """Check if keys in imported JSON match exptected keys."""
        if not ok_to_use_defaults:
            for i in expected_keys:
                if i not in j_keys:
                    errmsg = f'No value for {i} in JSON data'
                    print(errmsg, file=sys.stderr)
                    raise KeyError(errmsg)
        for i in j_keys:
            if i not in expected_keys:
                errmsg = f'Unexpected parameter {i} in JSON data'
                print(errmsg, file=sys.stderr)
                raise KeyError(errmsg)

    @staticmethod
    def check_dict_parse(self_data: dict[str, Any], json_data: dict[str, Any],
                         key: str, ok_to_use_defaults: bool,
                         unchecked_dicts: list[str]) -> None:
        """Check and parse dictionary."""
        if not isinstance(self_data, dict) and \
                not isinstance(json_data, dict):
            return
        if isinstance(self_data, dict):
            if not isinstance(json_data, dict):
                errmsg = f'Not dictionary for {key} in JSON data'
                print(errmsg, file=sys.stderr)
                raise KeyError(errmsg)
        if not isinstance(self_data, dict):
            errmsg = f'Unexpected dictionary for {key} in JSON data'
            print(errmsg, file=sys.stderr)
            raise KeyError(errmsg)
        if key in unchecked_dicts:
            return
        Config.check_key_match(list(self_data.keys()), list(json_data.keys()),
                               ok_to_use_defaults)
        for i in self_data.keys():
            if i in json_data:
                Config.check_dict_parse(self_data[i], json_data[i], i,
                                        ok_to_use_defaults, unchecked_dicts)

    def _json_parse_obj_hook(self, indict: dict[str, Any]) -> dict[str, Any]:
        """Convert str to correct type."""
        hookd = self._hook_dict
        if hookd is None:
            return indict  # pragma: no cover
        ret = deepcopy(indict)
        for key, value in ret.items():
            if key in hookd:
                parse_c = hookd[key]
                if not isinstance(value, parse_c.result_type):
                    ret[key] = parse_c.func(value, **parse_c.args)
        return ret

    def parse_json(self, from_json_text: str,
                   ok_to_use_defaults: bool = False) -> None:
        """Parse a string of JSON data and set self to that."""
        self._hook_dict = self.parse_converters()
        hook = self._json_parse_obj_hook if self._hook_dict is not None \
            else None
        data = None
        try:
            data = json.loads(from_json_text, object_hook=hook)
        except Exception as exc:
            if isinstance(exc, NotImplementedError):
                raise exc
            msg = 'Config.parse_json failed to load JSON from string/file.\n'
            msg += 'Probably incorrectly edited configuration,\n'
            msg += 'or using wrong file (not config file) as configuration.\n'
            msg += str(exc)
            print(msg, file=sys.stderr)
            if isinstance(exc, json.JSONDecodeError):
                raise ConfigBadJson(msg=msg, doc=exc.doc, pos=exc.pos) from exc
            raise ConfigBadJson(msg=msg, doc='', pos=0) from exc
        self_keys = [i for i in vars(self).keys() if not
                     callable(getattr(self, i)) and not i.startswith('_')]
        self.check_key_match(self_keys, data.keys(), ok_to_use_defaults)
        for i in self_keys:
            if i in data.keys():
                self.check_dict_parse(getattr(self, i), data[i], i,
                                      ok_to_use_defaults,
                                      self._unchecked_dicts)
                setattr(self, i, data[i])

    def as_json_string(self) -> str:
        """Return string with this configuration as JSON data."""
        data = {}
        self_keys = [i for i in vars(self).keys() if not
                     callable(getattr(self, i)) and not i.startswith('_')]
        for i in self_keys:
            data[i] = getattr(self, i)
        return json.dumps(data, sort_keys=True, indent=4, cls=ConfigEncoder)

    def read(self, from_json_filename: str,
             ok_to_use_defaults: bool = False) -> None:
        """Read configuration JSON from named file."""
        file_must_exist(filename=from_json_filename,
                        with_content_txt='with configuration JSON input')
        with open(from_json_filename, "r", encoding='UTF-8') as file:
            data = file.read()
            self.parse_json(data, ok_to_use_defaults)

    def write(self, to_json_filename: str) -> None:
        """Create named file with configuration as JSON."""
        with open(to_json_filename, "w", encoding='UTF-8') as file:
            text = self.as_json_string()
            file.write(text)

    @staticmethod
    def get_csv_dialect(name: Optional[str], delimiter: Optional[str],  # pylint: disable=too-many-arguments, line-too-long, too-many-branches # noqa: E501
                        quoting: Optional[str], quotechar: Optional[str],  # pylint: disable=too-many-arguments, line-too-long, too-many-branches # noqa: E501
                        lineterminator: Optional[str],  # pylint: disable=too-many-arguments, line-too-long, too-many-branches # noqa: E501
                        escapechar: Optional[str]  # pylint: disable=too-many-arguments, line-too-long, too-many-branches # noqa: E501
                        ) -> type[csv.Dialect]:  # pylint: disable=too-many-arguments, line-too-long, too-many-branches # noqa: E501
        """Get CSV dialect object matching arguments."""
        ret: Optional[type[csv.Dialect]] = None
        if name is None or name.lower() == 'csv.excel':
            ret = csv.excel
            ret.lineterminator = '\r\n'
        elif name.lower() == 'csv.excel_tab':
            ret = csv.excel_tab
            ret.lineterminator = '\r\n'
        elif name.lower() == 'csv.unix_dialect':
            ret = csv.unix_dialect
            ret.lineterminator = '\n'
        else:
            errmsg = f'Unknown csv dialect: {name}'
            print(errmsg, file=sys.stderr)
            raise KeyError(errmsg)
        if delimiter is not None:
            ret.delimiter = delimiter
        if quoting is None:
            ret.quoting = csv.QUOTE_MINIMAL
        elif quoting.lower() == 'csv.quote_all':
            ret.quoting = csv.QUOTE_ALL
        elif quoting.lower() == 'csv.quote_minimal':
            ret.quoting = csv.QUOTE_MINIMAL
        elif quoting.lower() == 'csv.quote_none':
            ret.quoting = csv.QUOTE_NONE
        elif quoting.lower() == 'csv.quote_nonnumeric':
            ret.quoting = csv.QUOTE_NONNUMERIC
        else:
            errmsg = f'Unknown csv quoting: {quoting}'
            print(errmsg, file=sys.stderr)
            raise KeyError(errmsg)
        if quotechar is None:
            ret.quotechar = '"'
        else:
            ret.quotechar = quotechar
        if lineterminator is not None:
            ret.lineterminator = lineterminator
        if escapechar is None:
            ret.escapechar = '\\'
        else:
            ret.escapechar = escapechar
        return ret

    @staticmethod
    def check_array_keys(name_of_cfg: str, array: list[dict[str, Any]],
                         mandatory_keys: list[str],
                         allowed_keys: Optional[list[str]] = None) -> None:
        """Check an array of dicts that all keys match allowed/mandatory.

        Check that all keys of the list of dict are allowed.
        A key is allowed if listed in mandatory_keys of allowed_keys.
        Check that all mandatory keys exist in all dicts in the list.
        @param nam_of_cfg The name of the configuration used in error message.
        @param array The array of dicts to check.
        @param mandatory_keys An array of the mandatory keys.
        @param allowed_keys An array of allowed keys excluding madatory keys.
        """
        to_allow = deepcopy(mandatory_keys)
        in_cfg = f' in config of {name_of_cfg}'
        if allowed_keys is not None:
            to_allow += deepcopy(allowed_keys)
        for i in array:
            for used_key in list(i.keys()):
                if used_key not in to_allow:
                    bad_k = f'Found non-allowed key "{used_key}"'
                    print(bad_k + in_cfg, file=sys.stderr)
                    raise KeyError(bad_k + in_cfg)
            for k in mandatory_keys:
                if k not in list(i.keys()):
                    miss = f'Missing key "{k}"'
                    print(miss + in_cfg, file=sys.stderr)
                    raise KeyError(miss + in_cfg)

    @staticmethod
    def value_of_type(input_value: Any, to_type: Any) -> Any:
        """Convert input to given type."""
        if isinstance(input_value, to_type):
            return input_value
        return to_type(input_value)

    @staticmethod
    def check_array_dicts(name_of_cfg: str,  # pylint: disable=too-many-locals
                          array: list[dict[str, Any]],
                          kind_key: str, kind_type: type,
                          dict_of_templates: Mapping[Keya, Mapping[str, type]]
                          ) -> None:
        """Check a list of dicts that all dicts matches a template.

        Check that all dicts in list of dicts matches one of the
        templates. The templates are each a dict and all keys must
        match one of the template dicts. For each key in matching
        template the value must be of the template value type.
        @param nam_of_cfg The name of the configuration used in error message.
        @param array The array of dicts to check.
        @param kind_key The key of the kind enum that determines template.
        @param kind_type The type of enum that is used as kind enum.
        @param dict_of_templates  The dict of template dicts.
        """
        Msgs = NamedTuple('Msgs', [('in_cfg', str), ('bad_arg', str),
                                   ('bad_templ', str)])
        msgs = Msgs(in_cfg=f' in config of {name_of_cfg} ',
                    bad_arg='argument not list of dicts',
                    bad_templ='Internal error: template not dict of dicts')
        if not isinstance(array, list):
            print(msgs.bad_arg + msgs.in_cfg + '(list_of_dicts)',
                  file=sys.stderr)
            raise KeyError(msgs.bad_arg + msgs.in_cfg + '(list_of_dicts)')
        if not isinstance(dict_of_templates, dict):
            print(msgs.bad_templ + msgs.in_cfg + '(dict_of_templates)',
                  file=sys.stderr)
            raise KeyError(msgs.bad_templ + msgs.in_cfg +
                           '(dict_of_templates)')
        for key1, template in dict_of_templates.items():
            if not isinstance(template, dict):
                msg = f' in template for {key1.name}'
                print(msgs.bad_templ + msgs.in_cfg + msg, file=sys.stderr)
                raise KeyError(msgs.bad_templ + msgs.in_cfg + msg)
        for i, row in enumerate(array):
            litem = f'(list index {i})'
            if not isinstance(row, dict):
                print(msgs.bad_arg + msgs.in_cfg + litem, file=sys.stderr)
                raise KeyError(msgs.bad_arg + msgs.in_cfg + litem)
            if kind_key not in row:
                msg = f'Key {kind_key} not in dict'
                print(msg + msgs.in_cfg + litem, file=sys.stderr)
                raise KeyError(msg + msgs.in_cfg + litem)
            kind = Config.value_of_type(row[kind_key], kind_type)
            for key2, valtype in dict_of_templates[kind].items():
                if key2 not in row:
                    msg = f'Key {key2} not in dict'
                    print(msg + msgs.in_cfg + litem, file=sys.stderr)
                    raise KeyError(msg + msgs.in_cfg + litem)
                if not isinstance(row[key2], valtype):
                    msg = f'Value for key {key2} = {row[key2]} '
                    msg += f'is not {valtype.__name__} '
                    msg += f'it is {type(row[key2]).__name__} '
                    print(msg + msgs.in_cfg + litem, file=sys.stderr)
                    raise KeyError(msg + msgs.in_cfg + litem)

    @staticmethod
    def get_converter_dict(enum_type: Type[Enum]) -> ParseConverter:
        """Get dict for converting to given enum_type."""
        return ParseConverter(result_type=enum_type,
                              func=string_to_enum_best_match,
                              args={'num_type': enum_type})
