#! /usr/local/bin/python3
"""Test the Config factory."""

# Copyright (c) 2024 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code

from tempfile import NamedTemporaryFile as ntf
from json import JSONDecodeError
import pytest
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.config_factory import \
    config_factory_from_enum, config_factory_from_json, \
    _config_factory_get_text, _config_factory_exit
from excel_list_transform.config_xls_list_refmt_name import \
    ConfigXlsListRefmtName
from excel_list_transform.config_xls_list_refmt_num import \
    ConfigXlsListRefmtNum


@pytest.mark.parametrize('num,typ',
                         [(ColumnRef.BY_NAME, ConfigXlsListRefmtName),
                          (ColumnRef.BY_NUMBER, ConfigXlsListRefmtNum)])
def test_config_fact_num_ok(capsys, num, typ):
    """Test OK cases of config_factory_from_enum."""
    ret = config_factory_from_enum(numerator=num)
    out, err = capsys.readouterr()
    assert isinstance(ret, typ)
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('num', [6, 'donald'])
def test_config_fact_num_nok(capsys, num):
    """Test not OK cases of config_factory_from_enum."""
    with pytest.raises(KeyError):
        _ = config_factory_from_enum(numerator=num)
    out, err = capsys.readouterr()
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('txt, fname, exc, excmsg, msg',
                         [('a', 'b', RuntimeError,
                           'Either JSON text or JSON file needed. ' +
                           'Both cannot be given.', None),
                          (None, None, RuntimeError,
                           'Either JSON text or JSON file needed. ' +
                           'Both cannot be None.', None),
                          (None, '/dev/a/b/c', SystemExit, None,
                           'File /dev/a/b/c with configuration JSON ' +
                           'input does not exist')])
def test_cfg_fact_get_text_nok(capsys,  # pylint: disable=too-many-arguments
                               txt, fname, exc, excmsg, msg):
    """Test not OK cases _config_factory_get_text."""
    with pytest.raises(exc) as cexc:
        _ = _config_factory_get_text(from_json_text=txt,
                                     from_json_filename=fname)
    out, err = capsys.readouterr()
    if msg is not None:
        assert msg in err
    if excmsg is not None:
        assert excmsg in str(cexc)
    assert '' == out


@pytest.mark.parametrize('txt', ['abc', 'def'])
def test_cfg_fact_get_text_ok_txt(capsys, txt):
    """Test OK case with text of _config_factory_get_text."""
    ret = _config_factory_get_text(from_json_text=txt,
                                   from_json_filename=None)
    out, err = capsys.readouterr()
    assert ret == txt
    assert '' == out
    assert '' == err


def test_cfg_fact_get_text_bad_enc(capsys):
    """Test bad UTF-8 encoding in file for _config_factory_get_text."""
    with ntf(mode='wb', delete_on_close=False) as tmpf:
        byt = [255, 1, 255, 0]
        tmpf.write(bytes(byt))
        tmpf.close()
        with pytest.raises(UnicodeDecodeError) as exc:
            _ = _config_factory_get_text(from_json_filename=tmpf.name,
                                         from_json_text=None)
        out, err = capsys.readouterr()
        assert 'invalid' in str(exc)
        assert '' == out
        assert '' == err


@pytest.mark.parametrize('txt', ['some text', 'another'])
def test_cfg_fact_get_text_ok_enc(capsys, txt):
    """Test good UTF-8 encoding in file for _config_factory_get_text."""
    with ntf(mode='w', delete_on_close=False) as tmpf:
        print(txt, file=tmpf)
        tmpf.close()
        ret = _config_factory_get_text(from_json_filename=tmpf.name,
                                       from_json_text=None)
        out, err = capsys.readouterr()
        assert ret.strip() == txt.strip()
        assert '' == out
        assert '' == err


@pytest.mark.parametrize('msg, exc, exctxt',
                         [('some text', None, None),
                          ('another', JSONDecodeError('msg', 'doc', 10),
                           '(char 10)'),
                          ('abc', UnicodeDecodeError('1', b'2', 3, 4, '5'),
                           "codec can't decode bytes")])
def test_cfg_fact_exit(capsys, msg, exc, exctxt):
    """Test _config_factory_exit."""
    with pytest.raises(SystemExit):
        _config_factory_exit(msg, exc=(exc if exc is not None else None))
    out, err = capsys.readouterr()
    assert '\nDid you specify an incorrect configuration file?\n' in err
    assert msg in err
    if exctxt is not None:
        assert exctxt in err
    assert '' == out


@pytest.mark.parametrize('text, exctxt, msg',
                         [(b'\xff\x01\xff\x00', "can't decode byte",
                           'Invalid UTF-8 in configuration data.'),
                          ('donald duck', 'Expecting value',
                           'Configuration JSON cannot be decoded'),
                          ('{"a": "b"}', None, 'No key "column_ref"'),
                          ('[{"a": "b"}, {"c": "d"}]', None,
                           'JSON data is not valid configuration. ' +
                           'Top level not dict')])
def test_cfg_fact_fr_json_nok(capsys, text, exctxt, msg):
    """Test not OK cases of config_factory_from_json."""
    with pytest.raises(SystemExit):
        _ = config_factory_from_json(from_json_text=text)
    out, err = capsys.readouterr()
    assert msg in err
    if exctxt is not None:
        assert exctxt in err
    assert '' == out


@pytest.mark.parametrize('kind', [ConfigXlsListRefmtName,
                                  ConfigXlsListRefmtNum])
def test_cfg_fact_fr_json_ok_txt(capsys, kind):
    """Test OK cases txt of config_factory_from_json."""
    orig = kind()
    txt = orig.as_json_string()
    cfg = config_factory_from_json(from_json_text=txt,
                                   from_json_filename=None)
    out, err = capsys.readouterr()
    assert isinstance(cfg, kind)
    assert '' == out
    assert '' == err


@pytest.mark.parametrize('kind', [ConfigXlsListRefmtName,
                                  ConfigXlsListRefmtNum])
def test_cfg_fact_fr_json_ok_file(capsys, kind):
    """Test OK cases txt of config_factory_from_json."""
    orig = kind()
    with ntf(delete_on_close=False) as file:
        fname = file.name
        file.close()
        orig.write(to_json_filename=fname)
        cfg = config_factory_from_json(from_json_text=None,
                                       from_json_filename=fname)
        out, err = capsys.readouterr()
        assert isinstance(cfg, kind)
        assert '' == out
        assert '' == err
