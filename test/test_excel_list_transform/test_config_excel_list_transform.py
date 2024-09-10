#! /usr/local/bin/python3
"""Test the ConfigExcelListTransform class."""

# Copyright (c) 2024 Tom Björkholm
# MIT License

# pylint: disable=duplicate-code


import pytest
from excel_list_transform.config_enums import ColumnRef
from excel_list_transform.config_excel_list_transform import \
    FileType, ConfigExcelListTransform, ColInfo


@pytest.mark.smoke
def test_config_exc_list_refrm_def(capsys):
    """Test default values of ConfigExcelListTransform."""
    col_to_use = ['street', 'street number', 'name', 'last name',
                  'Phone', 'Phone', 'Phone', 'Phone', 'Phone',
                  'Last Name']
    colinfo = ColInfo(split_last='right_name', insert_last=None,
                      col_to_use=col_to_use, tinfo='a',
                      s1=[], s6=[])
    cfg = ConfigExcelListTransform(ColumnRef.BY_NAME,
                                   colinfo=colinfo, tinfo='a')
    assert cfg.in_type == FileType.EXCEL
    assert cfg.out_type == FileType.EXCEL
    assert cfg.column_ref == ColumnRef.BY_NAME
    assert cfg.max_column_read == 20
    str_cfg = cfg.as_json_string()
    assert len(str_cfg) > 1
    assert 'in_type' in str_cfg
    zcfg = ConfigExcelListTransform(ColumnRef.BY_NAME,
                                    colinfo=colinfo, tinfo='a')
    assert cfg.__dict__ == zcfg.__dict__
    ycfg = ConfigExcelListTransform(ColumnRef.BY_NAME,
                                    colinfo=colinfo, tinfo='a',
                                    from_json_text=str_cfg)
    assert ycfg.__dict__ == cfg.__dict__
    assert cfg.out_csv_dialect['lineterminator'] is None
    assert ycfg.out_csv_dialect['lineterminator'] is None
    assert cfg.out_csv_dialect.keys() == ycfg.out_csv_dialect.keys()
    assert cfg.out_csv_dialect.keys() == zcfg.out_csv_dialect.keys()
    assert cfg.out_csv_dialect == ycfg.out_csv_dialect
    assert cfg.out_csv_dialect == zcfg.out_csv_dialect
    assert cfg.get_out_csv_dialect().lineterminator == \
        ycfg.get_out_csv_dialect().lineterminator
    assert ycfg.get_out_csv_dialect().lineterminator == '\r\n'
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('t',
                         ['{"out_type_" : "CSV"}',
                          '{"outfilen" : "out.dat"}'])
def test_config_exc_list_reform_read_incomplete4(capsys, t):
    """Test ConfigExcelListTransform read incomplete 4."""
    col_to_use = [15, 16, 1, 2, 5, 5, 5, 5, 5, 6]
    colinfo = ColInfo(split_last='store_single',
                      insert_last='name',
                      col_to_use=col_to_use, tinfo=2,
                      s1=[], s6=[])
    cfg = ConfigExcelListTransform(col_ref=ColumnRef.BY_NUMBER,
                                   colinfo=colinfo, tinfo=2)
    with pytest.raises(KeyError) as exc_info:
        cfg.parse_json(t, ok_to_use_defaults=True)
    assert exc_info.type is KeyError
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Unexpected' in err


@pytest.mark.parametrize('data,par',
                         [([1, 2, 3], 'test'),
                          ([1, 4, 3, 2], 'test')])
def test_check_no_dupl_ok(capsys, data, par):
    """Test check:_no_cuplicates for OK cases."""
    ConfigExcelListTransform._duplicates_not_allowed(data, par)  # pylint: disable=protected-access # noqa: E501
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'column': 2},
                            {'a': 1, 'column': 3},
                            {'a': 3, 'column': 4}], 'test')])
def test_check_no_dupl_key_ok(capsys, data, par):
    """Test check:_no_cuplicates for OK cases."""
    ConfigExcelListTransform._check_no_duplicate_single(data, par, 2)  # pylint: disable=protected-access # noqa: E501
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'columns': [1, 2]},
                            {'a': 1, 'columns': [3, 7]},
                            {'a': 3, 'columns': [4, 8]}], 'test')])
def test_check_no_dupl_mul_ok(capsys, data, par):
    """Test check:_no_cuplicates for OK cases."""
    ConfigExcelListTransform._check_no_duplicate_multi(data, par, 2)  # pylint: disable=protected-access # noqa: E501
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('data,par',
                         [([1, 2, 3, 2], 'test'),
                          ([1, 2, 3, 1], 'test')])
def test_check_no_dupl_nok(capsys, data, par):
    """Test check:_no_cuplicates for OK cases."""
    with pytest.raises(KeyError) as exc:
        ConfigExcelListTransform._duplicates_not_allowed(data, par)  # pylint: disable=protected-access # noqa: E501
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Duplicates not allowed in' in str(exc)
    assert 'Duplicates not allowed in' in err


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'column': 2},
                            {'a': 1, 'column': 3},
                            {'a': 3, 'column': 2}], 'test')])
def test_check_no_dupl_keyd_nok(capsys, data, par):
    """Test check:_no_cuplicates for OK cases."""
    with pytest.raises(KeyError) as exc:
        ConfigExcelListTransform._check_no_duplicate_single(data, par, 2)  # pylint: disable=protected-access # noqa: E501
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Duplicates not allowed in' in str(exc)
    assert 'Duplicates not allowed in' in err


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'columns': [1, 2]},
                            {'a': 1, 'columns': [3, 4]},
                            {'a': 3, 'columns': [2, 5]}], 'test')])
def test_check_no_dupl_num_nok(capsys, data, par):
    """Test check:_no_cuplicates for OK cases."""
    with pytest.raises(KeyError) as exc:
        ConfigExcelListTransform._check_no_duplicate_multi(data, par, 2)  # pylint: disable=protected-access # noqa: E501
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Duplicates not allowed in' in str(exc)
    assert 'Duplicates not allowed in' in err


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'columns': [2, 4]},
                            {'a': 1, 'columns': [5, 6]},
                            {'a': 3, 'columns': [7, 8]}], 'test')])
def test_check_increasing_ok(capsys, data, par):
    """Test check:_no_cuplicates for OK cases."""
    ConfigExcelListTransform._check_increasing_multi(data, par, 2)  # pylint: disable=protected-access # noqa: E501
    out, err = capsys.readouterr()
    assert out == ''
    assert err == ''


@pytest.mark.parametrize('data,par',
                         [([{'a': 1, 'columns': [2, 4]},
                            {'a': 1, 'columns': [3, 6]},
                            {'a': 3, 'columns': [7, 8]}], 'test')])
def test_check_increasing_nok(capsys, data, par):
    """Test check:_no_cuplicates for OK cases."""
    with pytest.raises(KeyError) as exc:
        ConfigExcelListTransform._check_increasing_multi(data, par, 2)  # pylint: disable=protected-access # noqa: E501
    out, err = capsys.readouterr()
    assert out == ''
    assert 'Increasing order needed in' in err
    assert 'Increasing order needed in' in str(exc)
