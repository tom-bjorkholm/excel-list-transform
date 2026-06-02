#! /usr/local/bin/python3
"""Test recode command and recode function."""

from tempfile import TemporaryDirectory
import pytest
from pytest import CaptureFixture
from excel_list_transform.recode import recode_cmd


@pytest.mark.parametrize('inenc', ['utf-8', 'iso8859-1', 'cp1252'])
@pytest.mark.parametrize('outenc', ['utf-8', 'iso8859-1', 'cp1252'])
@pytest.mark.parametrize('infile', ['a.txt', 'b.csv'])
@pytest.mark.parametrize('outfile', ['c.txt', 'd.csv'])
@pytest.mark.parametrize('text',
                         ['Hello \nand goodbye!',
                          'Some Swedish letters are ÅÄÖ and åäö',
                          'In Denmark they us æø and ÆØ'])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_recode_ok1(capsys: CaptureFixture[str], infile: str, inenc: str,
                    outfile: str, outenc: str, text: str) -> None:
    """Test OK cases 1 of recode command."""
    with TemporaryDirectory() as dirname:
        inf = dirname + '/' + infile
        outf = dirname + '/' + outfile
        with open(file=inf, encoding=inenc, mode='w') as file:
            file.write(text)
        args = ['-i', inf, '-o', outf, '-f', inenc, '-t', outenc]
        recode_cmd(args=args)
        out, err = capsys.readouterr()
        with open(file=outf, encoding=outenc, mode='r') as file:
            content = file.read()
            assert content == text
        assert '' == out
        assert f'Wrote file {outf}\n' == err


@pytest.mark.parametrize('inenc', ['utf-8', 'iso8859-1', 'cp1252'])
@pytest.mark.parametrize('outenc', ['utf-8', 'iso8859-1', 'cp1252'])
@pytest.mark.parametrize('infile', ['a.txt', 'b.csv'])
@pytest.mark.parametrize('outfile', ['c.txt', 'd.csv'])
@pytest.mark.parametrize('text',
                         ['Hello \nand goodbye!',
                          'Some Swedish letters are ÅÄÖ and åäö',
                          'In Denmark they us æø and ÆØ'])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_recode_ok2(capsys: CaptureFixture[str], infile: str, inenc: str,
                    outfile: str, outenc: str, text: str) -> None:
    """Test OK cases 2 of recode command."""
    with TemporaryDirectory() as dirname:
        inf = dirname + '/' + infile
        outf = dirname + '/' + outfile
        with open(file=inf, encoding=inenc, mode='w') as file:
            file.write(text)
        args = ['--input', inf, '--output', outf,
                '--from', inenc, '--to', outenc]
        recode_cmd(args=args)
        out, err = capsys.readouterr()
        with open(file=outf, encoding=outenc, mode='r') as file:
            content = file.read()
            assert content == text
        assert '' == out
        assert f'Wrote file {outf}\n' == err


@pytest.mark.parametrize('inenc', ['utf-8', 'iso8859-1', 'cp1252'])
@pytest.mark.parametrize('outenc', ['utf-8', 'iso8859-1', 'cp1252'])
@pytest.mark.parametrize('infile', ['a.txt', 'b.csv'])
@pytest.mark.parametrize('outfile', ['c.txt', 'd.csv'])
@pytest.mark.parametrize('text',
                         ['Hello \nand goodbye!',
                          'Some Swedish letters are ÅÄÖ and åäö',
                          'In Denmark they us æø and ÆØ'])
# pylint: disable-next=too-many-arguments,too-many-positional-arguments
def test_recode_ok3(capsys: CaptureFixture[str], infile: str, inenc: str,
                    outfile: str, outenc: str, text: str) -> None:
    """Test OK cases 3 of recode command."""
    with TemporaryDirectory() as dirname:
        inf = dirname + '/' + infile
        outf = dirname + '/' + outfile
        with open(file=inf, encoding=inenc, mode='w') as file:
            file.write(text)
        args = ['python3', '-m', 'recode.py',
                '--input', inf, '--output', outf,
                '--from', inenc, '--to', outenc]
        recode_cmd(args=args)
        out, err = capsys.readouterr()
        with open(file=outf, encoding=outenc, mode='r') as file:
            content = file.read()
            assert content == text
        assert '' == out
        assert f'Wrote file {outf}\n' == err


def test_recode_nok1(capsys: CaptureFixture[str]) -> None:
    """Test recode error case 1."""
    with TemporaryDirectory() as dirname:
        inf = dirname + '/a.txt'
        outf = dirname + '/a.txt'
        with open(file=inf, encoding='utf-8', mode='w') as file:
            file.write('Some text')
        args = ['--input', inf, '--output', outf,
                '--from', 'utf-8', '--to', 'iso8859-1']
        with pytest.raises(SystemExit):
            recode_cmd(args=args)
        out, err = capsys.readouterr()
        assert f'Error: Output file {outf} already exists.' in err
        assert '' == out


def test_recode_nok2(capsys: CaptureFixture[str]) -> None:
    """Test recode error case 2."""
    with TemporaryDirectory() as dirname:
        inf = dirname + '/a.txt'
        outf = dirname + '/b.txt'
        args = ['--input', inf, '--output', outf,
                '--from', 'utf-8', '--to', 'iso8859-1']
        with pytest.raises(SystemExit):
            recode_cmd(args=args)
        out, err = capsys.readouterr()
        assert f'Error: Input file {inf} not found.' in err
        assert '' == out


def test_recode_nok3(capsys: CaptureFixture[str]) -> None:
    """Test recode error case 3."""
    with TemporaryDirectory() as dirname:
        inf = dirname + '/a.txt'
        outf = dirname + '/b.txt'
        with open(file=inf, encoding='utf-8', mode='w') as file:
            file.write('Some text')
        args = ['--input', inf, '--output', outf,
                '--from', 'utf-9', '--to', 'iso8859-1']
        with pytest.raises(SystemExit):
            recode_cmd(args=args)
        out, err = capsys.readouterr()
        assert 'Error: Unrecognized input encoding utf-9' in err
        assert '' == out


def test_recode_nok4(capsys: CaptureFixture[str]) -> None:
    """Test recode error case 4."""
    with TemporaryDirectory() as dirname:
        inf = dirname + '/a.txt'
        outf = dirname + '/b.txt'
        with open(file=inf, encoding='utf-8', mode='w') as file:
            file.write('Some text')
        args = ['--input', inf, '--output', outf,
                '--from', 'utf-8', '--to', 'iso18859-172']
        with pytest.raises(SystemExit):
            recode_cmd(args=args)
        out, err = capsys.readouterr()
        assert 'Error: Unrecognized output encoding iso18859-172' in err
        assert '' == out


def test_recode_nok5(capsys: CaptureFixture[str]) -> None:
    """Test recode error case 4."""
    with TemporaryDirectory() as dirname:
        inf = dirname + '/a.txt'
        outf = dirname + '/b.txt'
        with open(file=inf, encoding='utf-8', mode='w') as file:
            file.write('Some text ÅÄÖåäö')
        args = ['--input', inf, '--output', outf,
                '--from', 'utf-8', '--to', 'ascii']
        with pytest.raises(SystemExit):
            recode_cmd(args=args)
        out, err = capsys.readouterr()
        assert 'Error: file content cannot be decoded/encoded' in err
        assert 'from utf-8 to ascii' in err
        assert '' == out
