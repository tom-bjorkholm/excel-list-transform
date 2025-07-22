#! /usr/local/bin/python3
"""Setup file specifying build of .whl."""

from setuptools import setup

setup(
  name='excel-list-transform',
  version='0.8.2',
  description='Transform a list in excel or CSV.',
  author='Tom Björkholm',
  author_email='klausuler_linnet0q@icloud.com',
  python_requires='<3.13,>=3.12',
  packages=['excel_list_transform'],
  package_dir={'excel_list_transform': 'src/excel_list_transform'},
  package_data={'excel_list_transform': ['src/py.typed']},
  install_requires=[
    'openpyxl >= 3.1.5',
    'types-openpyxl >= 3.1.5.20250602',
    'pylightxl >= 1.61',
    'XlsxWriter >= 3.2.5',
    'argcomplete >= 3.6.2',
    'pypi-simple >= 1.6.1',
    'requests >= 2.32.4',
    'types-requests >= 2.32.4.20250611',
    'packaging >= 25.0',
    'pip >= 25.1.1',
    'setuptools >= 80.9.0',
    'build >= 1.2.2.post1',
    'wheel>=0.45.1'
  ]
)
