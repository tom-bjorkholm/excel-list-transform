#! /usr/local/bin/python3
"""Setup file specifying build of apo_tools.whl with APO tools library."""

from setuptools import setup

setup(
  name='excel-list-transform',
  version='0.6.2',
  description='Transform a list in excel or CSV.',
  author='Tom Björkholm',
  author_email='klausuler_linnet0q@icloud.com',
  python_requires='>=3.12.6',
  packages=['excel_list_transform'],
  package_dir={'excel_list_transform': 'src/excel_list_transform'},
  package_data={'excel_list_transform': ['src/py.typed']},
  install_requires=[
    'openpyxl >= 3.1.5',
    'types-openpyxl >= 3.1.5.20240918',
    'pylightxl >= 1.61',
    'XlsxWriter >= 3.2.0',
    'pip >= 24.2',
    'setuptools >= 75.1.0',
    'build >= 1.2.2',
    'wheel>=0.44.0'
  ]
)
