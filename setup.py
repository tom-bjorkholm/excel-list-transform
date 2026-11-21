#! /usr/local/bin/python3
"""Setup file specifying build of .whl."""

from setuptools import setup

setup(
  name='excel-list-transform',
  version='1.3.1',
  description='Transform table data with configurable row and column changes.',
  author='Tom Björkholm',
  author_email='klausuler_linnet0q@icloud.com',
  python_requires='>=3.13',
  packages=['excel_list_transform'],
  package_dir={'excel_list_transform': 'src/excel_list_transform'},
  package_data={'excel_list_transform': ['src/py.typed']},
  install_requires=[
    'tableio-cfg-json >= 2.1',
    'tableio >= 1.3',
    'config-as-json >= 1.9',
    'versionreporter >= 0.5',
    'argcomplete >= 3.7.2',
    'edit-cfg-json >= 1.1.0',
    'edit-cfg-json-tk >= 1.1.0'
  ],
  extras_require={
    'textual': ['edit-cfg-json-textual >= 1.1.0']
  }
)
