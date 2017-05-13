#!/usr/bin/env python

import json
import os
from distutils.core import setup

here = os.path.abspath(os.path.dirname(__file__))
try:
    with open(os.path.join(here, 'README.md')) as fin:
        README = fin.read()
except IOError:
    README = ''

src_file = os.path.join(here, "gnumeric", ".constants.json")
with open(src_file) as src:
    constants = json.load(src)
    __author__ = constants['__author__']
    __author_email__ = constants["__author_email__"]
    __license__ = constants["__license__"]
    __maintainer_email__ = constants["__maintainer_email__"]
    __url__ = constants["__url__"]
    __version__ = constants["__version__"]

setup(name='Gnumeric-py',
      version=__version__,
      description='A python library for reading and writing Gnumeric files.',
      long_description=README,
      author=__author__,
      author_email=__author_email__,
      url=__url__,
      license=__license__,
      package_dir={'gnumeric': 'gnumeric'},
      packages=['gnumeric'],
      requires=['python (>=3.3.0)', 'lxml', 'dateutil'],
      data_files=[('', ['LICENSE', 'README.md', 'gnumeric/.constants.json'])],
      classifiers=['Development Status :: 3 - Alpha',
                   'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
                   'Operating System :: OS Independent',
                   'Programming Language :: Python :: 3',
                   'Topic :: Office/Business :: Financial :: Spreadsheet']
      )
