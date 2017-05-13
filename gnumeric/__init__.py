"""
Gnumeric-py: Reading and writing gnumeric files with python
Copyright (C) 2017 Michael Lipschultz

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <http://www.gnu.org/licenses/>.
"""

import json
import os

try:
    here = os.path.abspath(os.path.dirname(__file__))
    src_file = os.path.join(here, ".constants.json")
    with open(src_file) as src:
        constants = json.load(src)
        __author__ = constants['__author__']
        __author_email__ = constants["__author_email__"]
        __license__ = constants["__license__"]
        __maintainer_email__ = constants["__maintainer_email__"]
        __url__ = constants["__url__"]
        __version__ = constants["__version__"]
except IOError:
    # packaged
    pass

from gnumeric.workbook import Workbook


def load_workbook(filepath):
    return Workbook.load_workbook(filepath)
