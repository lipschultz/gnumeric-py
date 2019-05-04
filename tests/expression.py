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
import unittest

from gnumeric import Workbook


class TestSharedExpression(unittest.TestCase):
    def setUp(self):
        self.loaded_wb = Workbook.load_workbook('samples/test.gnumeric')
        self.loaded_ws = self.loaded_wb.get_sheet_by_name('SharedExpressions')
        self.wb = Workbook()
        self.ws = self.wb.create_sheet('Title')

