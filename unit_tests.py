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

from src.workbook import Workbook


class WorkbookTests(unittest.TestCase):
    def test_creating_empty_workbook_has_zero_sheets(self):
        wb = Workbook()
        self.assertEqual(len(wb), 0)
        self.assertEqual(len(wb.get_sheet_names()), 0)

    def test_creating_workbook_has_version_1_12_28(self):
        wb = Workbook()
        self.assertEqual(wb.version, "1.12.28")
