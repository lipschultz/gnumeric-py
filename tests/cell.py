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

from gnumeric import cell
from gnumeric.workbook import Workbook


class CellTests(unittest.TestCase):
    def setUp(self):
        self.loaded_wb = Workbook.load_workbook('samples/test.gnumeric')
        self.wb = Workbook()
        self.ws = self.wb.create_sheet('Title')

    def test_cell_types(self):
        ws = self.loaded_wb.get_sheet_by_name('CellTypes')
        for row in range(ws.max_row + 1):
            expected_type = ws.cell(row, 0).text
            if expected_type == 'Integer':
                expected_type = cell.VALUE_TYPE_INTEGER
            elif expected_type == 'Float':
                expected_type = cell.VALUE_TYPE_FLOAT
            elif expected_type == 'Equation':
                expected_type = cell.VALUE_TYPE_EXPR
            elif expected_type == 'Boolean':
                expected_type = cell.VALUE_TYPE_BOOLEAN
            elif expected_type == 'String':
                expected_type = cell.VALUE_TYPE_STRING
            elif expected_type == 'Error':
                expected_type = cell.VALUE_TYPE_ERROR
            elif expected_type == 'Empty':
                expected_type = cell.VALUE_TYPE_EMPTY

            test_cell = ws.cell(row, 1)
            self.assertEqual(test_cell.value_type, expected_type,
                             str(test_cell.text) + ' (row=' + str(row) + ') has type ' + str(test_cell.value_type)
                             + ', but expected ' + str(expected_type))

    def test_get_cell_values(self):
        ws = self.loaded_wb.get_sheet_by_name('CellTypes')
        for row in range(ws.max_row + 1):
            test_cell = ws.cell(row, 1)

            expected_value = test_cell.text
            if test_cell.value_type == cell.VALUE_TYPE_INTEGER:
                expected_value = int(expected_value)
            elif test_cell.value_type == cell.VALUE_TYPE_FLOAT:
                expected_value = float(expected_value)
            elif test_cell.value_type == cell.VALUE_TYPE_BOOLEAN:
                expected_value = bool(expected_value)
            elif test_cell.value_type == cell.VALUE_TYPE_EMPTY:
                expected_value = None

            if test_cell.value_type == cell.VALUE_TYPE_EXPR:
                self.assertEqual(test_cell.get_value().value, expected_value)
            else:
                self.assertEqual(test_cell.get_value(), expected_value,
                                 str(test_cell.text) + ' (row=' + str(row) + ') has type ' + str(test_cell.value_type)
                                 + ', but expected ' + str(expected_value))

    def test_get_string_cell_value_returns_parsed_value(self):
        ws = self.loaded_wb.get_sheet_by_name('Strings')
        expected_values = ('TBD', 'Blåbærgrød', 'Greek Α α', 'Greek Β β', 'Greek Γ γ', 'Greek Δ δ', 'Greek Ε ε', 'Greek Ζ ζ', 'Greek Η η', 'Greek Θ θ', 'Greek Ι ι',
                           'Greek Κ κ', 'Greek Λ λ', 'Greek Μ μ', 'Greek Ν ν', 'Greek Ξ ξ', 'Greek Ο ο', 'Greek Π π', 'Greek Ρ ρ', 'Greek Σ σς', 'Greek Τ τ', 'Greek Υ υ',
                           'Greek Φ φ', 'Greek Χ χ', 'Greek Ψ ψ', 'Greek Ω ω', 'TBD', '10', '-10', '1.23', '1e1', '1e+01', '1E+01', '1D+01', '=2+2', 'TRUE', 'FALSE', '#N/A',
                           '#DIV/0!', '#VALUE!', '#NAME?', '#NUM!', ' abc', 'abc ', 'abc"def', 'abc""def', "abc'def", "abc''def", "'abc", 'abc&def', 'abc&amp;def', 'abc<def',
                           'abc>def', 'abc	def', 'hi<!--there')
        for row in range(ws.max_row + 1):
            expected = expected_values[row]

            test_cell = ws.cell(row, 0)
            actual = test_cell.value

            self.assertEqual(cell.VALUE_TYPE_STRING, test_cell.value_type, f'Failed on row={row+1}, cell={actual}, expected={expected}')
            self.assertEqual(expected, actual, f'Failed on row={row+1}, cell={actual}, expected={expected}')

    def test_get_shared_expression_value_from_originating_cell(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')

        expected_value = '=sum(A2:A10)'
        expected_id = '1'
        expected_originating_cell = ws.cell(1, 1)
        expected_cells = [expected_originating_cell, ws.cell(4, 1)]

        c1 = ws.cell(1, 1)
        c1_val = c1.value
        self.assertEqual(c1_val.value, expected_value)
        self.assertEqual(c1_val.id, expected_id)
        self.assertEqual(c1_val.get_originating_cell(), expected_originating_cell)
        self.assertEqual(c1_val.get_all_cells(), expected_cells)

    def test_get_shared_expression_value_from_cell_using_expression(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')

        expected_value = '=sum(A2:A10)'
        expected_id = '1'
        expected_originating_cell = ws.cell(1, 1)
        expected_cells = [expected_originating_cell, ws.cell(4, 1)]

        c1 = ws.cell(4, 1)
        c1_val = c1.value
        self.assertEqual(c1_val.value, expected_value)
        self.assertEqual(c1_val.id, expected_id)
        self.assertEqual(c1_val.get_originating_cell(), expected_originating_cell)
        self.assertEqual(c1_val.get_all_cells(), expected_cells)

    def test_get_non_shared_expression_value(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')

        expected_value = '=product(BoundingRegion!D7:J13)'
        expected_id = None
        expected_originating_cell = ws.cell(3, 1)
        expected_cells = [expected_originating_cell]

        c1 = ws.cell(3, 1)
        c1_val = c1.value
        self.assertEqual(c1_val.value, expected_value)
        self.assertEqual(c1_val.id, expected_id)
        self.assertEqual(c1_val.get_originating_cell(), expected_originating_cell)
        self.assertEqual(c1_val.get_all_cells(), expected_cells)

    def test_set_cell_value_infer_bool(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(True)
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_BOOLEAN)

    def test_set_cell_value_infer_int(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(10)
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_INTEGER)

    def test_set_cell_value_infer_float(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(10.0)
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_FLOAT)

    def test_new_cell_is_empty(self):
        test_cell = self.ws.cell(0, 0)
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_EMPTY)

    def test_set_cell_value_infer_empty_from_empty_string(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('')
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_EMPTY)

    def test_set_cell_value_infer_empty_from_None(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(None)
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_EMPTY)

    def test_set_cell_value_infer_expression(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('=max()')
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_EXPR)

    def test_set_cell_value_infer_string(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('asdf')
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_STRING)

    def test_cell_type_changes_when_value_changes(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('asdf')
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_STRING)
        test_cell.set_value(17)
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_INTEGER)

    def test_save_int_into_cell_value_using_keep_when_cell_value_is_string(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('asdf')
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_STRING)
        test_cell.set_value(17, value_type='keep')
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_STRING)

    def test_explicitly_setting_value_type_uses_that_value_type(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(17, value_type=cell.VALUE_TYPE_STRING)
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_STRING)

    def test_setting_value_infers_type(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.value = 17
        self.assertEqual(test_cell.value_type, cell.VALUE_TYPE_INTEGER)

    def test_cells_equal(self):
        test_cell1 = self.ws.cell(0, 0)
        test_cell2 = self.ws.cell(0, 0)
        self.assertTrue(test_cell1 == test_cell2)

    def test_cells_unequal(self):
        test_cell1 = self.ws.cell(0, 0)
        test_cell2 = self.ws.cell(1, 0)
        self.assertFalse(test_cell1 == test_cell2)

    def test_cell_coordinate(self):
        row = 6
        col = 4
        c = self.ws.cell(row, col)
        self.assertEqual(c.coordinate, (row, col))

    def test_get_text_format(self):
        ws = self.loaded_wb.get_sheet_by_name('Mine & Yours Sheet[s]!')
        self.assertEqual(ws.cell(0, 0).text_format, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)')

    def test_copying_shared_expression_to_new_cell(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')

        expected_originating_cell = ws.cell(1, 1)

        new_cell = ws.cell(5, 5)
        new_cell.value = expected_originating_cell.value

        expected_value = expected_originating_cell.text
        expected_id = '1'
        expected_cells = [expected_originating_cell, ws.cell(4, 1), new_cell]

        nc_val = new_cell.value
        self.assertEqual(nc_val.value, expected_value)
        self.assertEqual(nc_val.id, expected_id)
        self.assertEqual(nc_val.get_originating_cell(), expected_originating_cell)
        self.assertEqual(nc_val.get_all_cells(), expected_cells)

        nc_val = expected_originating_cell.value
        self.assertEqual(nc_val.value, expected_value)
        self.assertEqual(nc_val.id, expected_id)
        self.assertEqual(nc_val.get_originating_cell(), expected_originating_cell)
        self.assertEqual(nc_val.get_all_cells(), expected_cells)

    def test_sharing_a_non_shared_expression(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')

        expected_originating_cell = ws.cell(3, 1)
        expected_id = str(max(int(k) for k in ws.get_expression_map()) + 1)

        new_cell = ws.cell(5, 5)
        new_cell.value = expected_originating_cell.value

        expected_value = expected_originating_cell.text
        expected_cells = [expected_originating_cell, new_cell]

        nc_val = new_cell.value
        self.assertEqual(nc_val.value, expected_value)
        self.assertEqual(nc_val.id, expected_id)
        self.assertEqual(nc_val.get_originating_cell(), expected_originating_cell)
        self.assertEqual(nc_val.get_all_cells(), expected_cells)

        nc_val = expected_originating_cell.value
        self.assertEqual(nc_val.value, expected_value)
        self.assertEqual(nc_val.id, expected_id)
        self.assertEqual(nc_val.get_originating_cell(), expected_originating_cell)
        self.assertEqual(nc_val.get_all_cells(), expected_cells)

    def test_deleting_originating_cell_of_shared_expression(self):
        pass
