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
import datetime

import pytest

from gnumeric import cell
from gnumeric.workbook import Workbook


TEST_GNUMERIC_FILE_PATH = 'samples/test.gnumeric'


@pytest.fixture()
def empty_worksheet():
    workbook = Workbook()
    worksheet = workbook.create_sheet('Title')
    yield worksheet


class TestCell:
    def test_cells_equal(self, empty_worksheet):
        test_cell1 = empty_worksheet.cell(0, 0)
        test_cell2 = empty_worksheet.cell(0, 0)

        assert test_cell1 == test_cell2

    def test_cells_unequal(self, empty_worksheet):
        test_cell1 = empty_worksheet.cell(0, 0)
        test_cell2 = empty_worksheet.cell(1, 0)

        assert test_cell1 != test_cell2

    def test_cell_coordinate(self, empty_worksheet):
        row = 6
        col = 4
        cell = empty_worksheet.cell(row, col)

        assert cell.coordinate == (row, col)

    def test_get_text_format(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)

        worksheet = workbook.get_sheet_by_name('Mine & Yours Sheet[s]!')
        assert worksheet.cell(0, 0).text_format == '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'


class TestCellType:
    def test_cell_types(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('CellTypes')

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
            assert test_cell.value_type == expected_type, f"{test_cell.text} (row={row}) has type {test_cell.value_type}, but expected {expected_type}"

    def test_set_cell_value_infer_bool(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.set_value(True)
        assert test_cell.value_type == cell.VALUE_TYPE_BOOLEAN

    def test_set_cell_value_infer_int(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.set_value(10)
        assert test_cell.value_type == cell.VALUE_TYPE_INTEGER

    def test_set_cell_value_infer_float(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.set_value(10.0)
        assert test_cell.value_type == cell.VALUE_TYPE_FLOAT

    def test_new_cell_is_empty(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        assert test_cell.value_type == cell.VALUE_TYPE_EMPTY

    def test_set_cell_value_infer_empty_from_empty_string(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.set_value('')
        assert test_cell.value_type == cell.VALUE_TYPE_EMPTY

    def test_set_cell_value_infer_empty_from_None(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.set_value(None)
        assert test_cell.value_type == cell.VALUE_TYPE_EMPTY

    def test_set_cell_value_infer_expression(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.set_value('=max(A1:A2)')
        assert test_cell.value_type == cell.VALUE_TYPE_EXPR

    def test_set_cell_value_infer_string(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.set_value('asdf')
        assert test_cell.value_type == cell.VALUE_TYPE_STRING

    def test_cell_type_changes_when_value_changes(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)

        test_cell.set_value('asdf')
        assert test_cell.value_type == cell.VALUE_TYPE_STRING

        test_cell.set_value(17)
        assert test_cell.value_type == cell.VALUE_TYPE_INTEGER

    def test_save_int_into_cell_value_using_keep_when_cell_value_is_string(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.set_value('asdf')
        assert test_cell.value_type == cell.VALUE_TYPE_STRING

        test_cell.set_value(17, value_type='keep')
        assert test_cell.value_type == cell.VALUE_TYPE_STRING

    def test_explicitly_setting_value_type_uses_that_value_type(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.set_value(17, value_type=cell.VALUE_TYPE_STRING)
        assert test_cell.value_type == cell.VALUE_TYPE_STRING

    def test_setting_value_infers_type(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.value = 17
        assert test_cell.value_type == cell.VALUE_TYPE_INTEGER

    def test_setting_true_stores_true_value(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.value = True

        assert test_cell.value_type == cell.VALUE_TYPE_BOOLEAN
        assert test_cell.value is True
        assert test_cell.text == 'TRUE'

    def test_setting_false_stores_false_value(self, empty_worksheet):
        test_cell = empty_worksheet.cell(0, 0)
        test_cell.value = False
        assert test_cell.value_type == cell.VALUE_TYPE_BOOLEAN
        assert test_cell.value is False
        assert test_cell.text == 'FALSE'


class TestCellValue:
    def test_get_cell_values(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('CellTypes')
        for row in range(ws.max_row + 1):
            test_cell = ws.cell(row, 1)

            expected_value = test_cell.text
            if test_cell.value_type == cell.VALUE_TYPE_INTEGER:
                expected_value = int(expected_value)
            elif test_cell.value_type == cell.VALUE_TYPE_FLOAT:
                expected_value = float(expected_value)
            elif test_cell.value_type == cell.VALUE_TYPE_BOOLEAN:
                expected_value = expected_value.lower() == 'true'
            elif test_cell.value_type == cell.VALUE_TYPE_EMPTY:
                expected_value = None
            # TODO: Cell values of Error cells

            if test_cell.value_type == cell.VALUE_TYPE_EXPR:
                assert test_cell.get_value().original_text == expected_value
            else:
                assert test_cell.get_value() == expected_value, f"{test_cell.text} (row={row}) has type {test_cell.value_type}, but expected {expected_value}"

    def test_get_string_cell_value_returns_parsed_value(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Strings')

        expected_values = ('TBD', 'Blåbærgrød', 'Greek Α α', 'Greek Β β', 'Greek Γ γ', 'Greek Δ δ', 'Greek Ε ε', 'Greek Ζ ζ', 'Greek Η η', 'Greek Θ θ', 'Greek Ι ι',
                           'Greek Κ κ', 'Greek Λ λ', 'Greek Μ μ', 'Greek Ν ν', 'Greek Ξ ξ', 'Greek Ο ο', 'Greek Π π', 'Greek Ρ ρ', 'Greek Σ σς', 'Greek Τ τ', 'Greek Υ υ',
                           'Greek Φ φ', 'Greek Χ χ', 'Greek Ψ ψ', 'Greek Ω ω', 'TBD', '10', '-10', '1.23', '1e1', '1e+01', '1E+01', '1D+01', '=2+2', 'TRUE', 'FALSE', '#N/A',
                           '#DIV/0!', '#VALUE!', '#NAME?', '#NUM!', ' abc', 'abc ', 'abc"def', 'abc""def', "abc'def", "abc''def", "'abc", 'abc&def', 'abc&amp;def', 'abc<def',
                           'abc>def', 'abc	def', 'hi<!--there')
        for row in range(ws.max_row + 1):
            expected = expected_values[row]

            test_cell = ws.cell(row, 0)
            actual = test_cell.value

            assert cell.VALUE_TYPE_STRING == test_cell.value_type, f'Failed on row={row+1}, cell={actual}, expected={expected}'
            assert expected == actual, f'Failed on row={row+1}, cell={actual}, expected={expected}'

    def test_getting_result(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('CellTypes')

        for row in range(ws.max_row + 1):
            test_cell = ws.cell(row, 1)

            expected_value = test_cell.text
            if test_cell.value_type == cell.VALUE_TYPE_INTEGER:
                expected_value = int(expected_value)
            elif test_cell.value_type == cell.VALUE_TYPE_FLOAT:
                expected_value = float(expected_value)
            elif test_cell.value_type == cell.VALUE_TYPE_BOOLEAN:
                expected_value = expected_value.lower() == 'true'
            elif test_cell.value_type == cell.VALUE_TYPE_EMPTY:
                expected_value = 0
            elif test_cell.value_type == cell.VALUE_TYPE_EXPR:
                expected_value = ws.cell(row, 2).value

            assert test_cell.result == expected_value, f"{test_cell.text} (row={row}) has type {test_cell.value_type}, but expected {expected_value}"

    def test_getting_dates(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Dates')

        for row in range(ws.max_row + 1):
            expected_cell = ws.cell(row, 1)
            expected_datetime = datetime.datetime.strptime(expected_cell.value, '%Y-%m-%dT%H:%M:%S')

            test_cell = ws.cell(row, 0)
            assert test_cell.is_datetime() is True
            assert expected_datetime == test_cell.result


class TestCellSharedExpression:
    def test_get_shared_expression_value_from_originating_cell(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')

        expected_value = '=sum(A2:A10)'
        expected_id = '1'
        expected_originating_cell = ws.cell(1, 1)
        expected_cells_using_expr = [expected_originating_cell, ws.cell(4, 1)]

        c1 = ws.cell(1, 1)
        c1_val = c1.value
        expected_cells_referenced_in_expr = set(ws.get_cell_collection('A2', 'A10', include_empty=True, create_cells=True))

        assert c1_val.original_text == expected_value
        assert c1_val.id == expected_id
        assert c1_val.get_originating_cell() == expected_originating_cell
        assert c1_val.get_all_cells() == expected_cells_using_expr
        assert c1_val.reference_coordinate_offset == (0, 0)
        assert c1_val.get_referenced_cells() == expected_cells_referenced_in_expr
        assert c1_val.value == 45

    def test_get_shared_expression_value_from_cell_using_expression(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')

        expected_value = '=sum(A2:A10)'
        expected_id = '1'
        expected_originating_cell = ws.cell(1, 1)
        expected_cells = [expected_originating_cell, ws.cell(4, 1)]

        c1 = ws.cell(4, 1)
        c1_val = c1.value
        expected_cells_referenced_in_expr = set(ws.get_cell_collection('A5', 'A13', include_empty=True, create_cells=True))

        assert c1_val.original_text == expected_value
        assert c1_val.id == expected_id
        assert c1_val.get_originating_cell() == expected_originating_cell
        assert c1_val.get_all_cells() == expected_cells
        assert c1_val.reference_coordinate_offset == (3, 0)
        assert c1_val.get_referenced_cells() == expected_cells_referenced_in_expr
        assert c1_val.value == 39

    def test_get_non_shared_expression_value(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')

        expected_value = '=product(BoundingRegion!D7:J13)'
        expected_id = None
        expected_originating_cell = ws.cell(3, 1)
        expected_cells = [expected_originating_cell]

        c1 = ws.cell(3, 1)
        c1_val = c1.value
        expected_cells_referenced_in_expr = set(workbook.get_sheet_by_name('BoundingRegion').get_cell_collection('D7', 'J13', include_empty=True, create_cells=True))

        assert c1_val.original_text == expected_value
        assert c1_val.id == expected_id
        assert c1_val.get_originating_cell() == expected_originating_cell
        assert c1_val.get_all_cells() == expected_cells
        assert c1_val.reference_coordinate_offset == (0, 0)
        assert c1_val.get_referenced_cells() == expected_cells_referenced_in_expr
        assert c1_val.value == 5040

    def test_copying_shared_expression_to_new_cell(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')

        expected_originating_cell = ws.cell(1, 1)

        new_cell = ws.cell(5, 5)
        new_cell.value = expected_originating_cell.value

        expected_value = expected_originating_cell.text
        expected_id = '1'
        expected_cells = [expected_originating_cell, ws.cell(4, 1), new_cell]

        nc_val = new_cell.value
        assert nc_val.original_text == expected_value
        assert nc_val.id == expected_id
        assert nc_val.get_originating_cell() == expected_originating_cell
        assert nc_val.get_all_cells() == expected_cells
        assert nc_val.reference_coordinate_offset == (4, 4)

        nc_val = expected_originating_cell.value
        assert nc_val.original_text == expected_value
        assert nc_val.id == expected_id
        assert nc_val.get_originating_cell() == expected_originating_cell
        assert nc_val.get_all_cells() == expected_cells
        assert nc_val.reference_coordinate_offset == (0, 0)

    def test_sharing_a_non_shared_expression(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')

        expected_originating_cell = ws.cell(3, 1)
        expected_id = str(max(int(k) for k in ws.get_expression_map()) + 1)

        new_cell = ws.cell(5, 5)
        new_cell.value = expected_originating_cell.value

        expected_value = expected_originating_cell.text
        expected_cells = [expected_originating_cell, new_cell]

        nc_val = new_cell.value
        assert nc_val.original_text == expected_value
        assert nc_val.id == expected_id
        assert nc_val.get_originating_cell() == expected_originating_cell
        assert nc_val.get_all_cells() == expected_cells
        assert nc_val.reference_coordinate_offset == (2, 4)

        nc_val = expected_originating_cell.value
        assert nc_val.original_text == expected_value
        assert nc_val.id == expected_id
        assert nc_val.get_originating_cell() == expected_originating_cell
        assert nc_val.get_all_cells() == expected_cells
        assert nc_val.reference_coordinate_offset == (0, 0)

    def test_deleting_originating_cell_of_shared_expression(self):
        pass
