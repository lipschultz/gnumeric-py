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

from gnumeric import sheet
from gnumeric.exceptions import UnsupportedOperationException
from gnumeric.workbook import Workbook


class SheetTests(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
        self.loaded_wb = Workbook.load_workbook('samples/test.gnumeric')

    def test_creating_sheet_stores_title(self):
        title = 'NewTitle'
        ws = self.wb.create_sheet(title)
        self.assertEqual(ws.title, title)

    def test_changing_title(self):
        title = 'NewTitle'
        ws = self.wb.create_sheet('OldTitle')
        ws.title = title
        self.assertEqual(ws.title, title)

    def test_equality_of_same_worksheet_is_true(self):
        ws = self.wb.create_sheet('Title')
        self.assertTrue(ws == ws)

    def test_different_worksheets_are_not_equal(self):
        ws1 = self.wb.create_sheet('Title1')
        ws2 = self.wb.create_sheet('Title2')
        self.assertFalse(ws1 == ws2)

    def test_worksheets_with_same_name_but_from_different_books_not_equal(self):
        wb2 = Workbook()
        title = 'Title'
        ws = self.wb.create_sheet(title)
        ws2 = wb2.create_sheet(title)
        self.assertFalse(ws == ws2)

    def test_getting_workbook(self):
        ws = self.wb.create_sheet('Title')
        self.assertEqual(ws.workbook, self.wb)

    def test_type_is_regular(self):
        self.assertEqual(self.loaded_wb['Sheet1'].type, sheet.SHEET_TYPE_REGULAR)

    def test_type_is_object(self):
        self.assertEqual(self.loaded_wb['Graph1'].type, sheet.SHEET_TYPE_OBJECT)

    def test_getting_min_row(self):
        ws = self.loaded_wb.get_sheet_by_index(1)
        self.assertEqual(ws.min_row, 6)

    def test_getting_min_col(self):
        ws = self.loaded_wb.get_sheet_by_index(1)
        self.assertEqual(ws.min_column, 3)

    def test_getting_max_row(self):
        ws = self.loaded_wb.get_sheet_by_index(1)
        self.assertEqual(ws.max_row, 12)

    def test_getting_max_col(self):
        ws = self.loaded_wb.get_sheet_by_index(1)
        self.assertEqual(ws.max_column, 9)

    def test_getting_min_row_raises_exception_on_chartsheet(self):
        ws = self.loaded_wb.get_sheet_by_name('Graph1')
        with self.assertRaises(UnsupportedOperationException):
            _ = ws.min_row

    def test_getting_min_col_raises_exception_on_chartsheet(self):
        ws = self.loaded_wb.get_sheet_by_name('Graph1')
        with self.assertRaises(UnsupportedOperationException):
            _ = ws.min_column

    def test_getting_max_row_raises_exception_on_chartsheet(self):
        ws = self.loaded_wb.get_sheet_by_name('Graph1')
        with self.assertRaises(UnsupportedOperationException):
            _ = ws.max_row

    def test_getting_max_col_raises_exception_on_chartsheet(self):
        ws = self.loaded_wb.get_sheet_by_name('Graph1')
        with self.assertRaises(UnsupportedOperationException):
            _ = ws.max_column

    def test_calculate_dimension(self):
        ws = self.loaded_wb.get_sheet_by_index(1)
        self.assertEqual(ws.calculate_dimension(), (6, 3, 12, 9))

    def test_calculate_dimension_raises_exception_on_chartsheet(self):
        ws = self.loaded_wb.get_sheet_by_name('Graph1')
        with self.assertRaises(UnsupportedOperationException):
            ws.calculate_dimension()

    def test_get_cell_at_index(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        c00 = ws.cell(1, 0)
        self.assertEqual(c00.row, 1)
        self.assertEqual(c00.column, 0)
        self.assertEqual(c00.text, '2')

    def test_getitem_returns_cell_at_provided_row_column(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        c00 = ws[1, 0]
        self.assertEqual(c00.row, 1)
        self.assertEqual(c00.column, 0)
        self.assertEqual(c00.text, '2')

    def test_getitem_returns_cell_at_specified_by_spreadsheet_notation(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        c00 = ws['A2']
        self.assertEqual(c00.row, 1)
        self.assertEqual(c00.column, 0)
        self.assertEqual(c00.text, '2')

    def test_get_non_existent_cell_at_index_creates_cell_with_None_text(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        c02 = ws.cell(0, 2)
        self.assertEqual(c02.row, 0)
        self.assertEqual(c02.column, 2)
        self.assertEqual(c02.text, None)

    def test_getting_non_existent_cell_outside_bounding_rectangle_does_not_increase_rectangle(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        ws.cell(15, 30)
        self.assertEqual(ws.calculate_dimension(), old_dimensions)

    def test_getting_non_existent_cell_outside_bounding_rectangle_and_assigning_empty_string_does_not_increase_rectangle(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        c = ws.cell(15, 30)
        c.set_value("")
        self.assertEqual(ws.calculate_dimension(), old_dimensions)

    def test_getting_non_existent_cell_outside_bounding_rectangle_and_assigning_None_does_not_increase_rectangle(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        c = ws.cell(15, 30)
        c.set_value(None)
        self.assertEqual(ws.calculate_dimension(), old_dimensions)

    def test_getting_non_existent_cell_outside_bounding_rectangle_and_assigning_string_does_increase_rectangle(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        c = ws.cell(15, 30)
        c.set_value("new value")
        new_dimensions = old_dimensions[:2] + (15, 30)
        self.assertEqual(ws.calculate_dimension(), new_dimensions)

    def test_getting_non_existent_cell_outside_bounding_rectangle_and_assigning_expression_does_increase_rectangle(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        c = ws.cell(15, 30)
        c.set_value("=max(A1:A5)")
        new_dimensions = old_dimensions[:2] + (15, 30)
        self.assertEqual(ws.calculate_dimension(), new_dimensions)

    def test_get_non_existent_cell_with_create_False_raises_exception(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        with self.assertRaises(IndexError):
            ws.cell(0, 2, create=False)
            # TODO: should also confirm that the cell is not created

    def test_get_cell_text_at_index(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        c00 = ws.cell_text(1, 0)
        self.assertEqual(c00, '2')

    def test_get_text_of_non_existent_cell_raises_exception(self):
        ws = self.loaded_wb.get_sheet_by_name('Sheet1')
        with self.assertRaises(IndexError):
            ws.cell_text(0, 2)
            # TODO: should also confirm that the cell is not created

    def test_get_content_cells(self):
        ws = self.wb.create_sheet("Test")
        cells = set()

        c = ws.cell(0, 0)
        c.set_value('string')
        cells.add(c)

        c = ws.cell(0, 1)
        c.set_value(-17)
        cells.add(c)

        c = ws.cell(0, 2)
        c.set_value(13.4)
        cells.add(c)

        _ = ws.cell(0, 3)

        c = ws.cell(0, 4)
        c.set_value('=max(A1:A5)')
        cells.add(c)

        self.assertEqual(set(ws.get_cell_collection()), cells)

    def test_get_content_cells_including_empty(self):
        ws = self.wb.create_sheet("Test")
        cells = set()

        c = ws.cell(0, 0)
        c.set_value('string')
        cells.add(c)

        c = ws.cell(0, 1)
        c.set_value(-17)
        cells.add(c)

        c = ws.cell(0, 2)
        c.set_value(13.4)
        cells.add(c)

        c = ws.cell(0, 3)
        cells.add(c)

        c = ws.cell(0, 4)
        c.set_value('=max(A1:A5)')
        cells.add(c)

        self.assertEqual(set(ws.get_cell_collection(include_empty=True)), cells)

    def test_get_expression_map_from_worksheet_with_expressions(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')
        expected_map = {'1': ((1, 1), '=sum(A2:A10)'),
                        '2': ((2, 1), '=counta(A$1:A$65536)')
                        }
        self.assertEqual(ws.get_expression_map(), expected_map)

    def test_get_expression_map_from_worksheet_with_no_expressions(self):
        ws = self.loaded_wb.get_sheet_by_name('BoundingRegion')
        expected_map = {}
        self.assertEqual(ws.get_expression_map(), expected_map)

    def test_get_all_cells_with_a_specific_expression_id(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')
        expected_cells = [ws.cell(1, 1), ws.cell(4, 1)]
        self.assertEqual(ws.get_all_cells_with_expression('1', sort='row'), expected_cells)

    def test_get_all_cells_with_a_specific_expression_id_that_does_not_exist(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')
        expected_cells = []
        self.assertEqual(ws.get_all_cells_with_expression("DOESN'T EXIST", sort='row'), expected_cells)

    def test_get_max_allowed_column(self):
        ws = self.loaded_wb.get_sheet_by_index(0)
        self.assertEqual(ws.max_allowed_column, 255)

    def test_get_max_allowed_row(self):
        ws = self.loaded_wb.get_sheet_by_index(0)
        self.assertEqual(ws.max_allowed_row, 65535)

    def test_get_cell_above_allowed_column(self):
        ws = self.wb.create_sheet('Sheet1')
        with self.assertRaises(IndexError):
            ws.cell(0, ws.max_allowed_column + 5)

    def test_get_cell_below_0_column(self):
        ws = self.wb.create_sheet('Sheet1')
        with self.assertRaises(IndexError):
            ws.cell(0, -1)

    def test_get_cell_above_allowed_row(self):
        ws = self.wb.create_sheet('Sheet1')
        with self.assertRaises(IndexError):
            ws.cell(ws.max_allowed_row + 5, 0)

    def test_get_cell_below_0_row(self):
        ws = self.wb.create_sheet('Sheet1')
        with self.assertRaises(IndexError):
            ws.cell(-1, 0)

    def test_valid_column_below_range(self):
        ws = self.wb.create_sheet('Title')
        self.assertFalse(ws.is_valid_column(-1))

    def test_valid_column_at_lower_bound(self):
        ws = self.wb.create_sheet('Title')
        self.assertTrue(ws.is_valid_column(0))

    def test_valid_column_within_range(self):
        ws = self.wb.create_sheet('Title')
        self.assertTrue(ws.is_valid_column(1))

    def test_valid_column_at_upper_bound(self):
        ws = self.wb.create_sheet('Title')
        self.assertTrue(ws.is_valid_column(ws.max_allowed_column))

    def test_valid_column_above_range(self):
        ws = self.wb.create_sheet('Title')
        self.assertFalse(ws.is_valid_column(ws.max_allowed_column + 1))

    def test_valid_row_below_range(self):
        ws = self.wb.create_sheet('Title')
        self.assertFalse(ws.is_valid_row(-1))

    def test_valid_row_at_lower_bound(self):
        ws = self.wb.create_sheet('Title')
        self.assertTrue(ws.is_valid_row(0))

    def test_valid_row_within_range(self):
        ws = self.wb.create_sheet('Title')
        self.assertTrue(ws.is_valid_row(1))

    def test_valid_row_at_upper_bound(self):
        ws = self.wb.create_sheet('Title')
        self.assertTrue(ws.is_valid_row(ws.max_allowed_row))

    def test_valid_row_above_range(self):
        ws = self.wb.create_sheet('Title')
        self.assertFalse(ws.is_valid_row(ws.max_allowed_row + 1))

    def test_get_cells_sorted_row_major(self):
        ws = self.wb.create_sheet('CellOrder')

        all_cells = set()
        cell = ws.cell(2, 2)
        cell.value = "3:C"
        all_cells.add(cell)
        cell = ws.cell(0, 2)
        cell.value = "1:C"
        all_cells.add(cell)
        cell = ws.cell(0, 0)
        cell.value = "1:A"
        all_cells.add(cell)
        cell = ws.cell(0, 1)
        cell.value = "1:B"
        all_cells.add(cell)
        cell = ws.cell(1, 2)
        cell.value = "2:C"
        all_cells.add(cell)
        cell = ws.cell(1, 1)
        cell.value = "2:B"
        all_cells.add(cell)

        ordered_cells = ws.get_cell_collection(sort='row')
        self.assertEqual(set(ordered_cells), all_cells)
        self.assertEqual(ordered_cells[0].row, 0)
        self.assertEqual(ordered_cells[0].column, 0)
        self.assertEqual(ordered_cells[1].row, 0)
        self.assertEqual(ordered_cells[1].column, 1)
        self.assertEqual(ordered_cells[2].row, 0)
        self.assertEqual(ordered_cells[2].column, 2)
        self.assertEqual(ordered_cells[3].row, 1)
        self.assertEqual(ordered_cells[3].column, 1)
        self.assertEqual(ordered_cells[4].row, 1)
        self.assertEqual(ordered_cells[4].column, 2)
        self.assertEqual(ordered_cells[5].row, 2)
        self.assertEqual(ordered_cells[5].column, 2)

    def test_get_cells_sorted_column_major(self):
        ws = self.wb.create_sheet('CellOrder')

        all_cells = set()
        cell = ws.cell(2, 2)
        cell.value = "3:C"
        all_cells.add(cell)
        cell = ws.cell(0, 2)
        cell.value = "1:C"
        all_cells.add(cell)
        cell = ws.cell(0, 0)
        cell.value = "1:A"
        all_cells.add(cell)
        cell = ws.cell(0, 1)
        cell.value = "1:B"
        all_cells.add(cell)
        cell = ws.cell(1, 2)
        cell.value = "2:C"
        all_cells.add(cell)
        cell = ws.cell(1, 1)
        cell.value = "2:B"
        all_cells.add(cell)

        ordered_cells = ws.get_cell_collection(sort='column')
        self.assertEqual(set(ordered_cells), all_cells)
        self.assertEqual(ordered_cells[0].row, 0)
        self.assertEqual(ordered_cells[0].column, 0)
        self.assertEqual(ordered_cells[1].row, 0)
        self.assertEqual(ordered_cells[1].column, 1)
        self.assertEqual(ordered_cells[2].row, 1)
        self.assertEqual(ordered_cells[2].column, 1)
        self.assertEqual(ordered_cells[3].row, 0)
        self.assertEqual(ordered_cells[3].column, 2)
        self.assertEqual(ordered_cells[4].row, 1)
        self.assertEqual(ordered_cells[4].column, 2)
        self.assertEqual(ordered_cells[5].row, 2)
        self.assertEqual(ordered_cells[5].column, 2)

    def test_get_empty_column(self):
        ws = self.wb.create_sheet('Title')
        col = ws.get_column(0)
        col = [c for c in col]

        self.assertEqual(col, [])

    @unittest.skip("takes too long to complete in a reasonable amount of time")
    def test_get_empty_col_and_create_cells_returns_cells_for_all_rows(self):
        ws = self.wb.create_sheet('Title')
        col = ws.get_column(0, create_cells=True)
        self.assertEqual(sum(1 for _ in col), ws.max_allowed_row + 1)

    def test_get_col_with_values_returns_only_existing_cells_in_sorted_order(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = "3:C"
        cell = ws.cell(3, 0)
        cell.value = "4:A"
        cell = ws.cell(0, 0)
        cell.value = "1:A"
        cell = ws.cell(1, 0)
        cell.value = "2:A"

        col = ws.get_column(0)
        self.assertEqual([c.text for c in col], ["1:A", "2:A", "4:A"])

    def test_get_col_within_range_only_returns_cells_within_that_range(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = "3:C"

        cell = ws.cell(3, 0)
        cell.value = "4:A"

        cell = ws.cell(0, 0)
        cell.value = "1:A"

        cell = ws.cell(1, 0)
        cell.value = "2:A"

        col = ws.get_column(0, min_row=1, max_row=2)
        self.assertEqual([c.text for c in col], ["2:A"])

    def test_get_col_and_create_cells_will_return_all_cells_in_sorted_order(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = "3:C"

        cell = ws.cell(3, 0)
        cell.value = "4:A"

        cell = ws.cell(0, 0)
        cell.value = "1:A"

        cell = ws.cell(1, 0)
        cell.value = "2:A"

        col = ws.get_column(0, max_row=10, create_cells=True)
        self.assertEqual([c.text for c in col], ["1:A", "2:A", None, "4:A"] + [None] * 7)

    def test_max_row_in_empty_row_is_negative_one(self):
        ws = self.wb.create_sheet('Title')
        self.assertEqual(ws.max_row_in_column(0), -1)

    def test_min_row_in_empty_row_is_negative_one(self):
        ws = self.wb.create_sheet('Title')
        self.assertEqual(ws.min_row_in_column(0), -1)

    def test_max_row_in_column(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = "3:C"

        cell = ws.cell(3, 0)
        cell.value = "4:A"

        cell = ws.cell(0, 0)
        cell.value = "1:A"

        cell = ws.cell(1, 0)
        cell.value = "2:A"

        self.assertEqual(ws.max_row_in_column(0), 3)

    def test_min_row_in_column(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = "3:C"

        cell = ws.cell(3, 0)
        cell.value = "4:A"

        cell = ws.cell(0, 0)
        cell.value = "1:A"

        cell = ws.cell(1, 0)
        cell.value = "2:A"

        self.assertEqual(ws.min_row_in_column(0), 0)

    def test_get_empty_row(self):
        ws = self.wb.create_sheet('Title')
        row = ws.get_row(0)
        row = [r for r in row]

        self.assertEqual(row, [])

    def test_get_empty_row_and_create_cells_returns_cells_for_all_columns(self):
        ws = self.wb.create_sheet('Title')
        row = ws.get_row(0, create_cells=True)
        row = [r for r in row]

        self.assertEqual(len(row), ws.max_allowed_column + 1)

    def test_get_row_with_values_returns_only_existing_cells_in_sorted_order(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 0)
        cell.value = "3:C"
        cell = ws.cell(0, 3)
        cell.value = "1:D"
        cell = ws.cell(0, 0)
        cell.value = "1:A"
        cell = ws.cell(0, 1)
        cell.value = "1:B"

        row = ws.get_row(0)
        self.assertEqual([r.text for r in row], ["1:A", "1:B", "1:D"])

    def test_get_row_within_range_only_returns_cells_within_that_range(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 0)
        cell.value = "3:C"

        cell = ws.cell(0, 3)
        cell.value = "1:D"

        cell = ws.cell(0, 0)
        cell.value = "1:A"

        cell = ws.cell(0, 1)
        cell.value = "1:B"

        row = ws.get_row(0, min_col=1, max_col=2)
        self.assertEqual([r.text for r in row], ["1:B"])

    def test_get_row_and_create_cells_will_return_all_cells_in_sorted_order(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = "3:C"
        cell = ws.cell(0, 3)
        cell.value = "1:D"
        cell = ws.cell(0, 0)
        cell.value = "1:A"
        cell = ws.cell(0, 1)
        cell.value = "1:B"

        row = ws.get_row(0, max_col=10, create_cells=True)
        self.assertEqual([r.text for r in row], ["1:A", "1:B", None, "1:D"] + [None] * 7)

    def test_max_column_in_empty_row_is_negative_one(self):
        ws = self.wb.create_sheet('Title')
        self.assertEqual(ws.max_column_in_row(0), -1)

    def test_min_column_in_empty_row_is_negative_one(self):
        ws = self.wb.create_sheet('Title')
        self.assertEqual(ws.min_column_in_row(0), -1)

    def test_max_column_in_row(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 0)
        cell.value = "3:C"

        cell = ws.cell(0, 3)
        cell.value = "1:D"

        cell = ws.cell(0, 0)
        cell.value = "1:A"

        cell = ws.cell(0, 1)
        cell.value = "1:B"

        self.assertEqual(ws.max_column_in_row(0), 3)

    def test_min_column_in_row(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 0)
        cell.value = "3:C"

        cell = ws.cell(0, 3)
        cell.value = "1:D"

        cell = ws.cell(0, 0)
        cell.value = "1:A"

        cell = ws.cell(0, 1)
        cell.value = "1:B"

        self.assertEqual(ws.min_column_in_row(0), 0)

    def test_removing_worksheet_deletes_it_from_workbook(self):
        [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        orig_num_worksheets = len(self.wb)

        ws2 = self.wb.get_sheet_by_name('Title2')

        ws2.remove_from_workbook()
        self.assertEqual(len(self.wb), orig_num_worksheets - 1)

    def test_removing_worksheet_keeps_other_worksheets(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(3)]

        ws1 = self.wb.get_sheet_by_name('Title1')
        ws1.remove_from_workbook()

        self.assertEqual(self.wb.get_sheet_by_name('Title0'), worksheets[0])
        self.assertEqual(self.wb.get_sheet_by_name('Title0').title, worksheets[0].title)
        self.assertEqual(self.wb.get_sheet_by_name('Title2'), worksheets[2])
        self.assertEqual(self.wb.get_sheet_by_name('Title2').title, worksheets[2].title)

    def test_removing_non_existing_cell_does_not_delete_cells(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        ws.delete_cell(10, 10)
        self.assertEqual(ws.get_cell_collection(), expected_cells)

    def test_removing_existing_non_expression_cell(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        expected_cells.remove(ws.cell(0, 1))
        ws.delete_cell(0, 1)
        self.assertEqual(ws.get_cell_collection(), expected_cells)

    def test_removing_non_shared_expression_cell(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        expected_cells.remove(ws.cell(3, 1))
        ws.delete_cell(3, 1)
        self.assertEqual(ws.get_cell_collection(), expected_cells)

    def test_removing_shared_expression_non_originating_cell(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        expected_cells.remove(ws.cell(4, 1))
        ws.delete_cell(4, 1)
        self.assertEqual(ws.get_cell_collection(), expected_cells)

    def test_removing_shared_expression_originating_cell_raises_exception(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        with self.assertRaises(UnsupportedOperationException):
            ws.delete_cell(1, 1)
        self.assertEqual(ws.get_cell_collection(), expected_cells)

    @unittest.skip('Still need to implement inserting cells/rows/columns')
    def test_inserting_cell_before_another_cell_does_not_cause_the_existing_cell_to_be_recreated(self):
        # This is needed because of the singleton nature of the Cell class
        self.assertTrue(False)

    @unittest.skip('Still need to implement removing cells/rows/columns')
    def test_removing_cell_before_another_cell_does_not_cause_the_existing_cell_to_be_recreated(self):
        # This is needed because of the singleton nature of the Cell class
        self.assertTrue(False)
