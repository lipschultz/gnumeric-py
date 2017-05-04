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
from datetime import datetime

from dateutil.tz import tzutc

from src.workbook import Workbook
from src import sheet, cell
from src.exceptions import DuplicateTitleException, WrongWorkbookException, UnsupportedOperationException


class WorkbookTests(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
        self.loaded_wb = Workbook.load_workbook('samples/test.gnumeric')

    def test_equality_of_same_workbook(self):
        self.assertTrue(self.wb == self.wb)

    def test_different_workbooks_are_not_equal(self):
        wb2 = Workbook()
        self.assertFalse(self.wb == wb2)

    def test_creating_empty_workbook_has_zero_sheets(self):
        self.assertEqual(len(self.wb), 0)
        self.assertEqual(len(self.wb.get_sheet_names()), 0)

    def test_creating_workbook_has_version_1_12_28(self):
        self.assertEqual(self.wb.version, "1.12.28")

    def test_creating_sheet_in_empty_book_adds_sheet_to_book(self):
        title = 'Title'
        ws = self.wb.create_sheet(title)
        self.assertEqual(len(self.wb), 1)
        self.assertEqual(self.wb.get_sheet_names(), [title])

    def test_creating_sheet_with_name_used_by_another_sheet_raises_exception(self):
        title = 'Title'
        self.wb.create_sheet(title)
        with self.assertRaises(DuplicateTitleException):
            self.wb.create_sheet(title)

    def test_creating_sheet_appends_to_list_of_sheets(self):
        titles = ['Title1', 'Title2']
        for title in titles:
            ws = self.wb.create_sheet(title)
        self.assertEqual(len(self.wb), 2)
        self.assertEqual(self.wb.get_sheet_names(), titles)

    def test_prepending_new_sheet(self):
        titles = ['Title1', 'Title2']
        for title in titles:
            ws = self.wb.create_sheet(title)
        self.assertEqual(self.wb.get_sheet_names(), titles)

    def test_inserting_new_sheet(self):
        titles = ['Title1', 'Title3', 'Title2']
        ws = self.wb.create_sheet(titles[0])
        ws = self.wb.create_sheet(titles[2])
        ws = self.wb.create_sheet(titles[1], 1)
        self.assertEqual(self.wb.get_sheet_names(), titles)

    def test_creation_date_on_new_workbook(self):
        before = datetime.now()
        wb = Workbook()
        after = datetime.now()
        self.assertTrue(before < wb.creation_date < after)

    def test_setting_creation_date(self):
        creation_date = datetime.now()
        self.wb.creation_date = creation_date
        self.assertEqual(self.wb.creation_date, creation_date)

    def test_getting_sheet_by_index_with_positive_index(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = 3
        ws = self.wb.get_sheet_by_index(index)
        self.assertEqual(ws, worksheets[index])

    def test_getting_sheet_by_index_with_negative_index(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = -2
        ws = self.wb.get_sheet_by_index(index)
        self.assertEqual(ws, worksheets[index])

    def test_getting_sheet_out_of_bounds_raises_exception(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        with self.assertRaises(IndexError):
            self.wb.get_sheet_by_index(len(worksheets) + 10)

    def test_getting_sheet_by_name(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws = self.wb.get_sheet_by_name('Title' + str(index))
        self.assertEqual(ws, worksheets[index])

    def test_getting_sheet_with_nonexistent_name_raises_exception(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        with self.assertRaises(KeyError):
            self.wb.get_sheet_by_name('NotASheet')

    def test_getting_sheet_with_getitem_using_positive_index(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = 3
        ws = self.wb[index]
        self.assertEqual(ws, worksheets[index])

    def test_getting_sheet_with_getitem_using_negative_index(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = -2
        ws = self.wb[index]
        self.assertEqual(ws, worksheets[index])

    def test_getting_sheet_with_getitem_out_of_bounds_raises_exception(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        with self.assertRaises(IndexError):
            self.wb[len(worksheets) + 10]

    def test_getting_sheet_with_getitem_using_bad_key_type_raises_exception(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        with self.assertRaises(TypeError):
            self.wb[2.0]

    def test_getting_sheet_with_getitem_by_name(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws = self.wb['Title' + str(index)]
        self.assertEqual(ws, worksheets[index])

    def test_getting_sheet_with_getitem_using_nonexistent_name_raises_exception(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        with self.assertRaises(KeyError):
            self.wb['NotASheet']

    def test_getting_index_of_worksheet(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        self.assertEqual(self.wb.get_index(ws2), index)

    def test_getting_index_of_worksheet_from_different_workbook_but_with_name_raises_exception(self):
        wb2 = Workbook()
        title = 'Title'
        ws = self.wb.create_sheet(title)
        ws2 = wb2.create_sheet(title)
        with self.assertRaises(WrongWorkbookException):
            self.wb.get_index(ws2)

    def test_deleting_worksheet(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        title = ws2.title
        self.wb.remove_sheet(ws2)
        self.assertEqual(len(self.wb), len(worksheets) - 1)
        self.assertEqual(len(self.wb.sheetnames), len(worksheets) - 1)
        self.assertNotIn(title, self.wb.sheetnames)
        self.assertTrue(all(ws2 != self.wb[i] for i in range(len(self.wb))))

    def test_deleting_worksheet_twice_raises_exception(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        title = ws2.title
        self.wb.remove_sheet(ws2)
        with self.assertRaises(ValueError):
            self.wb.remove_sheet(ws2)

    def test_deleting_worksheet_from_another_workbook_raises_exception(self):
        wb2 = Workbook()
        [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        ws2 = wb2.create_sheet('Title2')

        with self.assertRaises(ValueError):
            self.wb.remove_sheet(ws2)

    def test_deleting_worksheet_by_name(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        title = ws2.title
        self.wb.remove_sheet_by_name(title)
        self.assertEqual(len(self.wb), len(worksheets) - 1)
        self.assertEqual(len(self.wb.sheetnames), len(worksheets) - 1)
        self.assertNotIn(title, self.wb.sheetnames)
        self.assertTrue(all(ws2 != self.wb[i] for i in range(len(self.wb))))

    def test_deleting_worksheet_by_non_existent_name_raises_exception(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        title = "NotATitle"

        with self.assertRaises(KeyError):
            self.wb.remove_sheet_by_name(title)

        self.assertEqual(len(self.wb), len(worksheets))
        self.assertEqual(len(self.wb.sheetnames), len(worksheets))

    def test_deleting_worksheet_by_index(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        title = ws2.title
        self.wb.remove_sheet_by_index(index)
        self.assertEqual(len(self.wb), len(worksheets) - 1)
        self.assertEqual(len(self.wb.sheetnames), len(worksheets) - 1)
        self.assertNotIn(title, self.wb.sheetnames)
        self.assertTrue(all(ws2 != self.wb[i] for i in range(len(self.wb))))

    def test_deleting_worksheet_by_out_of_bounds_index_raises_exception(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        index = len(worksheets) + 10

        with self.assertRaises(IndexError):
            self.wb.remove_sheet_by_index(index)

        self.assertEqual(len(self.wb), len(worksheets))
        self.assertEqual(len(self.wb.sheetnames), len(worksheets))

    def test_getting_all_sheets(self):
        worksheets = [self.wb.create_sheet('Title' + str(i)) for i in range(5)]
        self.assertEqual(self.wb.sheets, worksheets)

    def test_getting_all_sheets_of_mixed_type(self):
        ws = self.loaded_wb.sheets
        selected_ws = [self.loaded_wb.get_sheet_by_name(n) for n in
                       ('Sheet1', 'BoundingRegion', 'CellTypes', 'Expressions', 'Mine & Yours Sheet[s]!', 'Graph1')]
        self.assertEqual(ws, selected_ws)

    def test_getting_worksheets_only(self):
        ws = self.loaded_wb.worksheets
        selected_ws = [self.loaded_wb.get_sheet_by_name(n) for n in
                       ('Sheet1', 'BoundingRegion', 'CellTypes', 'Expressions', 'Mine & Yours Sheet[s]!')]
        self.assertEqual(ws, selected_ws)

    def test_getting_chartsheets_only(self):
        ws = self.loaded_wb.chartsheets
        selected_ws = [self.loaded_wb.get_sheet_by_name(n) for n in ('Graph1',)]
        self.assertEqual(ws, selected_ws)

    def test_loading_compressed_file(self):
        self.assertEqual(self.loaded_wb.sheetnames,
                         ['Sheet1', 'BoundingRegion', 'CellTypes', 'Expressions', 'Mine & Yours Sheet[s]!', 'Graph1'])
        self.assertEqual(self.loaded_wb.creation_date, datetime(2017, 4, 29, 17, 56, 48, tzinfo=tzutc()))
        self.assertEqual(self.loaded_wb.version, '1.12.28')

    def test_loading_uncompressed_file(self):
        wb = Workbook.load_workbook('samples/sheet-names.xml')
        self.assertEqual(wb.sheetnames, ['Sheet1', 'Sheet2', 'Sheet3', 'Mine & Yours Sheet[s]!', 'Graph1'])
        self.assertEqual(wb.creation_date, datetime(2017, 4, 29, 17, 56, 48, tzinfo=tzutc()))
        self.assertEqual(wb.version, '1.12.28')


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
            val = ws.min_row

    def test_getting_min_col_raises_exception_on_chartsheet(self):
        ws = self.loaded_wb.get_sheet_by_name('Graph1')
        with self.assertRaises(UnsupportedOperationException):
            val = ws.min_column

    def test_getting_max_row_raises_exception_on_chartsheet(self):
        ws = self.loaded_wb.get_sheet_by_name('Graph1')
        with self.assertRaises(UnsupportedOperationException):
            val = ws.max_row

    def test_getting_max_col_raises_exception_on_chartsheet(self):
        ws = self.loaded_wb.get_sheet_by_name('Graph1')
        with self.assertRaises(UnsupportedOperationException):
            val = ws.max_column

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
            ws.cell(0, 2, False)
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

        c = ws.cell(0, 3)

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

        self.assertEqual(set(ws.get_cell_collection(True)), cells)

    def test_get_expression_map_from_worksheet_with_expressions(self):
        ws = self.loaded_wb.get_sheet_by_name('Expressions')
        expected_map = {'1': ((1, 1), '=sum(A2:A10)'),
                        '2': ((2, 1), '=counta(A$1:A$65536)'),
                        '3': ((3, 1), '=product(BoundingRegion!D7:J13)')
                        }
        self.assertEqual(ws.get_expression_map(), expected_map)

    def test_get_expression_map_from_worksheet_with_no_expressions(self):
        ws = self.loaded_wb.get_sheet_by_name('BoundingRegion')
        expected_map = {}
        self.assertEqual(ws.get_expression_map(), expected_map)

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
        self.assertEquals([r.text for r in row], ["1:A", "1:B", "1:D"])

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

        row = ws.get_row(0, 1, 2)
        self.assertEquals([r.text for r in row], ["1:B"])

    def test_get_row_and_create_cells_will_return_all_cells_in_sorted_order(self):
        ws = self.wb.create_sheet('Title')

        cell = ws.cell(2, 0)
        cell.value = "3:C"
        cell = ws.cell(0, 3)
        cell.value = "1:D"
        cell = ws.cell(0, 0)
        cell.value = "1:A"
        cell = ws.cell(0, 1)
        cell.value = "1:B"

        row = ws.get_row(0, max_col=10, create_cells=True)
        self.assertEquals([r.text for r in row], ["1:A", "1:B", None, "1:D"] + [None]*7)

    def test_max_column_in_empty_row_is_negative_one(self):
        ws = self.wb.create_sheet('Title')
        self.assertEqual(ws.max_column_in_row(0), -1)

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
            self.assertEqual(test_cell.type, expected_type, str(test_cell.text) + ' (row=' + str(row) + ') has type '
                             + str(test_cell.type) + ', but expected ' + str(expected_type))

    def test_get_cell_values(self):
        ws = self.loaded_wb.get_sheet_by_name('CellTypes')
        for row in range(ws.max_row + 1):
            test_cell = ws.cell(row, 1)

            expected_value = test_cell.text
            if test_cell.type == cell.VALUE_TYPE_INTEGER:
                expected_value = int(expected_value)
            elif test_cell.type == cell.VALUE_TYPE_FLOAT:
                expected_value = float(expected_value)
            elif test_cell.type == cell.VALUE_TYPE_BOOLEAN:
                expected_value = bool(expected_value)
            elif test_cell.type == cell.VALUE_TYPE_EMPTY:
                expected_value = None

            self.assertEqual(test_cell.get_value(), expected_value,
                             str(test_cell.text) + ' (row=' + str(row) + ') has type ' + str(test_cell.type)
                             + ', but expected ' + str(expected_value))

    def test_set_cell_value_infer_bool(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(True)
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_BOOLEAN)

    def test_set_cell_value_infer_int(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(10)
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_INTEGER)

    def test_set_cell_value_infer_float(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(10.0)
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_FLOAT)

    def test_new_cell_is_empty(self):
        test_cell = self.ws.cell(0, 0)
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_EMPTY)

    def test_set_cell_value_infer_empty_from_empty_string(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('')
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_EMPTY)

    def test_set_cell_value_infer_empty_from_None(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(None)
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_EMPTY)

    def test_set_cell_value_infer_expression(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('=max()')
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_EXPR)

    def test_set_cell_value_infer_string(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('asdf')
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_STRING)

    def test_cell_type_changes_when_value_changes(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('asdf')
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_STRING)
        test_cell.set_value(17)
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_INTEGER)

    def test_save_int_into_cell_value_using_keep_when_cell_value_is_string(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value('asdf')
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_STRING)
        test_cell.set_value(17, value_type='keep')
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_STRING)

    def test_explicitly_setting_value_type_uses_that_value_type(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.set_value(17, cell.VALUE_TYPE_STRING)
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_STRING)

    def test_setting_value_infers_type(self):
        test_cell = self.ws.cell(0, 0)
        test_cell.value = 17
        self.assertEqual(test_cell.type, cell.VALUE_TYPE_INTEGER)

    def test_cells_equal(self):
        test_cell1 = self.ws.cell(0, 0)
        test_cell2 = self.ws.cell(0, 0)
        self.assertTrue(test_cell1 == test_cell2)

    def test_cells_unequal(self):
        test_cell1 = self.ws.cell(0, 0)
        test_cell2 = self.ws.cell(1, 0)
        self.assertFalse(test_cell1 == test_cell2)
