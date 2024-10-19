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

import pytest

from gnumeric import sheet
from gnumeric.exceptions import UnsupportedOperationException
from gnumeric.workbook import Workbook

TEST_GNUMERIC_FILE_PATH = 'samples/test.gnumeric'


class TestSheetMetadata:
    def test_creating_sheet_stores_title(self):
        workbook = Workbook()
        title = 'NewTitle'
        ws = workbook.create_sheet(title)
        assert ws.title == title

    def test_changing_title(self):
        workbook = Workbook()
        title = 'NewTitle'
        ws = workbook.create_sheet('OldTitle')
        ws.title = title
        assert ws.title == title

    def test_equality_of_same_worksheet_is_true(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws == ws

    def test_different_worksheets_are_not_equal(self):
        workbook = Workbook()
        ws1 = workbook.create_sheet('Title1')
        ws2 = workbook.create_sheet('Title2')
        assert ws1 != ws2

    def test_worksheets_with_same_name_but_from_different_books_not_equal(self):
        workbook = Workbook()
        wb2 = Workbook()
        title = 'Title'
        ws = workbook.create_sheet(title)
        ws2 = wb2.create_sheet(title)
        assert ws != ws2

    def test_getting_workbook(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.workbook == workbook

    def test_type_is_regular(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        assert workbook['Sheet1'].type == sheet.SHEET_TYPE_REGULAR

    def test_type_is_object(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        assert workbook['Graph1'].type == sheet.SHEET_TYPE_OBJECT

    def test_getting_min_row(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_index(1)
        assert ws.min_row == 6

    def test_getting_min_col(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_index(1)
        assert ws.min_column == 3

    def test_getting_max_row(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_index(1)
        assert ws.max_row == 12

    def test_getting_max_col(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_index(1)
        assert ws.max_column == 9

    def test_getting_min_row_raises_exception_on_chartsheet(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Graph1')
        with pytest.raises(UnsupportedOperationException):
            _ = ws.min_row

    def test_getting_min_col_raises_exception_on_chartsheet(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Graph1')
        with pytest.raises(UnsupportedOperationException):
            _ = ws.min_column

    def test_getting_max_row_raises_exception_on_chartsheet(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Graph1')
        with pytest.raises(UnsupportedOperationException):
            _ = ws.max_row

    def test_getting_max_col_raises_exception_on_chartsheet(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Graph1')
        with pytest.raises(UnsupportedOperationException):
            _ = ws.max_column

    def test_calculate_dimension(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_index(1)
        assert ws.calculate_dimension() == (6, 3, 12, 9)

    def test_calculate_dimension_raises_exception_on_chartsheet(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Graph1')
        with pytest.raises(UnsupportedOperationException):
            ws.calculate_dimension()


class TestAccessCell:
    @classmethod
    def assert_equal_cell_sets_by_coordinates(
        cls, actual_cells, *, expected=None, expected_start=None, expected_end=None
    ):
        if expected_start is not None and expected_end is not None:
            start_row, start_col = expected_start
            end_row, end_col = expected_end
            expected_coordinates = {
                (r, c)
                for r in range(start_row, end_row + 1)
                for c in range(start_col, end_col + 1)
            }
            return cls.assert_equal_cell_sets_by_coordinates(
                actual_cells, expected=expected_coordinates
            )
        else:
            actual_coordinates = {(c.row, c.column) for c in actual_cells}
            assert set(expected) == actual_coordinates

    def test_get_cell_at_index(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        c00 = ws.cell(1, 0)
        assert c00.row == 1
        assert c00.column == 0
        assert c00.text == '2'

    def test_getitem_returns_cell_at_provided_row_column(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        c00 = ws[1, 0]
        assert c00.row == 1
        assert c00.column == 0
        assert c00.text == '2'

    def test_getitem_returns_cell_at_specified_by_spreadsheet_notation(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        c00 = ws['A2']
        assert c00.row == 1
        assert c00.column == 0
        assert c00.text == '2'

    def test_get_non_existent_cell_at_index_creates_cell_with_None_text(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        c02 = ws.cell(0, 2)
        assert c02.row == 0
        assert c02.column == 2
        assert c02.text is None

    def test_getting_non_existent_cell_outside_bounding_rectangle_does_not_increase_rectangle(
        self,
    ):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        ws.cell(15, 30)
        assert ws.calculate_dimension() == old_dimensions

    def test_getting_non_existent_cell_outside_bounding_rectangle_and_assigning_empty_string_does_not_increase_rectangle(
        self,
    ):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        c = ws.cell(15, 30)
        c.set_value('')
        assert ws.calculate_dimension() == old_dimensions

    def test_getting_non_existent_cell_outside_bounding_rectangle_and_assigning_None_does_not_increase_rectangle(
        self,
    ):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        c = ws.cell(15, 30)
        c.set_value(None)
        assert ws.calculate_dimension() == old_dimensions

    def test_getting_non_existent_cell_outside_bounding_rectangle_and_assigning_string_does_increase_rectangle(
        self,
    ):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        c = ws.cell(15, 30)
        c.set_value('new value')
        new_dimensions = old_dimensions[:2] + (15, 30)
        assert ws.calculate_dimension() == new_dimensions

    def test_getting_non_existent_cell_outside_bounding_rectangle_and_assigning_expression_does_increase_rectangle(
        self,
    ):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        old_dimensions = ws.calculate_dimension()
        c = ws.cell(15, 30)
        c.set_value('=max(A1:A5)')
        new_dimensions = old_dimensions[:2] + (15, 30)
        assert ws.calculate_dimension() == new_dimensions

    def test_get_non_existent_cell_with_create_False_raises_exception(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        with pytest.raises(IndexError):
            ws.cell(0, 2, create=False)
            # TODO: should also confirm that the cell is not created

    def test_get_cell_text_at_index(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        c00 = ws.cell_text(1, 0)
        assert c00 == '2'

    def test_get_text_of_non_existent_cell_raises_exception(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Sheet1')
        with pytest.raises(IndexError):
            ws.cell_text(0, 2)
            # TODO: should also confirm that the cell is not created

    def test_get_content_cells(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Test')
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

        assert set(ws.get_cell_collection()) == cells

    def test_get_content_cells_including_empty(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Test')
        expected_cells = set()

        c = ws.cell(0, 0)
        c.set_value('string')
        expected_cells.add(c.coordinate)

        c = ws.cell(0, 1)
        c.set_value(-17)
        expected_cells.add(c.coordinate)

        c = ws.cell(0, 2)
        c.set_value(13.4)
        expected_cells.add(c.coordinate)

        c = ws.cell(0, 3)
        expected_cells.add(c.coordinate)

        c = ws.cell(0, 4)
        c.set_value('=max(A1:A5)')
        expected_cells.add(c.coordinate)
        expected_cells.add((1, 0))
        expected_cells.add((2, 0))
        expected_cells.add((3, 0))

        c = ws.cell(4, 0)
        expected_cells.add(c.coordinate)

        actual = ws.get_cell_collection(include_empty=True)
        self.assert_equal_cell_sets_by_coordinates(actual, expected=expected_cells)

    def test_get_range_of_cells_by_passing_int_indices(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Test')
        cells = set()

        c = ws.cell(0, 0)
        c.set_value('string')

        c = ws.cell(0, 1)
        c.set_value(-17)
        cells.add(c)

        c = ws.cell(0, 2)
        c.set_value(13.4)
        cells.add(c)

        c = ws.cell(0, 3)

        c = ws.cell(0, 4)
        c.set_value('last')

        assert set(ws.get_cell_collection((0, 1), (0, 3))) == cells

    def test_get_range_of_cells_by_passing_coordinates(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Test')
        cells = set()

        c = ws.cell(0, 0)
        c.set_value('string')

        c = ws.cell(0, 1)
        c.set_value(-17)
        cells.add(c)

        c = ws.cell(0, 2)
        c.set_value(13.4)
        cells.add(c)

        c = ws.cell(0, 3)

        c = ws.cell(0, 4)
        c.set_value('last')

        assert set(ws.get_cell_collection('B1', 'D1')) == cells

    def test_get_range_of_cells_by_boundary_cells(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Test')
        cells = set()

        c = ws.cell(0, 0)
        c.set_value('string')

        c = ws.cell(0, 1)
        c.set_value(-17)
        cells.add(c)
        start_cell = c

        c = ws.cell(0, 2)
        c.set_value(13.4)
        cells.add(c)

        c = ws.cell(0, 3)
        end_cell = c

        c = ws.cell(0, 4)
        c.set_value('last')

        assert set(ws.get_cell_collection(start_cell, end_cell)) == cells

    def test_get_range_of_cells_by_boundary_cells_and_including_empties_and_creating_cells(
        self,
    ):
        workbook = Workbook()
        ws = workbook.create_sheet('Test')

        start_coord = (0, 0)
        end_coord = (2, 5)

        c = ws.cell(*start_coord)
        start_cell = c

        c = ws.cell(*end_coord)
        end_cell = c

        actual_cell_collection = ws.get_cell_collection(
            start_cell, end_cell, include_empty=True, create_cells=True
        )
        self.assert_equal_cell_sets_by_coordinates(
            actual_cell_collection, expected_start=start_coord, expected_end=end_coord
        )

    def test_get_range_of_cells_leaving_off_end_cell_returns_all_cells_greater_than_start_cells_position(
        self,
    ):
        workbook = Workbook()
        ws = workbook.create_sheet('Test')
        cells = set()

        c = ws.cell(0, 0)
        c.set_value('string')

        c = ws.cell(0, 1)
        c.set_value(-17)
        cells.add(c)
        start_cell = c

        c = ws.cell(0, 2)
        c.set_value(13.4)
        cells.add(c)

        c = ws.cell(0, 3)

        c = ws.cell(0, 4)
        c.set_value('last')
        cells.add(c)

        assert set(ws.get_cell_collection(start_cell)) == cells

    def test_get_range_of_cells_leaving_offstart_cell_returns_all_cells_less_than_end_cells_position(
        self,
    ):
        workbook = Workbook()
        ws = workbook.create_sheet('Test')
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
        end_cell = c

        c = ws.cell(0, 3)

        c = ws.cell(0, 4)
        c.set_value('last')

        assert set(ws.get_cell_collection(end=end_cell)) == cells

    def test_get_expression_map_from_worksheet_with_expressions(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')
        expected_map = {
            '1': ((1, 1), '=sum(A2:A10)'),
            '2': ((2, 1), '=counta(A$1:A$65536)'),
        }
        assert ws.get_expression_map() == expected_map

    def test_get_expression_map_from_worksheet_with_no_expressions(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('BoundingRegion')
        expected_map = {}
        assert ws.get_expression_map() == expected_map

    def test_get_all_cells_with_a_specific_expression_id(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')
        expected_cells = [ws.cell(1, 1), ws.cell(4, 1)]
        assert ws.get_all_cells_with_expression('1', sort='row') == expected_cells

    def test_get_all_cells_with_a_specific_expression_id_that_does_not_exist(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')
        expected_cells = []
        assert (
            ws.get_all_cells_with_expression("DOESN'T EXIST", sort='row')
            == expected_cells
        )

    def test_get_max_allowed_column(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_index(0)
        assert ws.max_allowed_column == 255

    def test_get_max_allowed_row(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_index(0)
        assert ws.max_allowed_row == 65535

    def test_get_cell_above_allowed_column(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Sheet1')
        with pytest.raises(IndexError):
            ws.cell(0, ws.max_allowed_column + 5)

    def test_get_cell_below_0_column(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Sheet1')
        with pytest.raises(IndexError):
            ws.cell(0, -1)

    def test_get_cell_above_allowed_row(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Sheet1')
        with pytest.raises(IndexError):
            ws.cell(ws.max_allowed_row + 5, 0)

    def test_get_cell_below_0_row(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Sheet1')
        with pytest.raises(IndexError):
            ws.cell(-1, 0)

    def test_valid_column_below_range(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_column(-1) is False

    def test_valid_column_at_lower_bound(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_column(0) is True

    def test_valid_column_within_range(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_column(1) is True

    def test_valid_column_at_upper_bound(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_column(ws.max_allowed_column) is True

    def test_valid_column_above_range(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_column(ws.max_allowed_column + 1) is False

    def test_valid_row_below_range(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_row(-1) is False

    def test_valid_row_at_lower_bound(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_row(0) is True

    def test_valid_row_within_range(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_row(1) is True

    def test_valid_row_at_upper_bound(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_row(ws.max_allowed_row) is True

    def test_valid_row_above_range(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.is_valid_row(ws.max_allowed_row + 1) is False

    def test_get_cells_sorted_row_major(self):
        workbook = Workbook()
        ws = workbook.create_sheet('CellOrder')

        all_cells = set()
        cell = ws.cell(2, 2)
        cell.value = '3:C'
        all_cells.add(cell)
        cell = ws.cell(0, 2)
        cell.value = '1:C'
        all_cells.add(cell)
        cell = ws.cell(0, 0)
        cell.value = '1:A'
        all_cells.add(cell)
        cell = ws.cell(0, 1)
        cell.value = '1:B'
        all_cells.add(cell)
        cell = ws.cell(1, 2)
        cell.value = '2:C'
        all_cells.add(cell)
        cell = ws.cell(1, 1)
        cell.value = '2:B'
        all_cells.add(cell)

        ordered_cells = ws.get_cell_collection(sort='row')
        assert set(ordered_cells) == all_cells
        assert ordered_cells[0].row == 0
        assert ordered_cells[0].column == 0
        assert ordered_cells[1].row == 0
        assert ordered_cells[1].column == 1
        assert ordered_cells[2].row == 0
        assert ordered_cells[2].column == 2
        assert ordered_cells[3].row == 1
        assert ordered_cells[3].column == 1
        assert ordered_cells[4].row == 1
        assert ordered_cells[4].column == 2
        assert ordered_cells[5].row == 2
        assert ordered_cells[5].column == 2

    def test_get_cells_sorted_column_major(self):
        workbook = Workbook()
        ws = workbook.create_sheet('CellOrder')

        all_cells = set()
        cell = ws.cell(2, 2)
        cell.value = '3:C'
        all_cells.add(cell)
        cell = ws.cell(0, 2)
        cell.value = '1:C'
        all_cells.add(cell)
        cell = ws.cell(0, 0)
        cell.value = '1:A'
        all_cells.add(cell)
        cell = ws.cell(0, 1)
        cell.value = '1:B'
        all_cells.add(cell)
        cell = ws.cell(1, 2)
        cell.value = '2:C'
        all_cells.add(cell)
        cell = ws.cell(1, 1)
        cell.value = '2:B'
        all_cells.add(cell)

        ordered_cells = ws.get_cell_collection(sort='column')
        assert set(ordered_cells) == all_cells
        assert ordered_cells[0].row == 0
        assert ordered_cells[0].column == 0
        assert ordered_cells[1].row == 0
        assert ordered_cells[1].column == 1
        assert ordered_cells[2].row == 1
        assert ordered_cells[2].column == 1
        assert ordered_cells[3].row == 0
        assert ordered_cells[3].column == 2
        assert ordered_cells[4].row == 1
        assert ordered_cells[4].column == 2
        assert ordered_cells[5].row == 2
        assert ordered_cells[5].column == 2

    def test_get_empty_column(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        col = ws.get_column(0)
        col = [c for c in col]

        assert col == []

    @pytest.mark.skip('takes too long to complete in a reasonable amount of time')
    def test_get_empty_col_and_create_cells_returns_cells_for_all_rows(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        col = ws.get_column(0, create_cells=True)
        assert sum(1 for _ in col) == ws.max_allowed_row + 1

    def test_get_col_with_values_returns_only_existing_cells_in_sorted_order(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = '3:C'
        cell = ws.cell(3, 0)
        cell.value = '4:A'
        cell = ws.cell(0, 0)
        cell.value = '1:A'
        cell = ws.cell(1, 0)
        cell.value = '2:A'

        col = ws.get_column(0)
        assert [c.text for c in col] == ['1:A', '2:A', '4:A']

    def test_get_col_within_range_only_returns_cells_within_that_range(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = '3:C'

        cell = ws.cell(3, 0)
        cell.value = '4:A'

        cell = ws.cell(0, 0)
        cell.value = '1:A'

        cell = ws.cell(1, 0)
        cell.value = '2:A'

        col = ws.get_column(0, min_row=1, max_row=2)
        assert [c.text for c in col] == ['2:A']

    def test_get_col_and_create_cells_will_return_all_cells_in_sorted_order(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = '3:C'

        cell = ws.cell(3, 0)
        cell.value = '4:A'

        cell = ws.cell(0, 0)
        cell.value = '1:A'

        cell = ws.cell(1, 0)
        cell.value = '2:A'

        col = ws.get_column(0, max_row=10, create_cells=True)
        assert [c.text for c in col] == ['1:A', '2:A', None, '4:A'] + [None] * 7

    def test_max_row_in_empty_row_is_negative_one(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.max_row_in_column(0) == -1

    def test_min_row_in_empty_row_is_negative_one(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.min_row_in_column(0) == -1

    def test_max_row_in_column(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = '3:C'

        cell = ws.cell(3, 0)
        cell.value = '4:A'

        cell = ws.cell(0, 0)
        cell.value = '1:A'

        cell = ws.cell(1, 0)
        cell.value = '2:A'

        assert ws.max_row_in_column(0) == 3

    def test_min_row_in_column(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = '3:C'

        cell = ws.cell(3, 0)
        cell.value = '4:A'

        cell = ws.cell(0, 0)
        cell.value = '1:A'

        cell = ws.cell(1, 0)
        cell.value = '2:A'

        assert ws.min_row_in_column(0) == 0

    def test_get_empty_row(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        row = ws.get_row(0)
        row = [r for r in row]

        assert row == []

    def test_get_empty_row_and_create_cells_returns_cells_for_all_columns(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        row = ws.get_row(0, create_cells=True)
        row = [r for r in row]

        assert len(row) == ws.max_allowed_column + 1

    def test_get_row_with_values_returns_only_existing_cells_in_sorted_order(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 0)
        cell.value = '3:C'
        cell = ws.cell(0, 3)
        cell.value = '1:D'
        cell = ws.cell(0, 0)
        cell.value = '1:A'
        cell = ws.cell(0, 1)
        cell.value = '1:B'

        row = ws.get_row(0)
        assert [r.text for r in row] == ['1:A', '1:B', '1:D']

    def test_get_row_within_range_only_returns_cells_within_that_range(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 0)
        cell.value = '3:C'

        cell = ws.cell(0, 3)
        cell.value = '1:D'

        cell = ws.cell(0, 0)
        cell.value = '1:A'

        cell = ws.cell(0, 1)
        cell.value = '1:B'

        row = ws.get_row(0, min_col=1, max_col=2)
        assert [r.text for r in row] == ['1:B']

    def test_get_row_and_create_cells_will_return_all_cells_in_sorted_order(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 2)
        cell.value = '3:C'
        cell = ws.cell(0, 3)
        cell.value = '1:D'
        cell = ws.cell(0, 0)
        cell.value = '1:A'
        cell = ws.cell(0, 1)
        cell.value = '1:B'

        row = ws.get_row(0, max_col=10, create_cells=True)
        assert [r.text for r in row] == ['1:A', '1:B', None, '1:D'] + [None] * 7

    def test_max_column_in_empty_row_is_negative_one(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.max_column_in_row(0) == -1

    def test_min_column_in_empty_row_is_negative_one(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')
        assert ws.min_column_in_row(0) == -1

    def test_max_column_in_row(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 0)
        cell.value = '3:C'

        cell = ws.cell(0, 3)
        cell.value = '1:D'

        cell = ws.cell(0, 0)
        cell.value = '1:A'

        cell = ws.cell(0, 1)
        cell.value = '1:B'

        assert ws.max_column_in_row(0) == 3

    def test_min_column_in_row(self):
        workbook = Workbook()
        ws = workbook.create_sheet('Title')

        cell = ws.cell(2, 0)
        cell.value = '3:C'

        cell = ws.cell(0, 3)
        cell.value = '1:D'

        cell = ws.cell(0, 0)
        cell.value = '1:A'

        cell = ws.cell(0, 1)
        cell.value = '1:B'

        assert ws.min_column_in_row(0) == 0

    def test_removing_worksheet_deletes_it_from_workbook(self):
        workbook = Workbook()
        for i in range(5):
            workbook.create_sheet('Title' + str(i))
        orig_num_worksheets = len(workbook)

        ws2 = workbook.get_sheet_by_name('Title2')

        ws2.remove_from_workbook()
        assert len(workbook) == orig_num_worksheets - 1

    def test_removing_worksheet_keeps_other_worksheets(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(3)]

        ws1 = workbook.get_sheet_by_name('Title1')
        ws1.remove_from_workbook()

        assert workbook.get_sheet_by_name('Title0') == worksheets[0]
        assert workbook.get_sheet_by_name('Title0').title == worksheets[0].title
        assert workbook.get_sheet_by_name('Title2') == worksheets[2]
        assert workbook.get_sheet_by_name('Title2').title == worksheets[2].title

    def test_removing_non_existing_cell_does_not_delete_cells(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        ws.delete_cell(10, 10)
        assert ws.get_cell_collection() == expected_cells

    def test_removing_existing_non_expression_cell(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        expected_cells.remove(ws.cell(0, 1))
        ws.delete_cell(0, 1)
        assert ws.get_cell_collection() == expected_cells

    def test_removing_non_shared_expression_cell(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        expected_cells.remove(ws.cell(3, 1))
        ws.delete_cell(3, 1)
        assert ws.get_cell_collection() == expected_cells

    def test_removing_shared_expression_non_originating_cell(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        expected_cells.remove(ws.cell(4, 1))
        ws.delete_cell(4, 1)
        assert ws.get_cell_collection() == expected_cells

    def test_removing_shared_expression_originating_cell_raises_exception(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.get_sheet_by_name('Expressions')
        expected_cells = ws.get_cell_collection()
        with pytest.raises(UnsupportedOperationException):
            ws.delete_cell(1, 1)
        assert ws.get_cell_collection() == expected_cells

    @pytest.mark.skip('Still need to implement inserting cells/rows/columns')
    def test_inserting_cell_before_another_cell_does_not_cause_the_existing_cell_to_be_recreated(
        self,
    ):
        # This is needed because of the singleton nature of the Cell class
        self.assertTrue(False)

    @pytest.mark.skip('Still need to implement removing cells/rows/columns')
    def test_removing_cell_before_another_cell_does_not_cause_the_existing_cell_to_be_recreated(
        self,
    ):
        # This is needed because of the singleton nature of the Cell class
        self.assertTrue(False)
