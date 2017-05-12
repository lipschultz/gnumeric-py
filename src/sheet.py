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
from operator import attrgetter

from lxml import etree

from src import cell
from src.exceptions import UnsupportedOperationException

NEW_CELL = b'''<?xml version="1.0" encoding="UTF-8"?><gnm:ROOT xmlns:gnm="http://www.gnumeric.org/v10.dtd">
<gnm:Cell Row="%(row)a" Col="%(col)a" ValueType="%(value_type)a"/>
</gnm:ROOT>'''

SHEET_TYPE_REGULAR = None
SHEET_TYPE_OBJECT = 'object'


class Sheet:
    __EMPTY_CELL_XPATH_SELECTOR = ('@ValueType="'
                                   + str(cell.VALUE_TYPE_EMPTY)
                                   + '" or (not(@ValueType) and not(@ExprID) and string-length(text())=0)')

    def __init__(self, sheet_name_element, sheet_element, workbook):
        self.__sheet_name = sheet_name_element
        self.__sheet = sheet_element
        self.__workbook = workbook

    def __get_cells(self):
        return self.__sheet.find('gnm:Cells', self.__workbook._ns)

    def __get_empty_cells(self):
        all_cells = self.__get_cells()
        return all_cells.xpath('./gnm:Cell[' + self.__EMPTY_CELL_XPATH_SELECTOR + ']',
                               namespaces=self.__workbook._ns)

    def __get_non_empty_cells(self):
        all_cells = self.__get_cells()
        return all_cells.xpath('./gnm:Cell[not(' + self.__EMPTY_CELL_XPATH_SELECTOR + ')]',
                               namespaces=self.__workbook._ns)

    def __get_cell_element(self, row, col):
        cells = self.__get_cells()
        return cells.find('gnm:Cell[@Row="%d"][@Col="%d"]' % (row, col), self.__workbook._ns)

    def __get_expression_id_cells(self):
        all_cells = self.__get_cells()
        return all_cells.xpath('./gnm:Cell[@ExprID]', namespaces=self.__workbook._ns)

    def __get_styles(self):
        return self.__sheet.xpath('./gnm:Styles', namespaces=self.__workbook._ns)[0]

    def __get_cell_style(self, cell_element):
        row = cell_element.get('Row')
        col = cell_element.get('Col')
        return self.__get_styles().xpath('./gnm:StyleRegion[@startCol<="' + col + '" and "' + col + '"<=@endCol '
                                         + 'and @startRow<="' + row + '" and "' + row + '"<=@endRow]',
                                         namespaces=self.__workbook._ns)[0]

    def __create_and_get_new_cell(self, row_idx, col_idx):
        """
        Creates a new cell, adds it to the worksheet, and returns it.
        """
        new_cell = etree.fromstring(
                NEW_CELL % {b'row': row_idx, b'col': col_idx,
                            b'value_type': cell.VALUE_TYPE_EMPTY}).getchildren()[0]
        cells = self.__get_cells()
        cells.append(new_cell)
        return new_cell

    def __cell_element_to_class(self, element):
        return cell.Cell(element, self.__get_cell_style(element), self, self.__workbook._ns)

    __ce2c = __cell_element_to_class

    @property
    def workbook(self):
        return self.__workbook

    def get_title(self):
        """
        The title, or name, of the worksheet
        """
        return self.__sheet_name.text

    def set_title(self, title):
        sheet_name = self.__sheet.find('gnm:Name', self.__workbook._ns)
        sheet_name.text = self.__sheet_name.text = title

    title = property(get_title, set_title)

    def remove_from_workbook(self):
        """
        Delete this sheet from its workbook.  Note that after this operation, this worksheet will be in an invalid state
        and should not be used.
        """
        self.__sheet_name.getparent().remove(self.__sheet_name)
        self.__sheet.getparent().remove(self.__sheet)

    @property
    def type(self):
        """
        The type of sheet:
         - `SHEET_TYPE_REGULAR` if a regular worksheet
         - `SHEET_TYPE_OBJECT` if an object (e.g. graph) worksheet
        """
        return self.__sheet_name.get('{%s}SheetType' % (self.__workbook._ns['gnm']))

    def __maxmin_rc(self, rc, mm_fn):
        """
        The abstracted method for `max_column`, `max_row`, `min_column`, and `min_row`
        :param rc: `str` indiciating whether this is for `"column"` or `"row"`.
        :param mm_fn: The function for whether to get the max or the min.
        """
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have ' + rc)

        content_cells = self.__get_non_empty_cells()
        return -1 if len(content_cells) == 0 else mm_fn(getattr(self.__ce2c(c), rc) for c in content_cells)

    @property
    def max_column(self):
        """
        The maximum column that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        """
        return self.__maxmin_rc('column', max)

    @property
    def max_row(self):
        """
        The maximum row that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        """
        return self.__maxmin_rc('row', max)

    @property
    def min_column(self):
        """
        The minimum column that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        """
        return self.__maxmin_rc('column', min)

    @property
    def min_row(self):
        """
        The minimum row that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        """
        return self.__maxmin_rc('row', min)

    def __maxmin_rc_in_cr(self, cr, mm_fn, idx):
        """
        The abstracted method for `max_column_in_row` and `max_row_in_column`.
        :param cr: `str` indicating whether to search the `"column"` or `"row"` for the max row/col.
        :param mm_fn: The function for whether to get the max or the min.
        :param idx: `int` indicating which column/row to search through
        """
        rc = 'column' if cr == 'row' else 'row'
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have ' + rc)

        content_cells = self.__get_non_empty_cells()
        content_cells = [self.__ce2c(c) for c in content_cells]
        content_cells = [getattr(c, rc) for c in content_cells if getattr(c, cr) == idx]
        return -1 if len(content_cells) == 0 else mm_fn(content_cells)

    def max_column_in_row(self, row):
        """
        Get the last column in `row` that has a value.  Returns -1 if the row is empty.  Raises
        UnsupportedOperationException when the sheet is a chartsheet.
        """
        return self.__maxmin_rc_in_cr('row', max, row)

    def max_row_in_column(self, column):
        """
        Get the last row in `column` that has a value.  Returns -1 if the column is empty.  Raises
        UnsupportedOperationException when the sheet is a chartsheet.
        """
        return self.__maxmin_rc_in_cr('column', max, column)

    def min_column_in_row(self, row):
        """
        Get the first column in `row` that has a value.  Returns -1 if the row is empty.  Raises
        UnsupportedOperationException when the sheet is a chartsheet.
        """
        return self.__maxmin_rc_in_cr('row', min, row)

    def min_row_in_column(self, column):
        """
        Get the first row in `column` that has a value.  Returns -1 if the column is empty.  Raises
        UnsupportedOperationException when the sheet is a chartsheet.
        """
        return self.__maxmin_rc_in_cr('column', min, column)

    def __max_allowed_rc(self, rc):
        """
        The abstracted method for `max_allowed_column` and `max_allowed_row`.
        :param rc: `str` indiciating whether this is for `"column"` or `"row"`.
        """
        rc = 'Cols' if rc == 'column' else 'Rows'
        key = '{%s}%s' % (self.__workbook._ns['gnm'], rc)
        return int(self.__sheet_name.get(key)) - 1

    @property
    def max_allowed_column(self):
        """
        The maximum column allowed in the worksheet.
        :return: `int`
        """
        return self.__max_allowed_rc('column')

    @property
    def max_allowed_row(self):
        """
        The maximum row allowed in the worksheet.
        :return: `int`
        """
        return self.__max_allowed_rc('row')

    def calculate_dimension(self):
        """
        The minimum bounding rectangle that contains all data in the worksheet

        Raises UnsupportedOperationException when the sheet is a chartsheet.

        :return: A four-tuple of ints: (min_row, min_col, max_row, max_col)
        """
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have rows or columns')
        return (self.min_row, self.min_column, self.max_row, self.max_column)

    def __is_valid_rc(self, rc, idx):
        """
        The abstracted method for `is_valid_column` and `is_valid_row`.
        :param idx: The column or row.
        :param rc: A string indiciating whether this is for a `"column"` or `"row"`.
        :return: bool
        """
        max_allowed = 'max_allowed_' + rc
        return 0 <= idx <= getattr(self, max_allowed)

    def is_valid_column(self, column):
        """
        Returns `True` if column is between `0` and `ws.max_allowed_column`, otherwise returns False
        :return: bool
        """
        return self.__is_valid_rc('column', column)

    def is_valid_row(self, row):
        """
        Returns `True` if row is between `0` and `ws.max_allowed_row`, otherwise returns False
        :return: bool
        """
        return self.__is_valid_rc('row', row)

    def cell(self, row_idx, col_idx, create=True):
        """
        Returns a Cell object for the cell at the specific row and column.

        If the cell does not exist, then an empty cell will be created and returned, unless `create` is `False` (in
        which case, `IndexError` is raised).  Note that the cell will not be added to the worksheet until it is not
        empty (since Gnumeric does not seem to store empty cells).
        """
        if not self.is_valid_row(row_idx):
            raise IndexError('Row (' + str(row_idx) + ') for cell is out of allowed bounds of [0, '
                             + str(self.max_allowed_row) + ']')
        elif not self.is_valid_column(col_idx):
            raise IndexError('Column (' + str(col_idx) + ') for cell is out of allowed bounds of [0, '
                             + str(self.max_allowed_column) + ']')

        cell_found = self.__get_cell_element(row_idx, col_idx)
        if cell_found is None:
            if create:
                cell_found = self.__create_and_get_new_cell(row_idx, col_idx)
            else:
                raise IndexError('No cell exists at position (%d, %d)' % (row_idx, col_idx))

        return self.__ce2c(cell_found)

    def cell_text(self, row_idx, col_idx):
        """
        Returns a the cell's text at the specific row and column.
        """
        return self.cell(row_idx, col_idx, create=False).text

    def __sort_cell_elements(self, cell_elements, row_major):
        """
        Sort the cells according to indices.  If `row_major` is True, then sorting will occur by row first, then within
        each row, columns will be sorted.  If `row_major` is False, then the opposite will happen: first sort by column,
        then by row within each column.

        :returns: A list of Cell objects
        """
        return self.__sort_cells([self.__ce2c(c) for c in cell_elements], row_major)

    def __sort_cells(self, cells, row_major):
        """
        Sort the cells according to indices.  If `row_major` is True, then sorting will occur by row first, then within
        each row, columns will be sorted.  If `row_major` is False, then the opposite will happen: first sort by column,
        then by row within each column.
        """
        key_fn = attrgetter('row', 'column') if row_major else attrgetter('column', 'row')
        return sorted(cells, key=key_fn)

    def get_cell_collection(self, include_empty=False, sort=False):
        """
        Return all cells as a list.

        If `include_empty` is False (default), then only cells with content will be included.  If `include_empty` is
        True, then empty cells that have been created will be included.

        Use `sort` to specify whether the cells should be sorted.  If `False` (default), then no sorting will take
        place.  If `sort` is `"row"`, then sorting will occur by row first, then by column within each row.  If `sort`
        is `"column"`, then the opposite will happen: first sort by column, then by row within each column.
        """
        if include_empty:
            cells = self.__get_cells()
        else:
            cells = self.__get_non_empty_cells()

        cells = [self.__ce2c(c) for c in cells]
        if sort:
            return self.__sort_cells(cells, sort == 'row')
        else:
            return cells

    def __get_rc(self, rc, idx, min_cr, max_cr, create_cells):
        """
        The abstracted method for `get_column` and `get_row`.
        :param rc: `str` indiciating whether this is for `"column"` or `"row"`.
        """
        cr = 'row' if rc == 'column' else 'column'
        rc_attr = rc[:3].title()
        cr_attr = cr[:3].title()
        if not self.__is_valid_rc(rc, idx):
            raise IndexError(rc.title() + ' (' + str(idx) + ') is out of allowed bounds of [0, '
                             + str(self.__max_allowed_rc(rc)) + ']')

        if not self.__is_valid_rc(cr, min_cr):
            raise IndexError('Min ' + cr + ' (' + str(min_cr) + ') is out of allowed bounds of [0, '
                             + str(self.__max_allowed_rc(cr)) + ']')

        if max_cr is not None and not self.__is_valid_rc(cr, max_cr):
            raise IndexError('Max ' + cr + ' (' + str(max_cr) + ') is out of allowed bounds of [0, '
                             + str(self.__max_allowed_rc(cr)) + ']')
        elif max_cr is None:
            max_cr = self.__maxmin_rc_in_cr(rc, max, idx)
            if max_cr == -1:
                max_cr = self.__max_allowed_rc(cr)

        existing_cells = self.__get_cells().xpath(
                './gnm:Cell[@' + rc_attr + '="' + str(idx)
                + '" and "' + str(min_cr) + '"<=@' + cr_attr + ' and @' + cr_attr + '<="' + str(max_cr) + '"]',
                namespaces=self.__workbook._ns)
        cells = self.__sort_cell_elements(existing_cells, False)

        if create_cells:
            cell_map = dict([(getattr(c, cr), c) for c in cells])
            return (cell_map[i] if i in cell_map else self.cell(i, idx) for i in range(min_cr, max_cr + 1))
        else:
            return (c for c in cells)

    def get_column(self, column, min_row=0, max_row=None, create_cells=False):
        """
        Get the cells in the specified column (index starting at 0).

        Use `min_row` and `max_row` to specify the range of cells within the column.  `min_row` defaults to 0.
        `max_row` defaults to `None`, meaning go to the last cell with a value in the column, but if the column is
        empty, then go to the last allowed row.  These bounds are inclusive.

        If `create_cells` is `True`, then any cells in the column that don't exist, will be created.  If `False` (the
        default), then only already-existing cells will be returned.  Note that existing cells that are empty will be
        returned if `create_cells` is `False`.

        Raises `IndexError` if `column` is outside of the valid range for columns, or if `min_row` or `max_row` is
        outside the valid range for rows.
        :return: generator
        """
        return self.__get_rc('column', column, min_row, max_row, create_cells)

    def get_row(self, row, min_col=0, max_col=None, create_cells=False):
        """
        Get the cells in the specified row (index starting at 0).

        Use `min_col` and `max_col` to specify the range of cells within the row.  `min_col` defaults to 0.  `max_col`
        defaults to `None`, meaning go to the last cell with a value in the row, but if the row is empty, then go to the
        last allowed column.  These bounds are inclusive.

        If `create_cells` is `True`, then any cells in the row that don't exist, will be created.  If `False` (the
        default), then only already-existing cells will be returned.  Note that existing cells that are empty will be
        returned if `create_cells` is `False`.

        Raises `IndexError` if `row` is outside of the valid range for rows, or if `min_col` or `max_col` is outside
        the valid range for columns.
        :return: generator
        """
        return self.__get_rc('row', row, min_col, max_col, create_cells)

    def get_expression_map(self):
        """
        In each worksheet, Gnumeric stores an expression/formula once (in the cell it's first used), then references it
        by an id in all other cells that use the expression.  This method will return a dict of
        expression ids -> ((cell_row, cell_col), expression).

        Note that this might not return all expressions in the worksheet.  If an expression is only used once, then it
        may not have an id and thus will not be returned by this method.
        """
        cells = self.__get_expression_id_cells()
        return dict([(c.get('ExprID'), ((int(c.get('Row')), int(c.get('Col'))), c.text)) for c in cells
                     if c.text is not None])

    def get_all_cells_with_expression(self, id, sort=False):
        """
        Returns a list of all cells referencing/using the expression with the provided id.

        Use `sort` to specify whether the cells should be sorted.  If `False` (default), then no sorting will take
        place.  If `sort` is `"row"`, then sorting will occur by row first, then by column within each row.  If `sort`
        is `"column"`, then the opposite will happen: first sort by column, then by row within each column.
        """
        cells = self.__get_expression_id_cells()
        cells = [self.__ce2c(c) for c in cells if c.get('ExprID') == id]
        if sort:
            return self.__sort_cells(cells, sort == 'row')
        else:
            return cells

    def delete_cell(self, row, col):
        """
        Deletes the cell at the specified row and column.  If the cell doesn't exist, then nothing will happen.  If the
        cell is the originating cell for an expression, then an exception is thrown (deleting these cells is not yet
        supported).
        """
        cell = self.__get_cell_element(row, col)
        if cell is None:
            return
        elif cell.get('ExprID') is not None and cell.text is not None:
            raise UnsupportedOperationException("Can't delete originating cell for an expression")

        all_cells = self.__get_cells()
        all_cells.remove(cell)

    def _clean_data(self):
        """
        Performs housekeeping on the data.  Only necessary when contents are being written to file.  Should not be
        called directly -- the workbook will call this automatically when writing to file.
        """

        # Delete empty cells
        all_cells = self.__get_cells()
        empty_cells = self.__get_empty_cells()
        for empty_cell in empty_cells:
            all_cells.remove(empty_cell)

        # Update max col and row
        self.__sheet.find('gnm:MaxCol', self.__workbook._ns).text = str(self.max_column)
        self.__sheet.find('gnm:MaxRow', self.__workbook._ns).text = str(self.max_row)

    def __eq__(self, other):
        return (self.__workbook == other.__workbook and
                self.__sheet_name == other.__sheet_name and
                self.__sheet == other.__sheet)

    def __str__(self):
        return self.title
