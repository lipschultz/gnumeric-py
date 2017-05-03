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
        self._sheet_name = sheet_name_element
        self._sheet = sheet_element
        self.__workbook = workbook

    def __get_cells(self):
        return self._sheet.find('gnm:Cells', self.__workbook._ns)

    def __get_empty_cells(self):
        all_cells = self.__get_cells()
        return all_cells.xpath('./gnm:Cell[' + self.__EMPTY_CELL_XPATH_SELECTOR + ']',
                               namespaces=self.__workbook._ns)

    def __get_non_empty_cells(self):
        all_cells = self.__get_cells()
        return all_cells.xpath('./gnm:Cell[not(' + self.__EMPTY_CELL_XPATH_SELECTOR + ')]',
                               namespaces=self.__workbook._ns)

    def __get_expression_id_cells(self):
        all_cells = self.__get_cells()
        return all_cells.xpath('./gnm:Cell[@ExprID]', namespaces=self.__workbook._ns)

    def __create_and_get_new_cell(self, row_idx, col_idx):
        '''
        Creates a new cell, adds it to the worksheet, and returns it.
        '''
        new_cell = etree.fromstring(
                NEW_CELL % {b'row': row_idx, b'col': col_idx,
                            b'value_type': cell.VALUE_TYPE_EMPTY}).getchildren()[0]
        cells = self.__get_cells()
        cells.append(new_cell)
        return new_cell

    @property
    def workbook(self):
        return self.__workbook

    def get_title(self):
        '''
        The title, or name, of the worksheet
        '''
        return self._sheet_name.text

    def set_title(self, title):
        sheet_name = self._sheet.find('gnm:Name', self.__workbook._ns)
        sheet_name.text = self._sheet_name.text = title

    title = property(get_title, set_title)

    @property
    def type(self):
        '''
        The type of sheet:
         - `SHEET_TYPE_REGULAR` if a regular worksheet
         - `SHEET_TYPE_OBJECT` if an object (e.g. graph) worksheet
        '''
        return self._sheet_name.get('{%s}SheetType' % (self.__workbook._ns['gnm']))

    @property
    def max_column(self):
        '''
        The maximum column that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        '''
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have max column')

        content_cells = self.__get_non_empty_cells()
        return -1 if len(content_cells) == 0 else max(cell.Cell(c, self).column for c in content_cells)

    @property
    def max_allowed_column(self):
        """
        The maximum column allowed in the worksheet.
        :return: `int`
        """
        key = '{%s}Cols' % (self.__workbook._ns['gnm'])
        return int(self._sheet_name.get(key))-1

    @property
    def max_row(self):
        '''
        The maximum row that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        '''
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have max row')

        content_cells = self.__get_non_empty_cells()
        return -1 if len(content_cells) == 0 else max(cell.Cell(c, self).row for c in content_cells)

    @property
    def max_allowed_row(self):
        """
        The maximum row allowed in the worksheet.
        :return: `int`
        """
        key = '{%s}Rows' % (self.__workbook._ns['gnm'])
        return int(self._sheet_name.get(key))-1

    @property
    def min_column(self):
        '''
        The minimum column that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        '''
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have min column')
        content_cells = self.__get_non_empty_cells()
        return -1 if len(content_cells) == 0 else min([cell.Cell(c, self).column for c in content_cells])

    @property
    def min_row(self):
        '''
        The minimum row that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        '''
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have min row')
        content_cells = self.__get_non_empty_cells()
        return -1 if len(content_cells) == 0 else min([cell.Cell(c, self).row for c in content_cells])

    def calculate_dimension(self):
        '''
        The minimum bounding rectangle that contains all data in the worksheet

        Raises UnsupportedOperationException when the sheet is a chartsheet.

        :return: A four-tuple of ints: (min_row, min_col, max_row, max_col)
        '''
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have rows or columns')
        return (self.min_row, self.min_column, self.max_row, self.max_column)

    def cell(self, row_idx, col_idx, create=True):
        '''
        Returns a Cell object for the cell at the specific row and column.

        If the cell does not exist, then an empty cell will be created and returned, unless `create` is `False` (in
        which case, `IndexError` is raised).  Note that the cell will not be added to the worksheet until it is not
        empty (since Gnumeric does not seem to store empty cells).
        '''
        cells = self.__get_cells()
        cell_found = cells.find('gnm:Cell[@Row="%d"][@Col="%d"]' % (row_idx, col_idx), self.__workbook._ns)
        if cell_found is None:
            if create:
                cell_found = self.__create_and_get_new_cell(row_idx, col_idx)
            else:
                raise IndexError('No cell exists at position (%d, %d)' % (row_idx, col_idx))

        return cell.Cell(cell_found, self)

    def cell_text(self, row_idx, col_idx):
        '''
        Returns a the cell's text at the specific row and column.
        '''
        return self.cell(row_idx, col_idx, create=False).text

    def get_cell_collection(self, include_empty=False):
        """
        Return all cells as a list.  The cells may be unordered.

        If `include_empty` is False (default), then only cells with content will be included.  If `include_empty` is
        True, then empty cells that have been created will be included.
        """
        if include_empty:
            cells = self.__get_cells()
        else:
            cells = self.__get_non_empty_cells()
        return [cell.Cell(c, self) for c in cells]

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
        self._sheet.find('gnm:MaxCol', self.__workbook._ns).text = str(self.max_column)
        self._sheet.find('gnm:MaxRow', self.__workbook._ns).text = str(self.max_row)

    def __eq__(self, other):
        return (self.__workbook == other.__workbook and
                self._sheet_name == other._sheet_name and
                self._sheet == other._sheet)

    def __str__(self):
        return self.title
