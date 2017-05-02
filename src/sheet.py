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
    def __init__(self, sheet_name_element, sheet_element, workbook):
        self._sheet_name = sheet_name_element
        self._sheet = sheet_element
        self.__workbook = workbook

    def __get_cells(self):
        return self._sheet.find('gnm:Cells', self.__workbook._ns)

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
        return int(self._sheet.find('gnm:MaxCol', self.__workbook._ns).text)

    @property
    def max_row(self):
        '''
        The maximum row that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        '''
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have max row')
        return int(self._sheet.find('gnm:MaxRow', self.__workbook._ns).text)

    @property
    def min_column(self):
        '''
        The minimum column that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        '''
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have min column')
        cells = self.__get_cells()
        return min([cell.Cell(c, self).column for c in cells])

    @property
    def min_row(self):
        '''
        The minimum row that still holds data.  Raises UnsupportedOperationException when the sheet is a chartsheet.
        :return: `int`
        '''
        if self.type == SHEET_TYPE_OBJECT:
            raise UnsupportedOperationException('Chartsheet does not have min row')
        cells = self.__get_cells()
        return min([cell.Cell(c, self).row for c in cells])

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

    def __eq__(self, other):
        return (self.__workbook == other.__workbook and
                self._sheet_name == other._sheet_name and
                self._sheet == other._sheet)
