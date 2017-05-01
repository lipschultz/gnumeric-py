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

from src.exceptions import UnrecognizedCellTypeException

VALUE_TYPE_EXPR = -10
VALUE_TYPE_EMPTY = 10
VALUE_TYPE_BOOLEAN = 20
VALUE_TYPE_INTEGER = 30
VALUE_TYPE_FLOAT = 40
VALUE_TYPE_ERROR = 50
VALUE_TYPE_STRING = 60
VALUE_TYPE_CELLRANGE = 70
VALUE_TYPE_ARRAY = 80


class Cell:
    def __init__(self, cell_element, worksheet):
        self.__cell = cell_element
        self.__worksheet = worksheet

    @property
    def column(self):
        '''
        The column this cell belongs to (0-indexed).
        '''
        return int(self.__cell.get('Col'))

    @property
    def row(self):
        '''
        The row this cell belongs to (0-indexed).
        '''
        return int(self.__cell.get('Row'))

    @property
    def text(self):
        '''
        Returns the raw value stored in the cell.  The text will be `None` if the cell is empty.
        :return: str
        '''
        return self.__cell.text

    @property
    def type(self):
        '''
        Returns the type of value stored in the cell:
         - VALUE_TYPE_EXPR = -10
         - VALUE_TYPE_EMPTY = 10
         - VALUE_TYPE_BOOLEAN = 20
         - VALUE_TYPE_INTEGER = 30
         - VALUE_TYPE_FLOAT = 40
         - VALUE_TYPE_ERROR = 50
         - VALUE_TYPE_STRING = 60
         - VALUE_TYPE_CELLRANGE = 70
         - VALUE_TYPE_ARRAY = 80
        '''
        value_type = self.__cell.get('ValueType')
        if value_type is not None:
            return int(value_type)
        elif self.__cell.get('ExprID') is not None or self.text.startswith('='):
            return VALUE_TYPE_EXPR
        else:
            raise UnrecognizedCellTypeException('Cell is: "' + str(etree.tostring(self.__cell)) + '"')

    def get_value(self):
        '''
        Gets the value stored in the cell, converted into the appropriate Python datatype when possible.
        '''
        value = self.text
        if self.type == VALUE_TYPE_BOOLEAN:
            return bool(value)
        elif self.type == VALUE_TYPE_INTEGER:
            return int(value)
        elif self.type == VALUE_TYPE_FLOAT:
            return float(value)
        else:
            return value