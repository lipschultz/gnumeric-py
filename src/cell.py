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
        :return: str or `None`
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

    def __set_type(self, value_type):
        if value_type == VALUE_TYPE_EXPR:
            if 'ValueType' in self.__cell.keys():
                self.__cell.attrib.pop('ValueType')
                # TODO: How should expression IDs be handled?
        else:
            self.__cell.set('ValueType', str(value_type))

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

    def set_value(self, value, value_type='infer'):
        '''
        Sets the value stored in the cell.

        If `value_type` is:
        - one of the `VALUE_TYPE_` constants, then that value type will be used
        - ``keep'`, then the cell will keep its type
        - `'infer'`, then it tries to guess the type based on the value being set:
            - if `value` is of type `bool`, then `VALUE_TYPE_BOOLEAN`
            - if `value` is of type `int`, then `VALUE_TYPE_INTEGER`
            - if `value` is of type `float`, then `VALUE_TYPE_FLOAT`
            - if `value` is an empty string or `None`, then `VALUE_TYPE_EMPTY`
            - if `value` is a string and starts with `=`, then `VALUE_TYPE_EXPR`
            - if `value` is anything else, then `VALUE_TYPE_STRING`

        The default `value_type` is `'infer'`.

        Warning: This method does no type checking, so it is possible to save a string into a cell whose type is
        `VALUE_TYPE_INTEGER`.  This could result in problems when opening the workbook in Gnumeric.
        '''
        val_types = {bool: VALUE_TYPE_BOOLEAN, int: VALUE_TYPE_INTEGER, float: VALUE_TYPE_FLOAT}
        if value_type == 'infer':
            if type(value) in val_types:
                value_type = val_types[type(value)]
            elif value in ('', None):
                value_type = VALUE_TYPE_EMPTY
            elif value[0] == '=':
                value_type = VALUE_TYPE_EXPR
            elif self.type in (
                    VALUE_TYPE_EMPTY, VALUE_TYPE_BOOLEAN, VALUE_TYPE_INTEGER, VALUE_TYPE_FLOAT, VALUE_TYPE_ERROR):
                value_type = VALUE_TYPE_STRING
            else:
                value_type = self.type
        elif value_type == 'keep':
            value_type = self.type

        if value_type == VALUE_TYPE_BOOLEAN:
            self.__cell.text = str(bool(value)).upper()
        elif value_type == VALUE_TYPE_EMPTY:
            self.__cell.text = None
        else:
            self.__cell.text = str(value)

        self.__set_type(value_type)

    value = property(get_value, set_value, doc='Get or set the value in the cell, converted into the correct type.')

    def __str__(self):
        return repr(self.value)

    def __repr__(self):
        return 'Cell[%s, (%d, %d), ws="%s"]' % (str(self), self.row, self.column, self.__worksheet.title)

    def __eq__(self, other):
        return (isinstance(other, Cell)
                and self.__worksheet == other.__worksheet and self.row == other.row and self.column == other.column)
