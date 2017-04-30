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
    def __init__(self, cell_element):
        self.__cell = cell_element

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
