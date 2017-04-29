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
