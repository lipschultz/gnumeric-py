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


SHEET_TYPE_REGULAR = None
SHEET_TYPE_OBJECT = 'object'

class Sheet:
    def __init__(self, sheet_name_element, sheet_element, workbook):
        self._sheet_name = sheet_name_element
        self._sheet = sheet_element
        self.__workbook = workbook

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

    def __eq__(self, other):
        return (self.__workbook == other.__workbook and
                self._sheet_name == other._sheet_name and
                self._sheet == other._sheet)
