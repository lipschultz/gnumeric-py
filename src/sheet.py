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
import xml.etree.ElementTree


class Sheet:
    def __init__(self, sheet_name_element, sheet_element, workbook):
        self.__sheet_name = sheet_name_element
        self.__sheet = sheet_element
        self.__workbook = workbook

    def get_title(self):
        '''
        The title, or name, of the worksheet
        '''
        return self.__sheet_name.text

    def set_title(self, title):
        sheet_name = self.__sheet.find('gnm:Name', self.__workbook._ns)
        sheet_name.text = self.__sheet_name.text = title

    title = property(get_title, set_title)

    def __eq__(self, other):
        return self.__sheet_name == other.__sheet_name and self.__sheet == other.__sheet