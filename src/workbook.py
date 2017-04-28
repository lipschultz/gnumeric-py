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

EMPTY_WORKBOOK = '''<?xml version="1.0" encoding="UTF-8"?>
<gnm:Workbook xmlns:gnm="http://www.gnumeric.org/v10.dtd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.gnumeric.org/v9.xsd">
  <gnm:Version Epoch="1" Major="12" Minor="28" Full="1.12.28"/>
  <gnm:Attributes>
    <gnm:Attribute>
      <gnm:name>WorkbookView::show_horizontal_scrollbar</gnm:name>
      <gnm:value>TRUE</gnm:value>
    </gnm:Attribute>
    <gnm:Attribute>
      <gnm:name>WorkbookView::show_vertical_scrollbar</gnm:name>
      <gnm:value>TRUE</gnm:value>
    </gnm:Attribute>
    <gnm:Attribute>
      <gnm:name>WorkbookView::show_notebook_tabs</gnm:name>
      <gnm:value>TRUE</gnm:value>
    </gnm:Attribute>
    <gnm:Attribute>
      <gnm:name>WorkbookView::do_auto_completion</gnm:name>
      <gnm:value>TRUE</gnm:value>
    </gnm:Attribute>
    <gnm:Attribute>
      <gnm:name>WorkbookView::is_protected</gnm:name>
      <gnm:value>FALSE</gnm:value>
    </gnm:Attribute>
  </gnm:Attributes>
  <office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.2">
    <office:meta>
      <meta:creation-date>2017-04-27T22:47:06Z</meta:creation-date>
    </office:meta>
  </office:document-meta>
  <gnm:Calculation ManualRecalc="0" EnableIteration="1" MaxIterations="100" IterationTolerance="0.001" FloatRadix="2" FloatDigits="53"/>
  <gnm:SheetNameIndex>
  </gnm:SheetNameIndex>
  <gnm:Geometry Width="1024" Height="383"/>
  <gnm:Sheets>
  </gnm:Sheets>
  <gnm:UIData SelectedTab="0"/>
</gnm:Workbook>
'''
DEFAULT_NAMESPACE = {'gnm': 'http://www.gnumeric.org/v10.dtd'}


class Workbook:
    def __init__(self):
        self.__root = xml.etree.ElementTree.fromstring(EMPTY_WORKBOOK)
        self.__ns = DEFAULT_NAMESPACE

    @property
    def version(self):
        version = self.__root.find('gnm:Version', self.__ns)
        return version.get('Full')

    def __sheet_name_elements(self):
        return self.__root.find('gnm:SheetNameIndex', self.__ns)

    def __sheet_elements(self):
        return self.__root.find('gnm:Sheets', self.__ns)

    def __len__(self):
        return len(self.__sheet_elements())

    def get_sheet_names(self):
        '''
        The list of sheet names, in the order they occur in the workbook.
        '''
        return [s.text for s in self.__sheet_name_elements()]