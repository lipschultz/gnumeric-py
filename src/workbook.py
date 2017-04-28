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
from xml.etree import ElementTree

from src.sheet import Sheet

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

NEW_SHEET_NAME = '''<?xml version="1.0" encoding="UTF-8"?><gnm:ROOT xmlns:gnm="http://www.gnumeric.org/v10.dtd"><gnm:SheetName gnm:Cols="256" gnm:Rows="65536"></gnm:SheetName></gnm:ROOT>'''
NEW_SHEET = '''<?xml version="1.0" encoding="UTF-8"?><gnm:ROOT xmlns:gnm="http://www.gnumeric.org/v10.dtd">
<gnm:Sheet DisplayFormulas="0" HideZero="0" HideGrid="0" HideColHeader="0" HideRowHeader="0" DisplayOutlines="1" OutlineSymbolsBelow="1" OutlineSymbolsRight="1" Visibility="GNM_SHEET_VISIBILITY_VISIBLE" GridColor="0:0:0">
  <gnm:Name></gnm:Name>
  <gnm:MaxCol>-1</gnm:MaxCol>
  <gnm:MaxRow>-1</gnm:MaxRow>
  <gnm:Zoom>1</gnm:Zoom>
  <gnm:Names>
    <gnm:Name>
      <gnm:name>Print_Area</gnm:name>
      <gnm:value>#REF!</gnm:value>
      <gnm:position>A1</gnm:position>
    </gnm:Name>
    <gnm:Name>
      <gnm:name>Sheet_Title</gnm:name>
      <gnm:value>&quot;Sheet1&quot;</gnm:value>
      <gnm:position>A1</gnm:position>
    </gnm:Name>
  </gnm:Names>
  <gnm:PrintInformation>
    <gnm:Margins>
      <gnm:top Points="120" PrefUnit="mm"/>
      <gnm:bottom Points="120" PrefUnit="mm"/>
      <gnm:left Points="72" PrefUnit="mm"/>
      <gnm:right Points="72" PrefUnit="mm"/>
      <gnm:header Points="72" PrefUnit="mm"/>
      <gnm:footer Points="72" PrefUnit="mm"/>
    </gnm:Margins>
    <gnm:Scale type="percentage" percentage="100"/>
    <gnm:vcenter value="0"/>
    <gnm:hcenter value="0"/>
    <gnm:grid value="0"/>
    <gnm:even_if_only_styles value="0"/>
    <gnm:monochrome value="0"/>
    <gnm:draft value="0"/>
    <gnm:titles value="0"/>
    <gnm:do_not_print value="0"/>
    <gnm:print_range value="GNM_PRINT_ACTIVE_SHEET"/>
    <gnm:order>d_then_r</gnm:order>
    <gnm:orientation>portrait</gnm:orientation>
    <gnm:Header Left="" Middle="&amp;[TAB]" Right=""/>
    <gnm:Footer Left="" Middle="Page &amp;[PAGE]" Right=""/>
    <gnm:paper>na_letter</gnm:paper>
    <gnm:comments placement="GNM_PRINT_COMMENTS_IN_PLACE"/>
    <gnm:errors PrintErrorsAs="GNM_PRINT_ERRORS_AS_DISPLAYED"/>
  </gnm:PrintInformation>
  <gnm:Styles>
    <gnm:StyleRegion startCol="0" startRow="0" endCol="255" endRow="65535">
      <gnm:Style HAlign="GNM_HALIGN_GENERAL" VAlign="GNM_VALIGN_BOTTOM" WrapText="0" ShrinkToFit="0" Rotation="0" Shade="0" Indent="0" Locked="1" Hidden="0" Fore="0:0:0" Back="FFFF:FFFF:FFFF" PatternColor="0:0:0" Format="General">
        <gnm:Font Unit="10" Bold="0" Italic="0" Underline="0" StrikeThrough="0" Script="0">Sans</gnm:Font>
      </gnm:Style>
    </gnm:StyleRegion>
  </gnm:Styles>
  <gnm:Cols DefaultSizePts="48"/>
  <gnm:Rows DefaultSizePts="12.75"/>
  <gnm:Selections CursorCol="0" CursorRow="0">
    <gnm:Selection startCol="0" startRow="0" endCol="0" endRow="0"/>
  </gnm:Selections>
  <gnm:Cells/>
  <gnm:SheetLayout TopLeft="A1"/>
  <gnm:Solver ModelType="0" ProblemType="0" MaxTime="60" MaxIter="1000" NonNeg="1" Discr="0" AutoScale="0" ProgramR="0" SensitivityR="0"/>
</gnm:Sheet>
</gnm:ROOT>
'''


class Workbook:
    def __init__(self):
        self.__root = ElementTree.fromstring(EMPTY_WORKBOOK)
        self._ns = DEFAULT_NAMESPACE

    @property
    def version(self):
        version = self.__root.find('gnm:Version', self._ns)
        return version.get('Full')

    def __sheet_name_elements(self):
        return self.__root.find('gnm:SheetNameIndex', self._ns)

    def __sheet_elements(self):
        return self.__root.find('gnm:Sheets', self._ns)

    def __len__(self):
        return len(self.__sheet_elements())

    def get_sheet_names(self):
        '''
        The list of sheet names, in the order they occur in the workbook.
        '''
        return [s.text for s in self.__sheet_name_elements()]

    @property
    def sheetnames(self):
        return self.get_sheet_names()

    def create_sheet(self, title):
        '''
        Create a new worksheet
        :param title: Title, or name, or worksheet
        :return: The worksheet
        '''
        sheet_name_element = ElementTree.fromstring(NEW_SHEET_NAME).getchildren()[0]
        sheet_element = ElementTree.fromstring(NEW_SHEET).getchildren()[0]

        self.__sheet_name_elements().append(sheet_name_element)
        self.__sheet_elements().append(sheet_element)

        ws = Sheet(sheet_name_element, sheet_element, self)
        ws.title = title
        return ws
