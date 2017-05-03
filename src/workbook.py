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
from datetime import datetime
from lxml import etree
import dateutil.parser
import gzip

from src import sheet
from src.exceptions import DuplicateTitleException, WrongWorkbookException
from src.sheet import Sheet

EMPTY_WORKBOOK = b'''<?xml version="1.0" encoding="UTF-8"?>
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
ALL_NAMESPACES = {'gnm': "http://www.gnumeric.org/v10.dtd",
                  'xsi': "http://www.w3.org/2001/XMLSchema-instance",
                  'office': "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
                  'xlink': "http://www.w3.org/1999/xlink",
                  'dc': "http://purl.org/dc/elements/1.1/",
                  'meta': "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
                  'ooo': "http://openoffice.org/2004/office"}

NEW_SHEET_NAME = b'''<?xml version="1.0" encoding="UTF-8"?><gnm:ROOT xmlns:gnm="http://www.gnumeric.org/v10.dtd"><gnm:SheetName gnm:Cols="256" gnm:Rows="65536"></gnm:SheetName></gnm:ROOT>'''
NEW_SHEET = b'''<?xml version="1.0" encoding="UTF-8"?><gnm:ROOT xmlns:gnm="http://www.gnumeric.org/v10.dtd">
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
    def __init__(self, workbook_root_element=None):
        self._ns = ALL_NAMESPACES
        if workbook_root_element is None:
            self.__root = etree.fromstring(EMPTY_WORKBOOK)
            self.creation_date = datetime.now()
        else:
            self.__root = workbook_root_element

    def __creation_date_element(self):
        return self.__root.find('office:document-meta/office:meta/meta:creation-date', self._ns)

    def __sheet_name_elements(self):
        return self.__root.find('gnm:SheetNameIndex', self._ns)

    def __sheet_elements(self):
        return self.__root.find('gnm:Sheets', self._ns)

    @property
    def version(self):
        version = self.__root.find('gnm:Version', self._ns)
        return version.get('Full')

    def get_creation_date(self):
        '''
        Date the workbook was created
        '''
        creation_element = self.__creation_date_element()
        return dateutil.parser.parse(creation_element.text)

    def set_creation_date(self, creation_datetime):
        '''
        :param creation_datetime: A datetime object representing when this workbook was created
        '''
        creation_element = self.__creation_date_element()
        creation_element.text = creation_datetime.isoformat()

    creation_date = property(get_creation_date, set_creation_date)

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

    def create_sheet(self, title, index=-1):
        '''
        Create a new worksheet
        :param title: Title, or name, or worksheet
        :param index: Where to insert the new sheet within the list of sheets. Default is `-1` (to append).
        :raises DuplicateTitleException: When a sheet with the same title already exists in the workbook
        :return: The worksheet
        '''
        if title in self.sheetnames:
            raise DuplicateTitleException('A sheet titled "%s" already exists' % (title))

        sheet_name_element = etree.fromstring(NEW_SHEET_NAME).getchildren()[0]
        sheet_element = etree.fromstring(NEW_SHEET).getchildren()[0]

        if index < 0:
            index = len(self) + index + 1
        self.__sheet_name_elements().insert(index, sheet_name_element)
        self.__sheet_elements().insert(index, sheet_element)

        ws = Sheet(sheet_name_element, sheet_element, self)
        ws.title = title
        return ws

    def get_sheet_by_index(self, index):
        '''
        Get the sheet at the specified index.
        :raises IndexError: When index is out of bounds
        '''
        return Sheet(self.__sheet_name_elements()[index],
                     self.__sheet_elements()[index],
                     self
                     )

    def get_sheet_by_name(self, name):
        '''
        Get the sheet with the specified title/name
        :raises KeyError: When not worksheet with that name exists
        '''
        names = self.sheetnames
        try:
            idx = names.index(name)
            return self.get_sheet_by_index(idx)
        except ValueError:
            raise KeyError('No sheet named "%s" exists' % (name))

    def __getitem__(self, key):
        '''
        Get a worksheet by name or by index.  If `key` is a string, then it is treated as a worksheet's name.  If it is
         an int, then it's treated as an index.

         Raises a `TypeError` if key is not and `int` or a `str`.  Can also raise the exceptions that
         `get_sheet_by_index` and `get_sheet_by_name` raise.
        '''
        if isinstance(key, str):
            return self.get_sheet_by_name(key)
        elif isinstance(key, int):
            return self.get_sheet_by_index(key)
        else:
            raise TypeError("Unexpected type (%s) for key: %s" % (type(key), str(key)))

    def get_index(self, ws):
        '''
        Given a worksheet, find its index in the workbook.
        '''
        index = self.sheetnames.index(ws.title)
        if ws != self.get_sheet_by_index(index):
            raise WrongWorkbookException("The worksheet does not belong to this workbook.")
        return index

    def index(self, ws):
        return self.get_index(ws)

    def remove_sheet_by_name(self, name):
        '''
        Remove the worksheet named `name` from the workbook.  See `get_sheet_by_name` and `remove_sheet` for exceptions
        thrown.
        '''
        self.remove_sheet(self.get_sheet_by_name(name))

    def remove_sheet_by_index(self, index):
        '''
        Remove the worksheet at `index` from the workbook.  See `get_sheet_by_index` and `remove_sheet` for exceptions
        thrown.
        '''
        self.remove_sheet(self.get_sheet_by_index(index))

    def remove_sheet(self, ws):
        '''
        Remove the worksheet from the workbook.
        '''
        try:
            self.__sheet_name_elements().remove(ws._sheet_name)
            self.__sheet_elements().remove(ws._sheet)
        except ValueError:
            raise ValueError("Worksheet is not part of workbook")

    def remove(self, ws):
        '''
        Remove the specified worksheet
        :param ws: Can be the worksheet to remove, the index of the worksheet, or the name of the worksheet.
        '''
        if isinstance(ws, int):
            self.remove_sheet_by_index(ws)
        elif isinstance(ws, str):
            self.remove_sheet_by_name(ws)
        else:
            self.remove_sheet(ws)

    def __delitem__(self, key):
        '''
        Remove the specified worksheet
        :param key: Can be the worksheet to remove, the index of the worksheet, or the name of the worksheet.
        '''
        self.remove(key)

    @property
    def sheets(self):
        '''
        Get a list of all sheets in the workbook.
        '''
        return [self.get_sheet_by_index(i) for i in range(len(self))]

    @property
    def chartsheets(self):
        '''
        Get list of only the chart sheets in the workbook.
        '''
        return [s for s in self.sheets if s.type == sheet.SHEET_TYPE_OBJECT]

    @property
    def worksheets(self):
        '''
        Get list of only the non-chart sheets in the workbook.
        '''
        return [s for s in self.sheets if s.type == sheet.SHEET_TYPE_REGULAR]

    def save(self, filepath, compress=9):
        '''
        Save the workbook to `filepath`.
        :param compress: The level of compression to apply to the file.  A value between 0 (no compression, but still
            write it as a gzip-compressed Gnumeric file) and 9 (slowest but most compressed; default).  A `False` value
            will write a uncompressed Gnumeric file.
        '''
        [s._prepare_for_saving() for s in self.sheets]
        xml = etree.tostring(self.__root)
        if compress:
            with gzip.open(filepath, mode='wb') as fout:
                fout.write(xml)
        else:
            with open(filepath, mode='wb') as fout:
                fout.write(xml)

    @classmethod
    def load_workbook(clas, filepath):
        '''
        Open the given filepath and return the workbook.
        '''
        if filepath.endswith('.xml'):
            open_method = open
        else:
            open_method = gzip.open

        with open_method(filepath, mode='rb') as fin:
            contents = fin.read()

        root = etree.fromstring(contents)
        return Workbook(root)
