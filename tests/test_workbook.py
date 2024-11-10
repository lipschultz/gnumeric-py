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

import gzip
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from dateutil.tz import tzutc

from gnumeric.exceptions import DuplicateTitleException, WrongWorkbookException
from gnumeric.workbook import Workbook

TEST_GNUMERIC_FILE_PATH = 'samples/test.gnumeric'
SHEET_NAMES = (
    'Sheet1',
    'BoundingRegion',
    'CellTypes',
    'Strings',
    'Errors',
    'Expressions',
    'Dates',
    'Mine & Yours Sheet[s]!',
)
GRAPH_NAMES = ('Graph1',)
ALL_NAMES = SHEET_NAMES + GRAPH_NAMES


TEST_SHEET_NAME_FILE_PATH = 'samples/sheet-names.xml'
SHEET_NAME_SHEET_NAMES = (
    'Sheet1',
    'Sheet2',
    'Sheet3',
    'Mine & Yours Sheet[s]!',
    'Graph1',
)


class TestWorkbook:
    def test_equality_of_same_workbook(self):
        workbook = Workbook()
        assert workbook == workbook

    def test_different_workbooks_are_not_equal(self):
        workbook = Workbook()
        wb2 = Workbook()
        assert workbook != wb2

    def test_creating_empty_workbook_has_zero_sheets(self):
        workbook = Workbook()
        assert len(workbook) == 0
        assert len(workbook.get_sheet_names()) == 0

    def test_creating_workbook_has_version_1_12_28(self):
        workbook = Workbook()
        assert workbook.version == '1.12.28'

    def test_creating_sheet_in_empty_book_adds_sheet_to_book(self):
        workbook = Workbook()
        title = 'Title'
        workbook.create_sheet(title)
        assert len(workbook) == 1
        assert workbook.get_sheet_names() == [title]

    def test_creating_sheet_with_name_used_by_another_sheet_raises_exception(self):
        workbook = Workbook()
        title = 'Title'
        workbook.create_sheet(title)
        with pytest.raises(DuplicateTitleException):
            workbook.create_sheet(title)

    def test_creating_sheet_appends_to_list_of_sheets(self):
        workbook = Workbook()
        titles = ['Title1', 'Title2']
        for title in titles:
            workbook.create_sheet(title)
        assert len(workbook) == 2
        assert workbook.get_sheet_names() == titles

    def test_prepending_new_sheet(self):
        workbook = Workbook()
        titles = ['Title1', 'Title2']
        for title in titles:
            workbook.create_sheet(title)
        assert workbook.get_sheet_names() == titles

    def test_inserting_new_sheet(self):
        workbook = Workbook()
        titles = ['Title1', 'Title3', 'Title2']
        workbook.create_sheet(titles[0])
        workbook.create_sheet(titles[2])
        workbook.create_sheet(titles[1], index=1)
        assert workbook.get_sheet_names() == titles

    def test_creation_date_on_new_workbook(self):
        before = datetime.now()
        workbook = Workbook()
        after = datetime.now()
        assert (before < workbook.creation_date < after) is True

    def test_setting_creation_date(self):
        workbook = Workbook()
        creation_date = datetime.now()
        workbook.creation_date = creation_date
        assert workbook.creation_date == creation_date

    def test_getting_sheet_by_index_with_positive_index(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = 3
        ws = workbook.get_sheet_by_index(index)
        assert ws == worksheets[index]

    def test_getting_sheet_by_index_with_negative_index(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = -2
        ws = workbook.get_sheet_by_index(index)
        assert ws == worksheets[index]

    def test_getting_sheet_out_of_bounds_raises_exception(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        with pytest.raises(IndexError):
            workbook.get_sheet_by_index(len(worksheets) + 10)

    def test_getting_sheet_by_name(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws = workbook.get_sheet_by_name('Title' + str(index))
        assert ws == worksheets[index]

    def test_getting_sheet_with_nonexistent_name_raises_exception(self):
        workbook = Workbook()
        for i in range(5):
            workbook.create_sheet(f'Title{i}')
        with pytest.raises(KeyError):
            workbook.get_sheet_by_name('NotASheet')

    def test_getting_sheet_with_getitem_using_positive_index(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = 3
        ws = workbook[index]
        assert ws == worksheets[index]

    def test_getting_sheet_with_getitem_using_negative_index(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = -2
        ws = workbook[index]
        assert ws == worksheets[index]

    def test_getting_sheet_with_getitem_out_of_bounds_raises_exception(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        with pytest.raises(IndexError):
            _ = workbook[len(worksheets) + 10]

    def test_getting_sheet_with_getitem_using_bad_key_type_raises_exception(self):
        workbook = Workbook()
        for i in range(5):
            workbook.create_sheet(f'Title{i}')
        with pytest.raises(TypeError):
            _ = workbook[2.0]

    def test_getting_sheet_with_getitem_by_name(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws = workbook['Title' + str(index)]
        assert ws == worksheets[index]

    def test_getting_sheet_with_getitem_using_nonexistent_name_raises_exception(self):
        workbook = Workbook()
        for i in range(5):
            workbook.create_sheet(f'Title{i}')
        with pytest.raises(KeyError):
            _ = workbook['NotASheet']

    def test_getting_index_of_worksheet(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        assert workbook.get_index(ws2) == index

    def test_getting_index_of_worksheet_from_different_workbook_but_with_name_raises_exception(
        self,
    ):
        workbook = Workbook()
        wb2 = Workbook()
        title = 'Title'
        workbook.create_sheet(title)
        ws2 = wb2.create_sheet(title)
        with pytest.raises(WrongWorkbookException):
            workbook.get_index(ws2)

    def test_deleting_worksheet(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        title = ws2.title
        workbook.remove_sheet(ws2)
        assert len(workbook) == len(worksheets) - 1
        assert len(workbook.sheetnames) == len(worksheets) - 1
        assert title not in workbook.sheetnames
        assert all(ws2 != workbook[i] for i in range(len(workbook)))

    def test_deleting_worksheet_twice_raises_exception(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        workbook.remove_sheet(ws2)
        with pytest.raises(ValueError):
            workbook.remove_sheet(ws2)

    def test_deleting_worksheet_from_another_workbook_raises_exception(self):
        workbook = Workbook()
        wb2 = Workbook()
        for i in range(5):
            workbook.create_sheet(f'Title{i}')
        ws2 = wb2.create_sheet('Title2')

        with pytest.raises(WrongWorkbookException):
            workbook.remove_sheet(ws2)

    def test_deleting_worksheet_by_name(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        title = ws2.title
        workbook.remove_sheet_by_name(title)
        assert len(workbook) == len(worksheets) - 1
        assert len(workbook.sheetnames) == len(worksheets) - 1
        assert title not in workbook.sheetnames
        assert all(ws2 != workbook[i] for i in range(len(workbook)))

    def test_deleting_worksheet_by_non_existent_name_raises_exception(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        title = 'NotATitle'

        with pytest.raises(KeyError):
            workbook.remove_sheet_by_name(title)

        assert len(workbook) == len(worksheets)
        assert len(workbook.sheetnames) == len(worksheets)

    def test_deleting_worksheet_by_index(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = 2
        ws2 = worksheets[index]
        title = ws2.title
        workbook.remove_sheet_by_index(index)
        assert len(workbook) == len(worksheets) - 1
        assert len(workbook.sheetnames) == len(worksheets) - 1
        assert title not in workbook.sheetnames
        assert all(ws2 != workbook[i] for i in range(len(workbook)))

    def test_deleting_worksheet_by_out_of_bounds_index_raises_exception(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        index = len(worksheets) + 10

        with pytest.raises(IndexError):
            workbook.remove_sheet_by_index(index)

        assert len(workbook) == len(worksheets)
        assert len(workbook.sheetnames) == len(worksheets)

    def test_getting_all_sheets(self):
        workbook = Workbook()
        worksheets = [workbook.create_sheet('Title' + str(i)) for i in range(5)]
        assert workbook.sheets == worksheets

    def test_getting_all_sheets_of_mixed_type(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.sheets
        selected_ws = [workbook.get_sheet_by_name(n) for n in ALL_NAMES]
        assert ws == selected_ws

    def test_getting_worksheets_only(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.worksheets
        selected_ws = [workbook.get_sheet_by_name(n) for n in SHEET_NAMES]
        assert ws == selected_ws

    def test_getting_chartsheets_only(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        ws = workbook.chartsheets
        selected_ws = [workbook.get_sheet_by_name(n) for n in GRAPH_NAMES]
        assert ws == selected_ws

    @pytest.mark.parametrize(
        'filepath', [str(TEST_GNUMERIC_FILE_PATH), Path(TEST_GNUMERIC_FILE_PATH)]
    )
    def test_loading_compressed_file(self, filepath):
        workbook = Workbook.load_workbook(filepath)
        assert workbook.sheetnames == list(ALL_NAMES)
        assert workbook.creation_date == datetime(
            2017, 4, 29, 17, 56, 48, tzinfo=tzutc()
        )
        assert workbook.version == '1.12.57'

    @pytest.mark.parametrize(
        'filepath', [str(TEST_SHEET_NAME_FILE_PATH), Path(TEST_SHEET_NAME_FILE_PATH)]
    )
    def test_loading_uncompressed_file(self, filepath):
        wb = Workbook.load_workbook(filepath)
        assert wb.sheetnames == list(SHEET_NAME_SHEET_NAMES)
        assert wb.creation_date == datetime(2017, 4, 29, 17, 56, 48, tzinfo=tzutc())
        assert wb.version == '1.12.28'

    def test_getting_active_sheet(self):
        workbook = Workbook.load_workbook(TEST_GNUMERIC_FILE_PATH)
        assert workbook.get_active_sheet() == workbook.get_sheet_by_name('Strings')

    def test_getting_active_sheet_from_empty_workbook(self):
        workbook = Workbook()
        assert workbook.get_active_sheet() is None

    def test_setting_active_sheet_by_index(self):
        workbook = Workbook()
        workbook.create_sheet('Sheet1')
        ws = workbook.create_sheet('Sheet2')
        workbook.create_sheet('Sheet3')
        workbook.set_active_sheet(1)
        assert workbook.get_active_sheet() == ws

    def test_setting_active_sheet_by_name(self):
        workbook = Workbook()
        workbook.create_sheet('Sheet1')
        ws = workbook.create_sheet('Sheet2')
        workbook.create_sheet('Sheet3')
        workbook.set_active_sheet('Sheet2')
        assert workbook.get_active_sheet() == ws

    def test_setting_active_sheet_by_sheet(self):
        workbook = Workbook()
        workbook.create_sheet('Sheet1')
        ws = workbook.create_sheet('Sheet2')
        workbook.create_sheet('Sheet3')
        workbook.set_active_sheet(ws)
        assert workbook.get_active_sheet() == ws


class TestWorkbookSave:
    def test_saving_compressed_file(self, monkeypatch):
        workbook = Workbook()

        mocked_open = MagicMock()
        monkeypatch.setattr(gzip, 'open', mocked_open)

        workbook.save('anyfile.gnumeric')

        mocked_open.assert_called_once_with(
            'anyfile.gnumeric', mode='wb', compresslevel=9
        )

    def test_saving_compressed_file_setting_compression_level(self, monkeypatch):
        workbook = Workbook()

        mocked_open = MagicMock()
        monkeypatch.setattr(gzip, 'open', mocked_open)

        workbook.save('anyfile.gnumeric', compress=2)

        mocked_open.assert_called_once_with(
            'anyfile.gnumeric', mode='wb', compresslevel=2
        )

    def test_saving_uncompressed_file(self, monkeypatch):
        workbook = Workbook()

        mocked_open = MagicMock()
        monkeypatch.setattr('builtins.open', mocked_open)

        workbook.save('anyfile.gnumeric', compress=False)

        mocked_open.assert_called_once_with('anyfile.gnumeric', mode='wb')
