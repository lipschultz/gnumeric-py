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

import pytest

from gnumeric import utils


class TestBasicUtility:
    def test_column_index_to_letter_works_for_single_char_columns(self):
        name = utils.column_to_spreadsheet(3)
        assert name == 'D'

    def test_column_index_to_letter_works_for_multi_char_columns(self):
        name = utils.column_to_spreadsheet(30)
        assert name == 'AE'

    def test_column_index_to_letter_raises_IndexError_when_negative_column_given(self):
        with pytest.raises(IndexError):
            utils.column_to_spreadsheet(-3)

    def test_column_index_to_letter_returns_absolute_when_absolute_ref_requested(self):
        name = utils.column_to_spreadsheet(30, True)
        assert name == '$AE'

    def test_column_letter_to_index_works_for_single_char_columns(self):
        idx = utils.column_from_spreadsheet('D')
        assert idx == 3

    def test_column_letter_to_index_works_for_multi_char_columns(self):
        idx = utils.column_from_spreadsheet('AE')
        assert idx == 30

    def test_column_letter_to_index_raises_IndexError_when_non_alphabetic_character_in_name(
        self,
    ):
        with pytest.raises(IndexError):
            utils.column_from_spreadsheet('A:E')

    def test_column_letter_to_index_ignores_absolute_references(self):
        idx = utils.column_from_spreadsheet('$A$E')
        assert idx == 30

    def test_row_to_spreadsheet_converts_valid_values(self):
        assert utils.row_to_spreadsheet(16) == '17'

    def test_row_to_spreadsheet_creates_absolute_reference(self):
        assert utils.row_to_spreadsheet(16, True) == '$17'

    def test_row_to_spreadsheet_raises_exception_on_less_than_0(self):
        with pytest.raises(IndexError):
            utils.row_to_spreadsheet(-1)

    def test_row_from_spreadsheet_converts_valid_values(self):
        assert utils.row_from_spreadsheet('15') == 14

    def test_row_from_spreadsheet_ignores_absolute_reference(self):
        assert utils.row_from_spreadsheet('$15') == 14

    def test_row_from_spreadsheet_raises_exception_on_less_than_1(self):
        with pytest.raises(IndexError):
            utils.row_from_spreadsheet('0')

    def test_row_from_spreadsheet_ignores_absolute_references(self):
        idx = utils.row_from_spreadsheet('$31')
        assert idx == 30

    def test_coordinate_to_spreadsheet_using_col_param(self):
        assert utils.coordinate_to_spreadsheet(17, 30, abs_ref_row=True) == 'AE$18'

    def test_coordinate_to_spreadsheet_using_just_coord_param(self):
        assert utils.coordinate_to_spreadsheet((17, 30), abs_ref_col=True) == '$AE18'

    def test_coordinate_from_spreadsheet(self):
        assert utils.coordinate_from_spreadsheet('AE$18') == (17, 30)
