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


def column_to_spreadsheet(col_int, abs_ref=False):
    """
        Convert 0-indexed column number into standard spreadsheet notation.  For example: `30` -> `'AE'`.

        Raises `IndexError` when `col_int` is negative.
    """
    if col_int < 0:
        raise IndexError('The column index must be positive: ' + str(col_int))

    column = chr((col_int % 26) + ord('A'))
    while col_int >= 26:
        col_int = col_int // 26 - 1
        column = chr((col_int % 26) + ord('A')) + column

    abs_col = '$' if abs_ref else ''
    return '%s%s' % (abs_col, column)


def column_from_spreadsheet(col_letter):
    """
    Convert standard spreadsheet notation for a column into a 0-indexed column number.  For example: `'AE'` -> `30`.

    Capitalization of the letters is ignored, so `'ae'` and `'Ae'` also convert to `30`.

    Raises `IndexError` is any character in `col_letter` is not a letter.
    """
    col_letter = col_letter.replace('$', '')
    index = 0
    for l in col_letter:
        if not l.isalpha():
            raise IndexError('Illegal character in column name: ' + l)
        index = index * 26 + (ord(l) - ord('A') + 1)
    return index - 1


def row_to_spreadsheet(row):
    """
    Convert a 0-indexed row index into the standard spreadsheet row (which is 1-indexed).

    Raises `IndexError` if row is less than 1.
    """
    if row < 0:
        raise IndexError('row must be greater than or equal to 0')

    return row + 1


def row_from_spreadsheet(row):
    """
    Convert the standard spreadsheet row numbers (1-indexed) to 0-indexed index.

    Raises `IndexError` if row is less than 1.
    """
    if row < 1:
        raise IndexError('row must be greater than or equal to 1')

    return row - 1
