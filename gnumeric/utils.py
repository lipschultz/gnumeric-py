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
from typing import Optional, Union, Tuple


def column_to_spreadsheet(col_int: int, abs_ref: bool = False) -> str:
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


def column_from_spreadsheet(col_letter: str) -> int:
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


def row_to_spreadsheet(row: int, abs_ref: bool = False) -> str:
    """
    Convert a 0-indexed row index into the standard spreadsheet row (which is 1-indexed).  If `abs_ref` is True, then
    the returned value will be an absolute reference to the row.

    Raises `IndexError` if row is less than 1.
    """
    if row < 0:
        raise IndexError('row must be greater than or equal to 0')

    abs_ref = '$' if abs_ref else ''
    return abs_ref + str(row + 1)


def row_from_spreadsheet(row: str) -> int:
    """
    Convert the standard spreadsheet row numbers (1-indexed) to 0-indexed index.

    Raises `IndexError` if row is less than 1.
    """
    row = int(row.replace('$', ''))
    if row < 1:
        raise IndexError('row must be greater than or equal to 1')

    return row - 1


def coordinate_to_spreadsheet(coord: Union[int, Tuple[int, int]], col: Optional[int] = None, abs_ref_row: bool = False, abs_ref_col: bool = False) -> str:
    """
    Convert a coordinate into spreadsheet notation.  If `col` is `None`, then `coord` is assumed to be a (row, column)
    tuple.  If `col` is not `None`, then `coord` is the row and col is the column.

    If `abs_ref_row` is True, then the resulting coordinate will have the row be an absolute reference.  Then same is
    true for `abs_ref_col` about the column.

    Example (6, 2) -> 'C7'
    """
    if col is None:
        coord, col = coord

    return column_to_spreadsheet(col, abs_ref_col) + row_to_spreadsheet(coord, abs_ref_row)


def coordinate_from_spreadsheet(coord: str) -> Tuple[int, int]:
    """
    Convert a coordinate from spreadsheet notation into a (row, column) tuple.

    Example `'C$7'` -> `(6, 2)`
    """
    coord = coord.replace('$', '')

    first_row_position = None
    i = 0
    while first_row_position is None and i < len(coord):
        if not coord[i].isalpha():
            first_row_position = i

        i += 1

    return row_from_spreadsheet(coord[first_row_position:]), column_from_spreadsheet(coord[:first_row_position])
