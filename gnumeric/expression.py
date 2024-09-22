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

from typing import Set, Tuple, Union

from gnumeric import expression_evaluation, utils
from gnumeric.evaluation_errors import EvaluationError


class Expression:
    # TODO: what to do if originating cell's text changes or is deleted?

    def __init__(self, id, worksheet, cell):
        """
        :param cell: The cell this Expression was created from.  It could be the originating cell or just a cell that
            uses this expression.  It's used when `id` is `None` (i.e. the expression isn't shared between cells).
        """
        self.__exprid = id
        self.__worksheet = worksheet
        self.__cell = cell

    def __get_raw_originating_cell(self):
        """
        Returns a (int, int, str) tuple, where the values are (row, column, text of the cell).
        """
        if self.__exprid is not None:
            coords, text = self.__worksheet.get_expression_map()[self.__exprid]
            return int(coords[0]), int(coords[1]), text
        else:
            return self.__cell.row, self.__cell.column, self.__cell.text

    @property
    def id(self):
        """
        The expression id used to uniquely identify the expression within the sheet.  This will be `None` if the
        expression isn't shared between cells.
        """
        return self.__exprid

    @property
    def original_text(self) -> str:
        """
        Returns the text of the expression, with cell references from the perspective of the cell where the expression
        is stored (i.e. the original cell).
        """
        return self.__get_raw_originating_cell()[2]

    @property
    def text(self):
        """
        Returns the text of the exprsesion, with cell references updated to be from the perspective of the cell using the expression.
        """
        # TODO: fix this to actually return what it's supposed to
        raise NotImplementedError

    @property
    def reference_coordinate_offset(self) -> Tuple[int, int]:
        """
        The (row, col) offset to translate the original coordinates into the coordinates based at the current cell.
        """
        original_coordinates = self.get_originating_cell_coordinate()
        current_coordinates = self.__cell.coordinate
        return current_coordinates[0] - original_coordinates[0], current_coordinates[
            1
        ] - original_coordinates[1]

    @property
    def value(self):
        """
        Returns the result of the expression's evaluation.
        """
        return expression_evaluation.evaluate(self.original_text, self.__cell)

    @property
    def worksheet(self):
        return self.__worksheet

    def get_originating_cell_coordinate(
        self, representation_format='index'
    ) -> Union[Tuple[int, int], str]:
        """
        Returns the cell coordinate for the cell Gnumeric is using to store the expression.

        :param representation_format: For spreadsheet notation, use `'spreadsheet'`, for 0-indexed (row, column)
            notation, use 'index' (default).
        :return: A `str` if `representation_format` is `'spreadsheet'` and a tuple of ints `(int, int)` if 'index'.
        """
        row, col = self.__get_raw_originating_cell()[:2]
        if representation_format == 'index':
            return row, col
        elif representation_format == 'spreadsheet':
            return utils.coordinate_to_spreadsheet(row, col)

    def get_originating_cell(self):
        """
        Returns the cell Gnumeric is using to store the expression.
        """
        return self.__worksheet.cell(
            *self.get_originating_cell_coordinate(), create=False
        )

    def get_all_cells(self, sort=False):
        """
        Returns a list of all cells using this expression.

        Use `sort` to specify whether the cells should be sorted.  If `False` (default), then no sorting will take
        place.  If `sort` is `"row"`, then sorting will occur by row first, then by column within each row.  If `sort`
        is `"column"`, then the opposite will happen: first sort by column, then by row within each column.
        """
        if self.__exprid is None:
            return [self.__cell]
        else:
            return self.__worksheet.get_all_cells_with_expression(
                self.__exprid, sort=sort
            )

    def get_referenced_cells(self) -> Union[Set, EvaluationError]:
        return expression_evaluation.get_referenced_cells(
            self.original_text, self.__cell
        )

    def __str__(self):
        return self.original_text

    def __repr__(self):
        return 'Expression(id=%s, text="%s", ws=%s, cell=(%d, %d))' % (
            self.id,
            self.original_text,
            self.__worksheet,
            self.__cell.row,
            self.__cell.column,
        )
