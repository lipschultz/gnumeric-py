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
import datetime
import time
from typing import Optional, Tuple, Union

from lxml import etree

from gnumeric.exceptions import UnrecognizedCellTypeException
from gnumeric.expression import Expression

VALUE_TYPE_EXPR = -10
VALUE_TYPE_EMPTY = 10
VALUE_TYPE_BOOLEAN = 20
VALUE_TYPE_INTEGER = 30
VALUE_TYPE_FLOAT = 40
VALUE_TYPE_ERROR = 50
VALUE_TYPE_STRING = 60
VALUE_TYPE_CELLRANGE = 70
VALUE_TYPE_ARRAY = 80


class Cell:
    _instances = {}

    _BASE_DATETIME = datetime.datetime(1899, 12, 30)

    def __new__(cls, cell_element, style_region_element, worksheet, ns):
        key = (cell_element, style_region_element, worksheet)
        instance = cls._instances.get(key)
        if not instance:
            instance = super(Cell, cls).__new__(cls)
            instance.__cached_value = None
            instance.__time_cached_value_set = None
            cls._instances[key] = instance
        return instance

    def __init__(self, cell_element, style_region_element, worksheet, ns):
        self.__cell = cell_element
        self.__style_region = style_region_element
        self.__worksheet = worksheet
        self.__ns = ns

    def __get_style_element(self):
        elements = self.__style_region.xpath('./gnm:Style', namespaces=self.__ns)
        return elements[0] if elements else None

    def __set_expression_id(self, expr_id: str) -> None:
        self.__cell.set('ExprID', expr_id)

    @property
    def worksheet(self):
        return self.__worksheet

    @property
    def column(self) -> int:
        """
        The column this cell belongs to (0-indexed).
        """
        return int(self.__cell.get('Col'))

    @property
    def row(self) -> int:
        """
        The row this cell belongs to (0-indexed).
        """
        return int(self.__cell.get('Row'))

    @property
    def coordinate(self) -> Tuple[int, int]:
        """
        The (row, column) of the cell.
        """
        return self.row, self.column

    @property
    def text(self) -> Optional[str]:
        """
        Returns the raw value stored in the cell.  The text will be `None` if the cell is empty.
        :return: str or `None`
        """
        return self.__cell.text

    @property
    def value_type(self) -> int:
        """
        Returns the type of value stored in the cell:
         - VALUE_TYPE_EXPR = -10
         - VALUE_TYPE_EMPTY = 10
         - VALUE_TYPE_BOOLEAN = 20
         - VALUE_TYPE_INTEGER = 30
         - VALUE_TYPE_FLOAT = 40
         - VALUE_TYPE_ERROR = 50
         - VALUE_TYPE_STRING = 60
         - VALUE_TYPE_CELLRANGE = 70
         - VALUE_TYPE_ARRAY = 80
        """
        value_type = self.__cell.get('ValueType')
        if value_type is not None:
            return int(value_type)
        elif self.__cell.get('ExprID') is not None or self.text.startswith('='):
            return VALUE_TYPE_EXPR
        else:
            raise UnrecognizedCellTypeException('Cell is: "' + str(etree.tostring(self.__cell)) + '"')

    def is_datetime(self) -> bool:
        style_format = self.text_format.lower()
        return self.value_type == VALUE_TYPE_FLOAT and any(pattern in style_format for pattern in ('yy', 'm', 'd', 'h', 's'))

    def __set_type(self, value_type: int):
        if value_type == VALUE_TYPE_EXPR:
            if 'ValueType' in self.__cell.keys():
                self.__cell.attrib.pop('ValueType')
        else:
            self.__cell.set('ValueType', str(value_type))

    def get_value(self, *, compute_expression: bool=False) -> Union[bool, int, float, Expression, str]:
        """
        Gets the value stored in the cell, converted into the appropriate Python datatype when possible.

        If the cell is an expression: If `compute_expression` is True, the the result of the expression is returned, otherwise an Expression object is returned.
        """
        value = self.text
        if self.value_type == VALUE_TYPE_BOOLEAN:
            return value.lower() == 'true'
        elif self.value_type == VALUE_TYPE_INTEGER:
            return int(value)
        elif self.value_type == VALUE_TYPE_FLOAT:
            return float(value)
        elif self.value_type == VALUE_TYPE_EXPR:
            expression = Expression(self.__cell.get('ExprID'), self.__worksheet, self)
            if not compute_expression:
                return expression
            else:
                if self.__cached_value is None:
                    self.__cached_value = 0
                    self.__cached_value = expression.value
                return self.__cached_value
        else:
            return value

    def set_value(self, value, *, value_type: str = 'infer') -> None:
        """
        Sets the value stored in the cell.

        If `value_type` is:
        - one of the `VALUE_TYPE_` constants, then that value type will be used
        - ``keep'`, then the cell will keep its type
        - `'infer'`, then it tries to guess the type based on the value being set:
            - if `value` is of type `bool`, then `VALUE_TYPE_BOOLEAN`
            - if `value` is of type `int`, then `VALUE_TYPE_INTEGER`
            - if `value` is of type `float`, then `VALUE_TYPE_FLOAT`
            - if `value` is an empty string or `None`, then `VALUE_TYPE_EMPTY`
            - if `value` is a string that starts with `=` or is an `Expression` object, then `VALUE_TYPE_EXPR`
            - if `value` is anything else, then `VALUE_TYPE_STRING`

        The default `value_type` is `'infer'`.

        Warning: This method does no type checking, so it is possible to save a string into a cell whose type is
        `VALUE_TYPE_INTEGER`.  This could result in problems when opening the workbook in Gnumeric.
        """
        val_types = {bool: VALUE_TYPE_BOOLEAN, int: VALUE_TYPE_INTEGER, float: VALUE_TYPE_FLOAT}
        if value_type == 'infer':
            if type(value) in val_types:
                value_type = val_types[type(value)]
            elif value in ('', None):
                value_type = VALUE_TYPE_EMPTY
            elif isinstance(value, Expression) or value[0] == '=':
                value_type = VALUE_TYPE_EXPR
            elif self.value_type in (VALUE_TYPE_EMPTY, VALUE_TYPE_BOOLEAN, VALUE_TYPE_INTEGER, VALUE_TYPE_FLOAT,
                                     VALUE_TYPE_ERROR):
                value_type = VALUE_TYPE_STRING
            else:
                value_type = self.value_type
        elif value_type == 'keep':
            value_type = self.value_type

        if value_type == VALUE_TYPE_BOOLEAN:
            self.__cell.text = str(bool(value)).upper()
        elif value_type == VALUE_TYPE_EMPTY:
            self.__cell.text = None
        elif value_type == VALUE_TYPE_EXPR and isinstance(value, Expression):
            if self.__worksheet != value.worksheet:
                raise NotImplementedError('Copying expression to different worksheet is not yet supported')  # TODO

            expr_id = str(value.id)
            if expr_id not in self.__worksheet.get_expression_map() and value.id is None:
                expr_id = str(max(int(k) for k in self.__worksheet.get_expression_map()) + 1)

            if len(value.get_all_cells()) == 1 and value.get_originating_cell() == self:
                # copying expression over itself, so don't add it to cells using this expression
                self.__cell.text = value.value
            else:
                # the expression is shared, so use the expression id
                if len(value.get_all_cells()) == 1:
                    value.get_originating_cell().__set_expression_id(expr_id)
                self.__cell.text = None
                self.__set_expression_id(expr_id)
        else:
            self.__cell.text = str(value)

        self.__set_type(value_type)
        self.__cached_value = self.get_value(compute_expression=True)

    value = property(get_value, set_value, doc='Get or set the value in the cell, converted into the correct type.')

    @property
    def result(self):
        """
        Gets the result of the cell.  For expressions, it returns the result of the expression.  For literals, it returns the literal.
        """
        value = self.get_value(compute_expression=True)
        if value is None:
            value = 0
        elif self.is_datetime():
            value = self._BASE_DATETIME + datetime.timedelta(days=value)
        return value

    @property
    def text_format(self) -> str:
        """
        The format string used to format the text in the cell for display.  This is the "Number Format" in Gnumeric.
        """
        return str(self.__style_region.xpath('./gnm:Style/@Format', namespaces=self.__ns)[0])

    def __str__(self) -> str:
        return repr(self.value)

    def __repr__(self) -> str:
        return 'Cell[%s, (%d, %d), ws="%s"]' % (str(self), self.row, self.column, self.__worksheet.title)

    def __eq__(self, other) -> bool:
        return (isinstance(other, Cell) and
                self.__worksheet == other.__worksheet and
                self.row == other.row and
                self.column == other.column)

    def __hash__(self) -> int:
        return hash(self.__cell)
