import math
import re
from typing import Set, Union

from lark import Lark, Transformer, v_args
from lark.exceptions import VisitError, UnexpectedCharacters

from gnumeric.evaluation_errors import EvaluationError, ExpressionEvaluationException
from gnumeric.formula_functions import mathematics, statistics
from gnumeric.utils import column_from_spreadsheet, row_from_spreadsheet, coordinate_to_spreadsheet

function_map = {
    'abs': abs,
    'len': lambda val: len(str(val)),
    **mathematics.functions,
    **statistics.functions,
}


class CellReference:
    def __init__(self, formula_cell, ss_row, ss_col, sheet, row_is_fixed, col_is_fixed):
        self.formula_cell = formula_cell
        self.ss_row = ss_row
        self.ss_col = ss_col
        self.sheet = sheet
        self.row_is_fixed = row_is_fixed
        self.col_is_fixed = col_is_fixed

    @property
    def cell(self):
        try:
            sheet = self.formula_cell.worksheet if self.sheet is None else self.formula_cell.worksheet.workbook.get_sheet_by_name(self.sheet)
        except KeyError:
            raise ExpressionEvaluationException(EvaluationError.REF)

        offset_row, offset_col = self.formula_cell.value.reference_coordinate_offset

        col = column_from_spreadsheet(self.ss_col)
        row = row_from_spreadsheet(self.ss_row)

        if not self.col_is_fixed:
            col += offset_col
        if not self.row_is_fixed:
            row += offset_row

        return sheet.cell(row, col, create=True)

    @classmethod
    def create_from_cell_reference(cls, formula_cell, sheet, ss_col, ss_row):
        col_fixed = ss_col.startswith('$')
        row_fixed = ss_row.startswith('$')
        return cls(formula_cell, ss_row, ss_col, sheet, row_is_fixed=row_fixed, col_is_fixed=col_fixed)


_grammar = f"""
    ?start: "=" root
          | "+" root
    ?root: logical_stmt
    ?logical_stmt: text_stmt
                 | logical_stmt logical_op text_stmt   -> logical
    ?text_stmt: sum
            | text_stmt "&" sum   -> concat
    ?sum: product
        | sum /[+-]/ product   -> arithmetic
    ?product: exponentiation
        | product /[*\\/]/ exponentiation  -> arithmetic
    ?exponentiation: atom
                   | atom /\\^/ exponentiation   -> arithmetic

    ?function_argument: root
                      | cell_range
    ?cell_range: sheet_cell_reference ":" cell_reference

    ?atom: NUMBER                                     -> number
         | "#REF!"                                    -> error_ref
         | ATOMIC_STR                                 -> atomic_string
         | FUNC_NAME "(" [ function_argument ( "," function_argument )* ] ")"   -> function
         | string
         | "(" root ")"
         // | sheet_cell_reference                       -> cell_lookup

    !logical_op: "=" | "<>" | "<" | "<=" | ">" | ">="

    ?sheet_cell_reference: (SHEETNAME "!")? cell_reference   -> sheet_cell_ref
    ?cell_reference: "$"? COLUMN "$"? ROW                    -> cell_ref
    COLUMN: LETTER~1..3
    ROW: DIGIT~1..5
    SHEETNAME: LETTER+

    string : ESCAPED_STRING

    // FUNC_NAME: {' | '.join(f'"{f}"i' for f in function_map.keys())}
    FUNC_NAME: LETTER ("_"|LETTER|DIGIT|".")*
    ATOMIC_STR: (LETTER | "$" | "'") ("_"|LETTER|DIGIT|"."|"!"|"$"|"'"|" ")*

    %import common.DIGIT
    %import common.SIGNED_NUMBER -> NUMBER
    %import common.LETTER
    %import common.ESCAPED_STRING
    %import common.WS_INLINE
    %ignore WS_INLINE
"""


def to_str(value) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    elif isinstance(value, bool):
        value = 'TRUE' if value else 'FALSE'
    return str(value)


@v_args(inline=True)
class ExpressionEvaluator(Transformer):
    def __init__(self, cell):
        self._cell = cell
        self.referenced_cells = set()

    def number(self, val):
        try:
            return int(val)
        except ValueError:
            return float(val)

    def string(self, a):
        return a.value[1:-1]

    def error_ref(self):
        raise ExpressionEvaluationException(EvaluationError.REF)

    def atomic_string(self, name):
        if name == 'TRUE':
            return True
        elif name == 'FALSE':
            return False
        else:
            res = re.match(r"('?[A-Za-z0-9_ ]+'?!)?(\$?[A-Za-z]{1,3})(\$?\d{1,5})", name)
            if res:
                sheet, col, row = res.groups()
                sheet = sheet[:-1] if sheet is not None else None

                if sheet is not None and sheet.startswith("'") and sheet.endswith("'"):
                    sheet = sheet[1:-1]

                cell_value = self.cell_lookup((sheet, col, row))
                if cell_value is not None and not isinstance(cell_value, (int, float, str)):
                    cell_value = cell_value.value
                return cell_value

        raise ExpressionEvaluationException(EvaluationError.NAME)

    def logical_op(self, op):
        return op.value

    def logical(self, a, op, b):
        if isinstance(a, str):
            a = str(a).lower()
        if isinstance(b, str):
            b = str(b).lower()

        # Numbers are always less than strings and strings always less than bools, so to simplify the logic below, just use different values for a and b if there's a type mismatch
        a_type = type(a) if type(a) != int else float
        b_type = type(b) if type(b) != int else float
        if a_type != b_type:
            type_val_mapper = {
                float: 0,
                str: 50,
                bool: 100
            }
            a = type_val_mapper[a_type]
            b = type_val_mapper[b_type]

        if op == '=':
            return a == b
        elif op == '<>':
            return a != b
        elif op == '<':
            return a < b
        elif op == '<=':
            return a <= b
        elif op == '>':
            return a > b
        elif op == '>=':
            return a >= b
        else:
            raise ValueError(f'Unrecognized logical operator: {op}')

    def arithmetic(self, a, op, b):
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
            raise ExpressionEvaluationException(EvaluationError.VALUE)

        if op == '+':
            return a + b
        elif op == '-':
            return a - b
        elif op == '*':
            return a * b
        elif op == '/':
            return a / b
        elif op == '^':
            try:
                return math.pow(a, b)
            except OverflowError as ex:
                raise ExpressionEvaluationException(EvaluationError.NUM)

    def cell_ref(self, *ref) -> CellReference:
        return CellReference.create_from_cell_reference(self._cell, None, ref[0].value, ref[1].value)

    def sheet_cell_ref(self, *ref) -> CellReference:
        if len(ref) > 1:
            ref[1].sheet = ref[0].value
            ref = ref[1]
        else:
            ref = ref[0]
        return ref

    def cell_range(self, start, end):
        start_cell = start.cell
        end_cell = end.cell
        cells = start_cell.worksheet.get_cell_collection(start_cell, end_cell, include_empty=True, create_cells=True)
        self.referenced_cells.update(cells)
        return cells

    def cell_lookup(self, ref):
        cell_ref = CellReference.create_from_cell_reference(self._cell, ref[0], ref[1], ref[2])
        cell = cell_ref.cell
        self.referenced_cells.add(cell)
        if cell is None:
            return 0
        return cell.result

    def concat(self, a, b):
        a = to_str(a)
        b = to_str(b)
        return a + b

    def function(self, name, *args):
        try:
            return function_map[name.lower()](*args)
        except KeyError as ex:
            raise ExpressionEvaluationException(EvaluationError.NAME) from ex
        except TypeError as ex:
            msg = str(ex)
            if msg.startswith('bad operand type'):
                raise ExpressionEvaluationException(EvaluationError.VALUE) from ex
            elif 'takes exactly' in str(ex):
                raise ExpressionEvaluationException(EvaluationError.NA) from ex
            print(type(ex), ex)
            raise ex


_parser = Lark(_grammar, start='start', parser='earley')


def parse(expression: str):
    try:
        tree = _parser.parse(expression)
    except UnexpectedCharacters:
        return EvaluationError.VALUE

    # print(tree.pretty())
    return tree


def _full_evaluation(expression: str, cell):
    tree = parse(expression)

    if isinstance(tree, EvaluationError):
        result, referenced_cells = tree, set()
    else:
        evaluator = ExpressionEvaluator(cell)
        try:
            result = evaluator.transform(tree)
            result, referenced_cells = result, evaluator.referenced_cells
        except VisitError as ex:
            if isinstance(ex.orig_exc, ZeroDivisionError):
                result, referenced_cells = EvaluationError.DIV0, set()
            elif isinstance(ex.orig_exc, ExpressionEvaluationException):
                result, referenced_cells = ex.orig_exc.error, set()
            else:
                print(type(ex.orig_exc), ex.orig_exc)
                raise ex.orig_exc

    return result, referenced_cells


def evaluate(expression: str, cell) -> Union[str, int, float, EvaluationError]:
    result, _ = _full_evaluation(expression, cell)
    return result


def get_referenced_cells(expression: str, cell) -> Set:
    _, references = _full_evaluation(expression, cell)
    return references
