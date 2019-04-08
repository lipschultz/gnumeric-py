import enum
import math

from lark import Lark, Transformer, v_args
from lark.exceptions import VisitError


class EvaluationError(enum.Enum):
    DIV0 = '#DIV/0!'
    VALUE = '#VALUE!'
    NA = '#N/A'
    NAME = '#NAME?'
    NUM = '#NUM!'  # =10000000000^1000000000
    REF = '#REF!'
    NULL = '#NULL!'  # occurs when the intersection of two areas don't actually intersect


class ExpressionEvaluationException(Exception):
    """
    Evaulating the expression has failed.
    """

    def __init__(self, error: EvaluationError, msg=None):
        msg = msg or error.value
        super().__init__(msg)
        self.error = error


function_map = {
    'abs': abs,
    'len': lambda val: len(str(val)),
}


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
    ?atom: NUMBER                                     -> number
         | string
         | "(" root ")"
         | FUNC_NAME "(" [ root ( "," root )* ] ")"   -> function
         | cell_reference                             -> cell_lookup
         | "#REF!"                                    -> error_ref
         | FUNC_NAME                                  -> atomic_string

    !logical_op: "=" | "<>" | "<" | "<=" | ">" | ">="

    ?cell_reference: (SHEETNAME "!")? "$"? COLUMN "$"? ROW   -> cell_ref
    COLUMN: LETTER~1..3
    ROW: DIGIT~1..5
    SHEETNAME: LETTER+

    string : ESCAPED_STRING

    // FUNC_NAME: {' | '.join(f'"{f}"i' for f in function_map.keys())}
    FUNC_NAME: LETTER ("_"|LETTER|DIGIT|".")*

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

    def __init__(self, sheet):
        self._sheet = sheet

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

    def cell_ref(self, *ref):
        ref = [r.value for r in ref]
        if len(ref) == 3:
            return tuple(ref)
        else:
            return (None, *ref)

    def cell_lookup(self, ref):
        return 'cell lookup:' + str(ref)

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


def evaluate(expression: str, sheet):
    evaluator = ExpressionEvaluator(sheet)

    tree = _parser.parse(expression)
    # print(tree.pretty())
    try:
        result = evaluator.transform(tree)
    except VisitError as ex:
        if isinstance(ex.orig_exc, ZeroDivisionError):
            return EvaluationError.DIV0
        elif isinstance(ex.orig_exc, ExpressionEvaluationException):
            return ex.orig_exc.error

        print(type(ex.orig_exc), ex.orig_exc)
        raise ex.orig_exc

    return result
