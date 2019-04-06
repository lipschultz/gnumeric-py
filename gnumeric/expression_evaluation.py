from lark import Lark, Transformer, v_args


function_map = {
    'abs': abs
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
        | sum "+" product   -> add
        | sum "-" product   -> sub
    ?product: exponentiation
        | product "*" exponentiation  -> mul
        | product "/" exponentiation  -> div
    ?exponentiation: atom
                   | atom "^" exponentiation   -> pow
    ?atom: NUMBER                                     -> number
         | string
         | "(" root ")"
         | FUNC_NAME "(" [ root ( "," root )* ] ")"   -> function
         | cell_reference                             -> cell_lookup

    ?cell_reference: (SHEETNAME "!")? "$"? COLUMN "$"? ROW   -> cell_ref
    COLUMN: CHARS~1..3
    ROW: DIGIT~1..5
    SHEETNAME: CHARS

    !logical_op: "=" | "<>" | "<" | "<=" | ">" | ">="

    string : ESCAPED_DQUOTE_STRING
           | ESCAPED_SQUOTE_STRING

    FUNC_NAME: {' | '.join(f'"{f}"i' for f in function_map.keys())}

    ESCAPED_SQUOTE_STRING : "'" _STRING_ESC_INNER "'"

    %import common.CNAME -> NAME
    %import common.SIGNED_NUMBER -> NUMBER
    %import common.DIGIT
    %import common.WORD -> CHARS
    %import common.ESCAPED_STRING -> ESCAPED_DQUOTE_STRING
    %import common._STRING_ESC_INNER
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
    from operator import add, sub, mul, truediv as div, neg
    pow = pow

    def __init__(self, cell):
        self._cell = cell

    def number(self, val):
        try:
            return int(val)
        except ValueError:
            return float(val)

    def string(self, a):
        return a.value[1:-1]

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
        return function_map[name](*args)


_parser = Lark(_grammar, start='start', parser='earley')


def evaluate(expression: str, cell):
    evaluator = ExpressionEvaluator(cell)
    result = evaluator.transform(_parser.parse(expression))
    return result
