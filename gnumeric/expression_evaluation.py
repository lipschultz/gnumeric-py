from lark import Lark, Transformer, v_args


function_map = {
    'abs': abs
}


_grammar = f"""
    ?start: "=" root
          | "+" root
    ?root: logical_stmt
    ?logical_stmt: text_stmt
                 | logical_stmt "=" text_stmt   -> eq
                 | logical_stmt "<" text_stmt   -> lt
                 | logical_stmt "<=" text_stmt   -> le
                 | logical_stmt ">" text_stmt   -> gt
                 | logical_stmt ">=" text_stmt   -> ge
                 | logical_stmt "<>" text_stmt   -> ne
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
    ?atom: NUMBER                                           -> number
         | string
         | "(" root ")"
         | FUNC_NAME "(" [ root ( "," root )* ] ")"   -> function
         | cell_reference                                   -> cell_lookup

    ?cell_reference: (SHEETNAME "!")? "$"? COLUMN "$"? ROW   -> cell_ref
    COLUMN: CHARS~1..3
    ROW: DIGIT~1..5
    SHEETNAME: CHARS

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


@v_args(inline=True)
class ExpressionEvaluator(Transformer):
    from operator import add, sub, mul, truediv as div, neg
    from operator import eq, lt, le, gt, ge, ne
    number = float
    pow = pow

    def __init__(self, cell):
        self._cell = cell

    def string(self, a):
        return a.value[1:-1]

    def cell_ref(self, *ref):
        ref = [r.value for r in ref]
        if len(ref) == 3:
            return tuple(ref)
        else:
            return (None, *ref)

    def cell_lookup(self, ref):
        return 'cell lookup:' + str(ref)

    def concat(self, a, b):
        return str(a) + str(b)

    def function(self, name, *args):
        return function_map[name](*args)


_parser = Lark(_grammar, start='start', parser='earley')


def evaluate(expression: str, cell):
    evaluator = ExpressionEvaluator(cell)
    result = evaluator.transform(_parser.parse(expression))
    return result
