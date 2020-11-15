from lark import Lark, Transformer, v_args


grammar = """
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

    string : ESCAPED_STRING
           | ESCAPED_SQUOTE_STRING

    FUNC_NAME: "abs"i | "test"i

    ESCAPED_SQUOTE_STRING : "'" _STRING_ESC_INNER "'"

    %import common.CNAME -> NAME
    %import common.SIGNED_NUMBER -> NUMBER
    %import common.DIGIT
    %import common.WORD -> CHARS
    %import common.ESCAPED_STRING
    %import common._STRING_ESC_INNER
    %import common.WS_INLINE
    %ignore WS_INLINE
"""


function_map = {
    'abs': abs,
    '1test': lambda: 15,
}


@v_args(inline=True)    # Affects the signatures of the methods
class TreeEvaluator(Transformer):
    from operator import add, sub, mul, truediv as div, neg
    from operator import eq, lt, le, gt, ge, ne
    number = float
    pow = pow

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

    # --- Text Operations ---
    def concat(self, a, b):
        return str(a) + str(b)

    def function(self, name, *args):
        return function_map[name](*args)


parser = Lark(grammar, start='start', parser='earley')  # lalr
tree_evaluator = TreeEvaluator()
evaluator = lambda text: tree_evaluator.transform(parser.parse(text))

def main():
    while True:
        try:
            s = input('> ')
        except EOFError:
            break
        print(evaluator(s))


def ptest(test_case, expected=None):
    result = evaluator(test_case)
    print(f'{test_case}: {result}')
    if expected is not None and result != expected:
        print(f'*** Mismatch: expected={expected}, actual={result}')


for tc in (('=abs(-1)', 1.0), ('=c1', "cell lookup:(None, 'c', '1')"), ('=$c$1', "cell lookup:(None, 'c', '1')"), ('=Investments!$c$1', "cell lookup:('Investments', 'c', '1')"), ('=abs(5-6)', 1.0), ('="asdf"&abs(-2)', 'asdf2.0'), ("='asdf'&abs(-2)", 'asdf2.0'), ("=2>0", True), ("=2=0", False), ('=abs(-1)=1', True), ):
    ptest(*tc)

#main()
