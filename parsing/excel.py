from lark import Lark, Transformer, v_args


try:
    input = raw_input   # For Python2 compatibility
except NameError:
    pass


grammar = """
    ?start: "=" text_op
          | "+" text_op
    ?text_op: sum
            | text_op "&" sum   -> concat
    ?sum: product
        | sum "+" product   -> add
        | sum "-" product   -> sub
    ?product: exponentiation
        | product "*" exponentiation  -> mul
        | product "/" exponentiation  -> div
    ?exponentiation: atom
                   | atom "^" exponentiation   -> pow
    ?atom: NUMBER                                           -> number
         | string_val                                       -> string
         | "(" text_op ")"
         | FUNC_NAME "(" [ text_op ( "," text_op )* ] ")"   -> function
         | cell_reference                                   -> cell_lookup

    ?cell_reference: (SHEETNAME "!")? "$"? COLUMN "$"? ROW   -> cell_ref
    COLUMN: CHARS~1..3
    ROW: DIGIT~1..5
    SHEETNAME: CHARS

    string_val: "'" /.+/ "'"
              | "\"" /.+/ "\""

    FUNC_NAME: "abs"i | "test"i

    %import common.CNAME -> NAME
    %import common.SIGNED_NUMBER -> NUMBER
    %import common.DIGIT
    %import common.WORD -> CHARS
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
    number = float
    pow = pow

    def string(self, a):
        print(a)
        return 'asdf'

    def concat(self, a, b):
        return str(a.value) + str(b.value)

    def function(self, name, *args):
        return function_map[name](*args)

    def cell_ref(self, *ref):
        ref = [r.value for r in ref]
        if len(ref) == 3:
            return tuple(ref)
        else:
            return (None, *ref)

    def cell_lookup(self, ref):
        return 'cell lookup:' + str(ref)


parser = Lark(grammar)
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


for tc in (('=abs(-1)', 1.0), ('=c1', "cell lookup:(None, 'c', '1')"), ('=$c$1', "cell lookup:(None, 'c', '1')"), ('=Investments!$c$1', "cell lookup:('Investments', 'c', '1')"), ('=abs(5-6)', 1.0), ('="asdf"&abs(-2)', 'asdf2')):
    ptest(*tc)

#main()
