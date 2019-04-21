import enum


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
