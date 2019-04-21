from gnumeric.evaluation_errors import EvaluationError, ExpressionEvaluationException


def flatten_just_type(values, types):
    keep_values = []
    for v in values:
        if isinstance(v, EvaluationError):
            raise ExpressionEvaluationException(v)
        elif isinstance(v, types):
            keep_values.append(v)
        elif isinstance(v, list):
            keep_values.extend(flatten_just_type(v, types))
        elif hasattr(v, 'get_value'):
            cell_value = v.get_value(compute_expression=True)
            keep_values.extend(flatten_just_type([cell_value], types))

    return keep_values


def get_just_numeric(values):
    return flatten_just_type(values, (int, float))
