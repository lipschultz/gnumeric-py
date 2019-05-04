import functools

from gnumeric.formula_functions.argument_helpers import get_just_numeric


def gnm_product(*values):
    values = get_just_numeric(values)
    if len(values) == 0:
        return 0
    else:
        return functools.reduce(lambda x, y: x*y, values)

def gnm_sum(*values):
    return sum(get_just_numeric(values))


local_objects = locals().copy()
functions = {name[4:]: obj for name, obj in local_objects.items() if name.startswith('gnm_')}
