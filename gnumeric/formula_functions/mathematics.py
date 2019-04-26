import itertools

from gnumeric.formula_functions.argument_helpers import get_just_numeric


def gnm_sum(*values):
    return sum(get_just_numeric(values))


local_objects = locals().copy()
functions = {name[4:]: obj for name, obj in local_objects.items() if name.startswith('gnm_')}
