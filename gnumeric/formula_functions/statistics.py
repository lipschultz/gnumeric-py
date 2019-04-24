from gnumeric.formula_functions.argument_helpers import get_just_numeric


def gnm_max(values):
    values = get_just_numeric(values)
    if len(values) == 0:
        return 0
    else:
        return max(values)


local_objects = locals().copy()
functions = {name[4:]: obj for name, obj in local_objects.items() if name.startswith('gnm_')}
