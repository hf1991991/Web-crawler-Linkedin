import json
from copy import deepcopy


def reorder_user_dict(old_dict):
    keys = [
        'a',
        'b',
        'c',
        'd'
    ]

    new_dict = {}

    for key in keys:
        new_dict[key] = deepcopy(old_dict[key])

    return new_dict

dict = {
    'c': 1,
    'd': 4,
    'b': 3,
    'a': 2
}

print(json.dumps(dict))

dict = reorder_user_dict(dict)

print(json.dumps(dict))
