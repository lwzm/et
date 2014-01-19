#!/usr/bin/env python3
# xls

import collections
import json
import glob
import logging
import pprint
import re
import time

from tornado import template

value_type_names = {
    int: "int",
    str: "string",
    float: "float",
    bool: "bool",
}

value_type_names2 = {}

def detect_common_type(lst):
    se = set(type(i) for i in lst)
    assert len(se) == 1, lst
    return value_type_names[se.pop()]

def value_type_conv(vs, col=None):
    lst = []
    is_dict = False
    values = []
    for v in vs:
        if isinstance(v, list):
            lst.extend(v)
        elif isinstance(v, dict):
            lst.extend(v.keys())
            values.extend(v.values())
            is_dict = True
            for k in v: v[k] = float(v[k]) # monkey patch, 20140116
        else:
            return value_type_names[type(v)]

    if is_dict:
        #if values:
        #    t = detect_common_type(values)
        #else:
        #    t = value_type_names[float]  # assume is float
        #if col is not None:
        #    value_type_names2[col] = t
        t = "float" # monkey patch, 20140116
        return "Dictionary<object, {}>".format(t)

    t = detect_common_type(lst)
    if col is not None:
        value_type_names2[col] = t
    return "{}[]".format(t)


def value_conv(v, col=None):
    if isinstance(v, list):
        out = "new {}[]{{{}}}".format(value_type_names2[col],
                                      ",".join(value_conv(i) for i in v))
    elif isinstance(v, dict):
        out = "new Dictionary<object, {}>{{{}}}".format(
            #value_type_names2[col],
            "float", # monkey patch, 20140116
            #",".join("{{{},{}}}".format(value_conv(k), value_conv(v))
            ",".join("{{{},{}f}}".format(value_conv(k), float(v)) # monkey patch, 20140116
                     for k, v in sorted(v.items())))
    elif isinstance(v, str):
        out = json.dumps(v)
    elif isinstance(v, bool):
        out = str(v).lower()
    elif isinstance(v, float):
        out = "{}f".format(v)
    else:
        out = str(v)
    return out

if __name__ == "__main__":
    root = collections.OrderedDict()
    idx_key_map = {
        #"monsters": "id",
        #"skills_result": "id",
        #"skills_display": "id",
    }

    with open("game_info.tpl") as f:
        t = template.Template(f.read())

    for i in sorted(glob.glob("tmp/*")):
        k = i.partition("/")[2]
        with open(i) as f:
            root[k] = [collections.OrderedDict(sorted(_.items()))
                       for _ in json.load(f)]

    s = t.generate(
        value_type_conv=value_type_conv,
        value_conv=value_conv,
        idx_key_map=idx_key_map,
        root=root).decode()
    s = re.sub(r"\s+\n", "\n", s)

    with open("../s/GameConfig.cs", "w") as f:
        f.write(s)
