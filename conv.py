#!/usr/bin/env python3
# xls

import collections
import configparser
import functools
import itertools
import json
import glob
import logging
import os
import pprint
import re
import shelve
import sys
import time
import traceback

from tornado import template

value_type_names = {
    int: "int",
    str: "string",
    float: "double",
    bool: "bool",
}

def value_type_conv(vs):
    lst = []
    values = []
    for v in vs:
        if isinstance(v, list):
            lst.extend(v)
        elif isinstance(v, dict):
            lst.extend(v.keys())
            values.extend(v.values())
        else:
            return value_type_names[type(v)]

    se = set(type(i) for i in lst)
    if len(se) != 1:
        logging.warning(lst)
        return "object[]"

    if values:
        se = set(type(i) for i in values)
        if len(se) != 1:
            logging.warning(values)
            return "xxx"
        return "Dictionary<object, " + value_type_names[se.pop()] + ">"

    return value_type_names[se.pop()] + "[]"

def value_conv(v, t=None):
    out = str(v)
    if isinstance(v, list):
        out = "new int[]{" + ",".join(str(i) for i in v) + "}" #todo
    elif isinstance(v, dict):
        out = "new Dictionary<object, int>{" + \
                ",".join("{%s,%s}" % (value_conv(k), value_conv(v))
                         for k, v in v.items()) + \
                "}" #todo
    elif isinstance(v, (str, bool)):
        out = json.dumps(v, ensure_ascii=False)
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
