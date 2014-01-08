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
    for v in vs:
        if isinstance(v, list):
            if not v:
                continue
            v = v[0]
            return value_type_names[type(v)] + "[]"
        else:
            return value_type_names[type(v)]
    else:
        return value_type_names[int] + []

def value_conv(v, t=None):
    out = str(v)
    if isinstance(v, list):
        out = "new int[]{" + ",".join(str(i) for i in v) + "}" #todo
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
