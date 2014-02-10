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
        else:
            return value_type_names[type(v)]

    if is_dict:
        return "Dictionary<object, BuffVO>"

    t = detect_common_type(lst)
    if col is not None:
        value_type_names2[col] = t
    return "{}[]".format(t)


def to_BuffVO(v):
    if not isinstance(v, list):
        v = [v]
    return "new BuffVO({})".format(",".join(map(value_conv, map(float, v))))

def value_conv(v, col=None):
    if isinstance(v, list):
        out = "new {}[]{{{}}}".format(value_type_names2[col],
                                      ",".join(value_conv(i) for i in v))
    elif isinstance(v, dict):
        out = "new Dictionary<object, BuffVO>{{{}}}".format(
            ",".join("{{{},{}}}".format(value_conv(k), to_BuffVO(v)) # monkey patch, 20140116
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
        "heroes": "id_lv",
        #"monsters": "id",
        #"skills_result": "id",
        #"skills_display": "id",
    }

    ignores = {
        "hero_strengthen",
        "hero_strengthen_gold_cost_1",
        "hero_strengthen_gold_cost_2",
        "hero_strengthen_gold_cost_3",
        "hero_strengthen_gold_cost_4",
    }

    simples = {
        #"hero_strengthen_1": ("level", "gold_cost"),
        #"hero_strengthen_2": ("level", "gold_cost"),
        #"hero_strengthen_3": ("level", "gold_cost"),
        #"hero_strengthen_4": ("level", "gold_cost"),
    }

    client_file = "GameConfig.cs"
    with open(client_file) as f:
        t = template.Template(f.read())

    for i in sorted(glob.glob("tmp/*")):
        k = i.partition("/")[2]
        if k not in ignores and k not in simples:
            with open(i) as f:
                root[k] = [collections.OrderedDict(sorted(_.items()))
                           for _ in json.load(f)]
                if k == "heroes":   # patch heroes
                    for i in root[k]:
                        i["id_lv"] = i["id"] * 1000 + i["lv"]

    s = t.generate(
        value_type_conv=value_type_conv,
        value_conv=value_conv,
        idx_key_map=idx_key_map,
        root=root).decode()
    s = re.sub(r"\s+\n", "\n", s)

    with open("../s/" + client_file, "w") as f:
        f.write(s)


    from tornado.util import ObjectDict
    root = ObjectDict()
    for i in sorted(glob.glob("tmp/*")):
        k = i.partition("/")[2]
        with open(i) as f:
            root[k] = [ObjectDict(_) for _ in json.load(f)]

    def export_b_config(name):
        with open(name) as f:
            t = template.Template(f.read())
        s = t.generate(root=root).decode()
        s = re.sub(r"\s+\n", "\n", s)
        with open("../s/srv/" + name, "w") as f:
            f.write(s)

    export_b_config("src/b_config.erl")
    export_b_config("src/b_proto.erl")
    export_b_config("src/message_code.erl")
    export_b_config("include/b_config.hrl")
