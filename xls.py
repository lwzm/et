#!/usr/bin/env python3
# xls

import collections
import configparser
import functools
import itertools
import json
import logging
import os
import pprint
import re
import shelve
import sys
import time
import traceback

import xlrd


class Default(dict):
    def __missing__(self, key):
        return key

eval_namespace = Default()

S = eval_namespace.__setitem__  # for eval in `test`


dump_sorted = functools.partial(json.dumps, ensure_ascii=False,
                                separators = (",", ": "),
                                sort_keys=True, indent=4)

dump_raw = functools.partial(json.dumps, ensure_ascii=False,
                             separators = (",", ": "), indent=4)


# global progress status
progress = collections.OrderedDict.fromkeys(["xls", "sheet", "column", "row"])


class ListDefaultDict(collections.OrderedDict):
    def __missing__(self, k):
        self[k] = []
        return self[k]

xls_tasks = ListDefaultDict()
uniq_tasks = ListDefaultDict()
sort_tasks = ListDefaultDict()
ref_file_tasks = ListDefaultDict()
json_outputs = ListDefaultDict()


workbooks_mtimes = collections.Counter()

log_level = logging.DEBUG
log_level = logging.INFO

logging.basicConfig(
    level=log_level,
    format="%(levelname)s:%(message)s",
    )


def check(prefix=""):
    """check
    uniq_tasks
    sort_tasks
    ref_file_tasks
    """
    logging.info("check uniq_tasks:")
    for pos, lst in uniq_tasks.items():
        logging.debug((pos, lst))
        if len(frozenset(lst)) != len(lst):
            logging.warning(pos)

    logging.info("check sort_tasks:")
    for pos, lst in sort_tasks.items():
        logging.debug((pos, lst))
        if sorted(lst) != list(lst):
            logging.warning(pos)

    logging.info("check ref_file_tasks:")
    for folder, files in ref_file_tasks.items():
        fs = set(os.listdir(os.path.join(prefix, folder)))
        for f, pos in files:
            logging.debug((pos, f, folder))
            if f not in fs:
                logging.warning(pos)

cellname_match = re.compile(r"([A-Z]+)(\d+)").match

def cellname_to_index(name, _A2Z="ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
    """
    >>> cellname_to_index("A1")
    (0, 0)
    >>> cellname_to_index("Z1")
    (25, 0)
    >>> cellname_to_index("ZZ1")
    (701, 0)
    >>> cellname_to_index("AAA111")
    (702, 110)
    """
    col, row = cellname_match(name).groups()
    base = len(_A2Z)
    n = m = 0
    for i in map(_A2Z.index, col):
        n = m + i
        m = (n + 1) * base
    return n, int(row) - 1

    

def bb_time(raw):
    """use value returned by xldate_as_tuple() directly"""
    assert isinstance(raw, tuple), raw
    assert len(raw) == 6, raw
    return raw

rc_key_match = re.compile(r"([a-z]+)(\d*)").match


def bb_mess(raw):
    """see bb.xls"""
    mess = []   # all
    w0 = []    # for weight things, one `weight loop`, thing
    w1 = []    # for weight things, one `weight loop`, weight

    def flush():
        if w0 and w1:
            mess.append([w0[:], w1[:]])
            del w0[:], w1[:]

    for unit in filter(None, (_.strip() for _ in raw.split("\n"))):
        keys = ([_.strip() for _ in unit.split("#")] + ["100%"])[:3]
        assert len(keys) == 3, unit

        term = list(eval(keys[0] + ",", None, eval_namespace))
        #assert term[0] in {"item", "gold"}, term[0]
        n = keys[1]

        if n.isdigit():
            n = int(n)
        else:
            compile(n, "just test", "eval")
        term.append(n)

        rate = keys[2]

        if rate[-1] == "%":
            mess.append([term, float(rate[:-1]) / 100])
            flush()
        else:
            w0.append(term)
            w1.append(float(rate))

    flush()  # flush tail

    return mess


list_fmt = "[{}]".format
dict_fmt = "{{{}}}".format

def bb_req(raw):
    """see bb.xls"""
    req = list(_.strip() for _ in raw.split("#"))
    l = len(req)
    assert l == 2 or l == 3, req
    x = req[1].strip()
    if x.isdigit():   # L:N[:R]
        if l == 3:
            compile(req[2], "just test", "eval")
            req[1], req[2] = req[2], int(x)
        else:
            req.insert(1, "True")
            req[2] = int(req[2])
    else:   # L:E[:N]
        compile(x, "just test", "eval")
        req[0] = eval(list_fmt(req[0]), None, eval_namespace)  # multi-log-activated
        if l == 3:
            req[2] = int(req[2])
        else:
            req.append(1)  # for plan status: (0/1) or (1/1)
    return req


bb_types = {
    # my custom formated list and dict
    "list": lambda raw: eval(list_fmt(raw), None, eval_namespace),
    "dict": lambda raw: eval(dict_fmt(raw), None, eval_namespace),
    "time": bb_time,
    "mess": bb_mess,
    "req": bb_req,
}


def note_text_to_attr(text):
    """type and test and ...
    type: int, float, str, bool, list, dict, time, mess, req, ...
    test: any python statement
    ...
    """
    attr = {}
    for line in filter(None, (_.strip() for _ in text.split("\n"))):
        token = line.split(":", 1)
        if len(token) != 2:
            continue
        k, v = (_.strip() for _ in token)
        if k == "type":
            attr["type"] = eval(v, None, bb_types)
        elif k == "test":
            attr["test"] = v
        elif k == "ref":
            attr["ref"] = v
        elif k == "uniq":
            attr["uniq"] = True
        elif k == "sort":
            attr["sort"] = True
    return attr

def get_custom_attrs(sheet):
    """cells' custom attrubutes
    range will be used: row 2 - end
    """
    cell_note_map = sheet.cell_note_map
    o = {}
    for k, v in cell_note_map.items():
        if k[0] == 0:  # ignore row 1, already parsed in get_keys_attrs
            continue
        txt = v.text
        out = note_text_to_attr(txt)
        logging.debug(
            "note_text_to_attr_{}({!r}) = {}".format(xlrd.cellname(*k), txt, out))
        o[k] = out
    return o


def get_keys_attrs(sheet):
    """keys and attrs
    keys come from row 1, a normal string is ok
    attrs live with these keys, note text
    """
    progress["row"] = 1
    keys = list(itertools.takewhile(lambda x: isinstance(x, str) and x,
                                    sheet.row_values(0)))
    assert keys, progress
    assert len(set(keys)) == len(keys), keys

    attrs = []
    cell_note_map = sheet.cell_note_map
    for colx in range(len(keys)):
        colname = xlrd.colname(colx)
        progress["column"] = colname
        note = cell_note_map.get((0, colx))
        if note:
            txt = note.text
            out = note_text_to_attr(txt)
            logging.debug(
                "note_text_to_attr_{}({!r}) = {}".format(colname, txt, out))
            attrs.append(out)
        else:
            attrs.append({})
    return keys, attrs


def apply_attrs(values, attrs, custom_attrs, rowx):
    """convert and check cell.value
    if error occurs, set progress["error"] and break
    """
    fmt = "{} -> {} -> {}".format
    o = []
    colx = 0
    for x, attr in zip(values, attrs):
        custom_attr = custom_attrs.get((rowx, colx))
        if custom_attr:
            attr = attr.copy()
            attr.update(custom_attr)

        colname = xlrd.colname(colx)
        progress["column"] = colname
        abs_colname = fmt(progress["xls"], progress["sheet"], colname)

        if attr:
            #
            _type = attr.get("type")
            if _type:
                x = _type(x)
            #
            _test = attr.get("test")
            if _test:
                assert eval(_test, None, locals()), _test  # `S` and `o` can be used here
            #
            if attr.get("uniq"):
                uniq_tasks[abs_colname].append(x)
            #
            if attr.get("sort"):
                sort_tasks[abs_colname].append(x)
            #
            _ref = attr.get("ref")
            if _ref:
                abs_cellname = fmt(progress["xls"], progress["sheet"],
                                   xlrd.cellname(rowx, colx))
                ref_file_tasks[_ref].append([x, abs_cellname])

        o.append(x)
        colx += 1

    return o


def combine(keys, attrs, rows_values, custom_attrs):
    """
    pprint.pprint(keys)
    pprint.pprint(attrs)
    pprint.pprint(rows_values)
    """
    o = []
    for rowx, values in enumerate(rows_values, 1):
        progress["row"] = rowx + 1   # human readable
        v = apply_attrs(values, attrs, custom_attrs, rowx)
        #od = collections.OrderedDict(zip(keys, v))
        o.append(dict(zip(keys, v)))
    return o


def filter_cell_value(type, value, datemode=0):
    if type == xlrd.XL_CELL_NUMBER:
        if value.is_integer():
            value = int(value)
    elif type == xlrd.XL_CELL_DATE:
        value = xlrd.xldate_as_tuple(value, datemode)
    elif type == xlrd.XL_CELL_BOOLEAN:
        value = bool(value)
    else:
        value = value.strip()
    return value


def parse_sheet(sheet):
    keys, attrs = get_keys_attrs(sheet)
    custom_attrs = get_custom_attrs(sheet)
    #pprint.pprint(keys)
    #pprint.pprint(attrs)
    #pprint.pprint(custom_attrs)
    rows_values = []
    for i in range(1, sheet.nrows):
        if sheet.row_values(i)[0] != "":
            rows_values.append(list(map(filter_cell_value,
                                        sheet.row_types(i),
                                        sheet.row_values(i))))
        else:
            break
    return combine(keys, attrs, rows_values, custom_attrs)


def walk(directory, dir_filter=None, file_filter=None):
    """return a list of all files in this directory, sorted"""
    all_files = []
    for root, dirs, files in os.walk(directory):
        dirs[:] = sorted(filter(dir_filter, dirs))
        all_files.extend(map(lambda f: os.path.abspath(os.path.join(root, f)),
                         sorted(filter(file_filter, files))))
    return all_files


def main():
    cfg = configparser.ConfigParser()
    cfg.read("idx")

    for i in cfg.sections():
        for j in cfg[i]:
            xls_tasks[cfg[i][j]].append([i, j])

    #pprint.pprint(xls_tasks)

    def parse(xls, sheet):
        progress["xls"] = xls
        progress["sheet"] = sheet
        sheet = xlrd.open_workbook(xls).sheet_by_name(sheet)
        return parse_sheet(sheet)

    for k, v in xls_tasks.items():
        for xls, sheet in v:
            logging.info("parse_sheet  {} - {}  to {}".format(xls, sheet, k))
            json_outputs[k].extend(parse(os.path.join("files", xls), sheet))

    for k, v in json_outputs.items():
        with open(os.path.join("jsons", k), "w") as f:
            f.write(dump_sorted(v))

    check()
    #pprint.pprint(names)
    #pprint.pprint(_eval_cache)

if __name__ == "__main__":
    try:
        main()
    except Exception:
        fmt = "{:>10} -> {}".format
        for k, v in progress.items():
            print(fmt(k, v))
        raise
