#!/usr/bin/env python3
# xls

import collections
import configparser
import functools
import itertools
import json
import os
import re
import shelve
import sys
import time
import traceback

import xlrd

from pprint import pprint
#pprint = functools.partial(pprint, width=40)

class Dict(dict):
    def __missing__(self, k):
        return k

eval_env = Dict()

dump_sorted = functools.partial(json.dumps, ensure_ascii=False,
                                separators = (",", ": "),
                                sort_keys=True, indent=4)

dump_raw = functools.partial(json.dumps, ensure_ascii=False,
                             separators = (",", ": "), indent=4)


# global progress status
progress = {
    #"xls": None,
    #"sheet": None,
    #"row": None,
    #"column": None,
    #"title": None,
    #"error": None,
}

xls_tasks = collections.defaultdict(list)
uniq_tasks = collections.defaultdict(list)
sort_tasks = collections.defaultdict(list)
ref_file_tasks = collections.defaultdict(list)
ref_cell_tasks = collections.defaultdict(list)
json_outputs = collections.defaultdict(list)

workbooks_mtimes = collections.Counter()


def view(xls_file):
    for k, v in get_values(xls_file).items():
        print("%s:" % k, dump_raw(v))

    print()

    for k, v in get_notes(xls_file).items():
        print("%s:" % k, dump_sorted(v))

def get_values(xls):
    """for hp"""
    wb = xlrd.open_workbook(xls)
    workbook = collections.OrderedDict()
    for s in wb.sheets():
        cell_note_map = s.cell_note_map
        note = cell_note_map.get((0, 0))
        if note:
            sheet = collections.OrderedDict(title=note.author)
            workbook[s.name] = sheet
            for row in range(s.nrows):
                rows = []
                sheet["row %d" % (row + 1,)] = rows
                for col in range(s.ncols):
                    rows.append(s.cell(row, col).value)
    return workbook

def get_notes(xls):
    """for me"""
    wb = xlrd.open_workbook(xls)
    all_notes = collections.OrderedDict()
    for s in wb.sheets():
        cell_note_map = s.cell_note_map
        head_note = cell_note_map.get((0, 0))
        if head_note and head_note.author:
            notes = {xlrd.cellname(*k): v.text for k, v in cell_note_map.items()}
            notes["@"] = "%s -> %s" % (xls, s.name)
            all_notes[head_note.author] = notes
    return all_notes


def check(prefix=""):
    """check
    uniq_tasks
    sort_tasks
    ref_file_tasks
    ref_cell_tasks
    """
    for pos, lst in uniq_tasks.items():
        if len(frozenset(lst)) != len(lst):
            print("unique is false:", pos)
    for pos, lst in sort_tasks.items():
        if sorted(lst) != list(lst):
            print("sorted is false:", pos)
    for folder, files in ref_file_tasks.items():
        fs = set(os.listdir(os.path.join(prefix, folder)))
        for f, pos in files:
            if f not in fs:
                print("%s: %s is not in %s" % (pos, f, folder))

def bb_time(raw):
    """use value returned by xldate_as_tuple() directly"""
    assert isinstance(raw, tuple), raw
    assert len(raw) == 6, raw
    return raw

rc_key = re.compile(r"([a-z]+)(\d*)")

def bb_mess(raw):
    """see bb.xls"""
    units = raw.split()
    mess = []   # all
    w0 = []    # for weight things, one `weight loop`, thing
    w1 = []    # for weight things, one `weight loop`, weight

    for i in units:
        rw = []
        flush = True
        keys = i.split(":")
        assert keys[0][0].isalpha(), keys[0]

        i1, i2 = rc_key.match(keys[0]).groups()
        i3 = keys[1]

        rw.append(i1)
        if i2:
            rw.append(int(i2))
        if i3.isdigit():
            i3 = int(i3)
        else:
            compile(i3, "just test", "eval")
        rw.append(i3)

        if len(keys) > 2:
            i4 = keys[2]
            if i4[-1] == "%":
                rw = [rw, float(i4[:-1]) / 100]
            else:
                w0.append(rw)
                w1.append(float(i4) if "." in i4 else int(i4))
                flush = False
        if flush:
            if w0 and w1:
                mess.append([w0[:], w1[:]])
                del w0[:], w1[:]
            mess.append(rw)

    if w0 and w1:   # flush tail
        mess.append([w0[:], w1[:]])
        del w0[:], w1[:]

    return mess


def bb_req(raw):
    """see bb.xls"""
    req = raw.split(":")
    l = len(req)
    assert req[0][0].isalpha(), "invalid `%s`" % (req[0],)
    assert l == 2 or l == 3, "length must be 2 or 3" % (req,)
    n = req[1].strip()
    if n.isdigit():   # L:N[:R]
        req[1] = int(n)
        if l == 3:
            compile(req[2], "just test", "eval")
    else:   # L:E[:N]
        compile(n, "just test", "eval")
    return req


bb_types = {
    "list": lambda raw: eval("[{}]".format(raw), None, eval_env),
    "dict": lambda raw: eval("{" + raw + "}", None, eval_env),
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
            attr["test"] = v#compile(v, v, "eval")
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
        o[k] = note_text_to_attr(v.text)
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

    attrs = []
    cell_note_map = sheet.cell_note_map
    for colx in range(len(keys)):
        progress["column"] = xlrd.colname(colx)
        note = cell_note_map.get((0, colx))
        if note:
            attrs.append(note_text_to_attr(note.text))
        else:
            attrs.append({})
    return keys, attrs


def apply_attrs(values, attrs, custom_attrs, rowx):
    """convert and check cell.value
    if error occurs, set progress["error"] and break
    """
    o = []
    colx = 0
    for x, attr in zip(values, attrs):
        custom_attr = custom_attrs.get((rowx, colx))
        if custom_attr:
            attr = attr.copy()
            attr.update(custom_attr)

        colname = xlrd.colname(colx)
        progress["column"] = colname
        abs_colname = "%s -> %s -> %s" \
                      % (progress["xls"], progress["sheet"], colname)

        if attr:
            #
            _type = attr.get("type")
            if _type:
                try:
                    x = _type(x)
                except Exception:
                    progress["error"] = "ERROR TYPE:\n%s" % traceback.format_exc()
                    break
            #
            _test = attr.get("test")
            if _test:
                try:
                    if not eval(_test):
                        progress["error"] = "test %r faild: x is %r" \
                                            % (_test.co_filename, x)
                        break
                except Exception:
                    progress["error"] = "ERROR TEST:\n%s" % traceback.format_exc()
                    break
            #
            _uniq = attr.get("uniq")
            if _uniq:
                uniq_tasks[abs_colname].append(x)
            #
            _sort = attr.get("sort")
            if _sort:
                sort_tasks[abs_colname].append(x)
            #
            _ref = attr.get("ref")
            if _ref:
                abs_cellname = "%s -> %s -> %s" \
                               % (progress["xls"], progress["sheet"],
                                  xlrd.cellname(rowx, colx))
                if "/" in _ref:
                    ref_file_tasks[_ref].append([x, abs_cellname])
                else:
                    ref_cell_tasks[_ref].append([x, abs_cellname])

        o.append(x)
        colx += 1
    else:   # error is broken(break statement), all is well here
        return o


def combine(keys, attrs, rows_values, custom_attrs):
    """
    pprint(keys)
    pprint(attrs)
    pprint(rows_values)
    """
    o = []
    for rowx, values in enumerate(rows_values, 1):
        progress["row"] = rowx + 1   # human readable
        v = apply_attrs(values, attrs, custom_attrs, rowx)
        if not v:
            pprint(progress)
            break
        #od = collections.OrderedDict(zip(keys, v))
        o.append(dict(zip(keys, v)))
    else:
        return o

def filter_cell_value(type, value, datemode=0):
    if type == xlrd.XL_CELL_NUMBER and value.is_integer():
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
    #print(keys)
    #pprint(attrs)
    #print(custom_attrs)
    rows_values = [
        list(map(filter_cell_value, sheet.row_types(i), sheet.row_values(i)))
        for i in range(1, sheet.nrows)
    ]
    return combine(keys, attrs, rows_values, custom_attrs)


def walk(directory, dir_filter=None, file_filter=None):
    """return a list of all files in this directory, sorted"""
    import os
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

    #pprint(xls_tasks)

    def parse(xls, sheet):
        progress["xls"] = xls
        progress["sheet"] = sheet
        sheet = xlrd.open_workbook(xls).sheet_by_name(sheet)
        return parse_sheet(sheet)

    for k, v in xls_tasks.items():
        for xls, sheet in v:
            json_outputs[k].extend(parse(xls, sheet))

    for k, v in json_outputs.items():
        with open(os.path.join("output", k), "w") as f:
            f.write(dump_sorted(v))

    check("..")

if __name__ == "__main__":
    main()
