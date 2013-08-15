#!/usr/bin/env python
# xls

import collections
import functools
import itertools
import json
import re
import shelve
import sys

import xlrd

from pprint import pprint
pprint = functools.partial(pprint, width=40)

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

def main(xls_file):
    for k, v in get_values(xls_file).items():
        print("%s:" % k, dump_raw(v))

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


strip = lambda s: s.strip()

try_int = lambda s: int(s) if s.isdigit() else s

def quoted(s):
    """quote a string, unless this thing could be a number"""
    try:
        float(s)
    except ValueError:
        s = repr(s)
    return s

def bb_list(raw):
    """accept ONLY spliter comma(",")
    """
    if not isinstance(raw, str):
        raw = str(raw)
    values = map(quoted, map(strip, raw.split(",")))
    return eval("[%s]" % ', '.join(values))

def bb_key_value(raw):
    k, v = map(quoted, map(strip, raw.split(":")))
    return "%s: %s" % (k, v)

def bb_dict(raw):
    """accept ONLY spliter CR("\n")
    """
    values = map(bb_key_value, raw.split("\n"))
    return eval("{%s}" % ', '.join(values))

def bb_time(raw):
    """use value returned by xldate_as_tuple() directly"""
    assert isinstance(raw, tuple), raw
    assert len(raw) == 6, raw
    return raw

rc_key = re.compile(r"([a-z]+)(\d*)")

def bb_reward(raw):
    """see bb.xls"""
    units = map(strip, raw.split())
    rws = []   # for all rewards
    w0 = []    # for weight rewards, one `weight loop`, reward
    w1 = []    # for weight rewards, one `weight loop`, weight

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
        rw.append(int(i3) if i3.isdigit() else i3)

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
                rws.append([w0[:], w1[:]])
                del w0[:], w1[:]
            rws.append(rw)

    if w0 and w1:   # flush tail
        rws.append([w0[:], w1[:]])
        del w0[:], w1[:]

    return rws


def bb_require(raw):
    """see bb.xls"""
    rq = raw.split(":")
    l = len(rq)
    assert rq[0][0].isalpha(), "invalid `%s`" % (rq[0],)
    assert l == 2 or l == 3, "length must be 2 or 3" % (rq,)
    n = rq[1].strip()
    if n.isdigit():   # L:N[:R]
        rq[1] = int(n)
        if l == 3:
            compile(rq[2], "", "eval")   # try to compile it
    else:   # L:E[:N]
        compile(n, "", "eval")
        if l == 3:
            rq[2] = int(rq[2])
    return rq


bb_types = {
    "list": bb_list,
    "dict": bb_dict,
    "time": bb_time,
    "reward": bb_reward,
    "require": bb_require,
}

def list_to_tuple(v):
    if isinstance(v, (list, tuple)):
        v = tuple(list_to_tuple(i) for i in v)
    return v


def note_text_to_attr(text):
    """type, test, ...
    type: int, float, str, bool, list, time, reward, require, ...
    test: any python statement
    ...
    """
    attr = {}
    for line in filter(None, map(strip, text.split("\n"))):
        token = line.split(":", 1)
        if len(token) != 2:
            continue
        k, v = map(strip, token)
        if k == "type":
            attr["type"] = eval(v, None, bb_types)
        elif k == "test":
            attr["test"] = compile(v, v, "eval")

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
            attrs.append(None)

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
        progress["column"] = xlrd.colname(colx)
        colx += 1

        if attr:
            #
            _type = attr.get("type")
            if _type:
                try:
                    x = _type(x)
                except Exception as e:
                    progress["error"] = "type error: %s, %s" \
                                        % (type(e).__name__, e)
                    break
            #
            _test = attr.get("test")
            if _test:
                try:
                    if not eval(_test):
                        progress["error"] = "test %r faild: x is %r" \
                                            % (_test.co_filename, x)
                        break
                except Exception as e:
                    progress["error"] = "test error: type error: %s, %s" \
                                        % (type(e).__name__, e)
                    break
            #
        o.append(x)
    else:   # all is well
        return o


def combine(keys, attrs, rows_values, custom_attrs):
    """
    pprint(keys)
    pprint(attrs)
    pprint(rows_values)
    """
    o = []
    for rowx, values in enumerate(rows_values, 1):
        progress["row"] = rowx
        v = apply_attrs(values, attrs, custom_attrs, rowx)
        if not v:
            pprint(progress)
            break
        od = collections.OrderedDict(zip(keys, v))
        o.append(od)
    else:
        return o

def filter_cell(type, value, datemode=0):
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
    rows_values = [list(map(filter_cell, sheet.row_types(i), sheet.row_values(i)))
                   for i in range(1, sheet.nrows)]
    return combine(keys, attrs, rows_values, custom_attrs)

def parse(xls, db):
    workbook = xlrd.open_workbook(xls)
    progress.clear()
    progress["xls"] = xls
    for sheet in workbook.sheets():
        cell_note_map = sheet.cell_note_map
        head_note = cell_note_map.get((0, 0))
        if head_note:
            progress["sheet"] = sheet.name
            title = head_note.author
            if title:
                progress["title"] = title
                s = parse_sheet(sheet)
                pprint(s)
                if s:
                    db[title] = s
                else:
                    print("beiju:")
                    pprint(progress)
                    continue



if __name__ == "__main__":
    sys.argv.append("b.xls")
    #main(sys.argv[1])
    db = shelve.open("bb")
    parse(sys.argv[1], db)
    db.close()
