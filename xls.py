#!/usr/bin/env python
# xls

import collections
import functools
import itertools
import json
import shelve
import sys

import xlrd

from pprint import pprint

dump_sorted = functools.partial(json.dumps, ensure_ascii=False,
                                separators = (",", ": "),
                                sort_keys=True, indent=4)

dump_raw = functools.partial(json.dumps, ensure_ascii=False,
                             separators = (",", ": "), indent=4)


# global progress status
progress = {
    "xls": None,
    "sheet": None,
    "row": None,
    "column": None,
    "title": None,
    "error": None,
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

# quote a string, unless this thing could be a number
quoted = lambda s: s if s.isdigit() else repr(s)
# or:
def quoted(s):
    """accept float format as a `number`"""
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

bb_types = {
    "list": bb_list,
    "dict": bb_dict,
}


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
    expr = "x != 212"
    o = []
    colx = 0
    #print(custom_attrs)
    for x, attr in zip(values, attrs):
        custom_attr = custom_attrs.get((rowx, colx))
        #print(custom_attr)
        if custom_attr:
            attr = custom_attr
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
            if _test and not eval(_test):
                progress["error"] = "test %r faild: x is %r" \
                                    % (_test.co_filename, x)
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
                    break



if __name__ == "__main__":
    sys.argv.append("b.xls")
    #main(sys.argv[1])
    db = shelve.open("bb")
    parse(sys.argv[1], db)
    db.close()
