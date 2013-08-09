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

cell_type = {
    xlrd.XL_CELL_TEXT: str,
    xlrd.XL_CELL_NUMBER: float,
    xlrd.XL_CELL_DATE: float,
}


def note_text_to_attrs(s):
    """TODO"""
    return s

def get_keys_attrs(sheet):
    """keys and attrs
    keys come from row 1, a normal string is ok
    attrs live with these keys, note text
    """
    keys = list(itertools.takewhile(lambda x: isinstance(x, str) and x,
                                    sheet.row_values(0)))
    assert keys, keys

    attrs = []
    cell_note_map = sheet.cell_note_map
    for i in range(len(keys)):
        note = cell_note_map.get((0, i))
        if note:
            attrs.append(note_text_to_attrs(note.text)) #dummy
        else:
            attrs.append(None)

    return keys, attrs

def apply_attrs(values, attrs):
    """TODO"""
    out = []
    for x, attr in zip(values, attrs):
        attr = {"type": int, "test": compile("x < 30", "test", "eval")} #dummy
        type = attr.get("type")
        if type:
            x = type(x)
        test = attr.get("test")
        if test and not eval(test):
            print("err")
        out.append(x)
    return out


def combine(keys, attrs, values):
    pprint(keys)
    pprint(attrs)
    pprint(values)
    rows = []
    for raw in values:
        v = apply_attrs(raw, attrs)
        row = collections.OrderedDict(zip(keys, v))
        rows.append(row)
    pprint(rows)

def filter_int(x):
    if isinstance(x, float) and x.is_integer():
        return int(x)
    else:
        return x

def parse_sheet(sheet):
    keys, attrs = get_keys_attrs(sheet)
    values = [list(map(filter_int, sheet.row_values(i)))
              for i in range(1, sheet.nrows)]
    combine(keys, attrs, values)

def parse(xls, db):
    wb = xlrd.open_workbook(xls)
    for s in wb.sheets():
        cell_note_map = s.cell_note_map
        head_note = cell_note_map.get((0, 0))
        if head_note:
            title = head_note.author
            if title:
                db[title] = parse_sheet(s)



if __name__ == "__main__":
    sys.argv.append("b.xls")
    #main(sys.argv[1])
    db = shelve.open("bb")
    parse(sys.argv[1], db)
    db.close()
