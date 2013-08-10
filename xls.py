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
    "error": "",
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

cell_type = {
    xlrd.XL_CELL_TEXT: str,
    xlrd.XL_CELL_NUMBER: float,
    xlrd.XL_CELL_DATE: float,
}


def note_text_to_attrs(s):
    """TODO"""
    return s

def get_custom_attrs(sheet):
    cell_note_map = sheet.cell_note_map
    return {}

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
    for i in range(len(keys)):
        progress["column"] = xlrd.colname(i)
        note = cell_note_map.get((0, i))
        if note:
            attrs.append(note_text_to_attrs(note.text)) #dummy
        else:
            attrs.append(None)

    return keys, attrs

def apply_attrs(values, attrs):
    """convert and check cell.value
    if error occurs, set progress["error"] and break
    """
    expr = "x != 22"
    dummy_attr = {"type": int, "test": compile(expr, expr, "eval")}
    o = []
    col = 0
    for x, attr in zip(values, attrs):
        attr = dummy_attr #dummy
        progress["column"] = xlrd.colname(col)
        col += 1
        #
        _type = attr.get("type")
        if _type:
            try:
                x = _type(x)
            except Exception as e:
                progress["error"] = "type %r error: %s, %s" \
                                    % (_type, type(e).__name__, e)
                break
        #
        _test = attr.get("test")
        if _test and not eval(_test):
            progress["error"] = "test %r faild: x is %r" \
                                % ( _test.co_filename, x)
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
    for row, values in enumerate(rows_values, 2):
        progress["row"] = row
        v = apply_attrs(values, attrs)
        if not v:
            pprint(progress)
            break
        od = collections.OrderedDict(zip(keys, v))
        o.append(od)
    else:
        return o

def filter_int(x):
    if isinstance(x, float) and x.is_integer():
        return int(x)
    else:
        return x

def parse_sheet(sheet):
    keys, attrs = get_keys_attrs(sheet)
    custom_attrs = get_custom_attrs(sheet)
    rows_values = [list(map(filter_int, sheet.row_values(i)))
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
