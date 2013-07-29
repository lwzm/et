#!/usr/bin/env python
# xls

import collections
import functools
import json
import sys

import xlrd

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



if __name__ == "__main__":
    sys.argv.append("b.xls")
    main(sys.argv[1])
