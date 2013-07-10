import itertools
import pandas as pd
from xlwt import Workbook


def to_excel(tryp):
    sheetname = tryp.excel['sheetname']
    filename = tryp.excel['filename']
    wb = Workbook()
    ws = wb.add_sheet(sheetname)
    write_index(ws, tryp)
    write_columns_labels(ws, tryp)
    write_values_labels(ws, tryp)
    write_values(ws, tryp)
    wb.save(filename)


def _labels(indexes, index_width, total_width, axis=0):
    labels = []
    for ir in range(index_width):
        lseries = pd.Series(map(lambda x: x[ir], indexes))
        if ir <= total_width:
            lseries = lseries.drop_duplicates()

        for il, idx in enumerate(lseries.index):
            label = lseries[idx].decode("utf-8")
            if il == len(lseries.index) - 1:
                if axis == 0:
                    labels.append((idx, len(indexes) - 1, ir, ir, label))
                if axis == 1:
                    labels.append((ir, ir, idx, len(indexes) - 1, label))
            else:
                if axis == 0:
                    labels.append((idx, lseries.index[il + 1] - 1, ir, ir,
                                   label))
                if axis == 1:
                    labels.append((ir, ir, idx, lseries.index[il + 1] - 1,
                                   label))
    return labels


def write_index(ws, tryp):
    indexes = tryp.crosstab.index
    index_width = len(tryp.rows)
    total_width = len(tryp.rows_totals)
    for label in _labels(indexes, index_width, total_width, 0):
        ws.write_merge(*label)


def write_columns_labels(ws, tryp):
    rows = tryp.rows
    indexes = tryp.crosstab.columns
    index_width = len(tryp.columns)
    total_width = len(tryp.columns_totals)
    for label in _labels(indexes, index_width, total_width, 1):
        ws.write_merge(label[0],
                       label[1],
                       label[2] + len(rows),
                       label[3] + len(rows),
                       label[4])


def write_values_labels(ws, tryp):
    rows = tryp.rows
    columns = tryp.columns
    labels = tryp.labels
    crosstab = tryp.crosstab

    for i, cc in enumerate(crosstab.columns):
        r = len(columns)
        c = len(rows) + i
        if isinstance(cc,  basestring):
            label = cc
        else:
            label = cc[-1]
        if labels:
            label = labels['values'][label]
        label = str(label).decode("utf-8")
        ws.write(r, c, label)


def write_values(ws, tryp):
    rows = tryp.rows
    columns = tryp.columns
    crosstab = tryp.crosstab

    for iv, value in enumerate(crosstab.values):
        for il, label in enumerate(value):
            r = iv + len(columns) + 1
            c = il + len(rows)
            ws.write(r, c, label)
