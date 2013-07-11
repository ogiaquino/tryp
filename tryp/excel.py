import pandas as pd
from xlwt import Workbook


def to_excel(tryp):
    sheetname = tryp.excel['sheetname']
    filename = tryp.excel['filename']
    wb = Workbook()
    ws = wb.add_sheet(sheetname)
    write_index(ws, tryp)
    write_columns(ws, tryp)
    write_values_labels(ws, tryp)
    write_values(ws, tryp)
    wb.save(filename)


def _merge_labels(indexes, index_width, total_width):
    labels = {}
    for ir in range(index_width):
        labels[ir] = []
        lseries = pd.Series(map(lambda x: x[ir], indexes))
        if ir <= total_width:
            lseries = lseries.drop_duplicates()

        for il, idx in enumerate(lseries.index):
            label = lseries[idx].decode("utf-8")
            if il == len(lseries.index) - 1:
                labels[ir].append((idx, len(indexes) - 1, label))
            else:
                labels[ir].append((idx, lseries.index[il + 1] - 1,
                                   label))
    return labels


def write_index(ws, tryp):
    indexes = tryp.crosstab.index
    index_width = len(tryp.rows)
    total_width = len(tryp.rows_totals)
    labels = _merge_labels(indexes, index_width, total_width)

    for k in sorted(labels.keys()):
        for label in labels[k]:
            r1 = label[0]
            r2 = label[1]
            c1 = k
            c2 = k
            label = label[2]
            ws.write_merge(r1, r2, c1, c2, label)


def write_columns(ws, tryp):
    rows = tryp.rows
    indexes = tryp.crosstab.columns
    index_width = len(tryp.columns)
    total_width = len(tryp.columns_totals)
    labels = _merge_labels(indexes, index_width, total_width)

    for k in sorted(labels.keys()):
        for label in labels[k]:
            r1 = k
            r2 = k
            c1 = label[0] + len(rows)
            c2 = label[1] + len(rows)
            label = label[2]
            ws.write_merge(r1, r2, c1, c2, label)


def write_values_labels(ws, tryp):
    rows = tryp.rows
    columns = tryp.columns
    labels = tryp.labels
    crosstab = tryp.crosstab

    for i, cc in enumerate(map(lambda x: x[-1], crosstab.columns)):
        r = len(columns)
        c = len(rows) + i
        label = cc
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
