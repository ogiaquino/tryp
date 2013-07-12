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


def merge_labels(indexes, index_width, total_width):
    labels = {}

    def __labels(k, series):
        labels[k] = []
        lseries = series.index.tolist()
        lseries.append(len(indexes))
        for il, idx in enumerate(lseries[:-1]):
            labels[k].append((idx, lseries[il + 1] - 1, series[idx]))

    for ir in range(total_width):
        series = pd.Series(zip(*indexes)[ir])
        series = series.drop_duplicates()
        __labels(ir, series)

    for ir in range(total_width, index_width):
        series = pd.Series(zip(*indexes)[ir])
        __labels(ir, series)

    return labels


def write_index(ws, tryp):
    columns = tryp.columns
    indexes = tryp.crosstab.index
    index_width = len(tryp.rows)
    total_width = len(tryp.rows_totals)
    labels = merge_labels(indexes, index_width, total_width)

    for k in sorted(labels.keys()):
        for label in labels[k]:
            r1 = label[0] + len(columns) + 1
            r2 = label[1] + len(columns) + 1
            c1 = k
            c2 = k
            label = label[2]
            ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"))


def write_columns(ws, tryp):
    rows = tryp.rows
    indexes = tryp.crosstab.columns
    index_width = len(tryp.columns)
    total_width = len(tryp.columns_totals)
    labels = merge_labels(indexes, index_width, total_width)

    for k in sorted(labels.keys()):
        for label in labels[k]:
            r1 = k
            r2 = k
            c1 = label[0] + len(rows)
            c2 = label[1] + len(rows)
            label = label[2]
            ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"))


def write_values_labels(ws, tryp):
    rows = tryp.rows
    columns = tryp.columns
    labels = tryp.labels
    values_labels = tryp.crosstab.values_labels

    for i, cc in enumerate(values_labels):
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
