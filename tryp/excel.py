import itertools
from xlwt import Workbook


def to_excel(tryp):
    sheetname = tryp.excel['sheetname']
    filename = tryp.excel['filename']
    wb = Workbook()
    ws = wb.add_sheet(sheetname)
    write_rows_labels(ws, tryp)
    write_columns_labels(ws, tryp)
    write_values_labels(ws, tryp)
    write_values(ws, tryp)
    wb.save(filename)


def write_rows_labels(ws, tryp):
    rows = tryp.rows
    columns = tryp.columns
    crosstab = tryp.crosstab
    labels_rows = []

    for i in range(len(rows)):
        sn = i
        ci = [x[i] for x in crosstab.index]
        if i < len(rows)-1:
            label_group = [list(g) for k, g in itertools.groupby(ci)]
        else:
            label_group = [[x] for x in ci]
        count = -1
        index = []
        for group in label_group:
            index.append([])
            for g in group:
                count = count + 1
                index[-1].append(count)

        for x in index:
            r1 = x[0] + len(columns) + 1
            r2 = x[-1] + len(columns) + 1
            c1 = i
            c2 = i
            label = ci[x[0]]
            label = str(label).decode("utf-8")
            ws.write_merge(r1, r2, c1, c2, label)


def write_columns_labels(ws, tryp):
    rows = tryp.rows
    columns = tryp.columns
    crosstab = tryp.crosstab

    for i in range(len(columns)):
        sn = i
        cc = [x[i] for x in crosstab.columns]
        label_group = [list(g) for k, g in itertools.groupby(cc)]
        count = -1
        index = []
        for group in label_group:
            index.append([])
            for g in group:
                count = count + 1
                index[-1].append(count)
        for x in index:
            r1 = i
            r2 = i
            c1 = x[0] + len(rows)
            c2 = x[-1] + len(rows)
            label = cc[x[0]]
            label = str(label).decode("utf-8")
            ws.write_merge(r1, r2, c1, c2, label)


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
