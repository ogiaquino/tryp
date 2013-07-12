import pandas as pd
from xlwt import Workbook


def to_excel(tryp):
    sheetname = tryp.excel['sheetname']
    filename = tryp.excel['filename']
    wb = Workbook()
    tryp.ws = wb.add_sheet(sheetname)
    write_axes(tryp)
    write_values(tryp)
    wb.save(filename)


def write_axes(tryp):
    def _write_axes(idx):
        r1 = idx['r1']
        r2 = idx['r2']
        c1 = idx['c1']
        c2 = idx['c2']
        label = idx['label']
        tryp.ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"))

    for idx in _index(tryp):
        _write_axes(idx)

    for idx in _columns(tryp):
        _write_axes(idx)


def write_values(tryp):
    def _write_values(idx):
        r = idx['r']
        c = idx['c']
        label = idx['label']
        tryp.ws.write(r, c, label)

    for idx in _values_labels(tryp):
        _write_values(idx)

    for idx in _values(tryp):
        _write_values(idx)


def _merge_labels(indexes, index_width, total_width):
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


def _index(tryp):
    columns = tryp.columns
    indexes = tryp.crosstab.index
    index_width = len(tryp.rows)
    total_width = len(tryp.rows_totals)
    labels = _merge_labels(indexes, index_width, total_width)

    for k in sorted(labels.keys()):
        for label in labels[k]:
            r1 = label[0] + len(columns) + 1
            r2 = label[1] + len(columns) + 1
            c1 = k
            c2 = k
            label = label[2]
            yield {'r1': r1, 'r2': r2, 'c1': c1, 'c2': c2, 'label': label}


def _columns(tryp):
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
            yield {'r1': r1, 'r2': r2, 'c1': c1, 'c2': c2, 'label': label}


def _values_labels(tryp):
    rows = tryp.rows
    columns = tryp.columns
    labels = tryp.labels
    values_labels = tryp.crosstab.values_labels

    for i, cc in enumerate(values_labels):
        r = len(columns)
        c = len(rows) + i
        label = cc
        yield {'r': r, 'c': c, 'label': label}


def _values(tryp):
    rows = tryp.rows
    columns = tryp.columns
    crosstab = tryp.crosstab

    for iv, value in enumerate(crosstab.values):
        for il, label in enumerate(value):
            r = iv + len(columns) + 1
            c = il + len(rows)
            yield {'r': r, 'c': c, 'label': label}
