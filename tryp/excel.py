import pandas as pd
from xlwt import Workbook


def to_excel(ct):
    sheetname = ct.excel['sheetname']
    filename = ct.excel['filename']
    wb = Workbook()
    ws = wb.add_sheet(sheetname)
    write_axes(ct, ws)
    write_values(ct, ws)
    wb.save(filename)


def write_axes(ct, ws):
    def _write_axes(idx):
        r1 = idx['r1']
        r2 = idx['r2']
        c1 = idx['c1']
        c2 = idx['c2']
        label = idx['label']
        ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"))

    for idx in _index(ct):
        _write_axes(idx)

    for idx in _columns(ct):
        _write_axes(idx)


def write_values(ct, ws):
    def _write_values(idx):
        r = idx['r']
        c = idx['c']
        label = idx['label']
        ws.write(r, c, label)

    for idx in _values_labels(ct):
        _write_values(idx)

    for idx in _values(ct):
        _write_values(idx)


def _merge_indexes(indexes, index_width, total_width):
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


def _index(ct):
    columns = ct.levels.columns
    index_width = len(ct.levels.rows)
    total_width = len(ct.rows_totals)
    labels = _merge_indexes(ct.df.index, index_width, total_width)

    for k in sorted(labels.keys()):
        for label in labels[k]:
            r1 = label[0] + len(columns) + 1
            r2 = label[1] + len(columns) + 1
            c1 = k
            c2 = k
            label = label[2]
            yield {'r1': r1, 'r2': r2, 'c1': c1, 'c2': c2, 'label': label}


def _columns(ct):
    rows = ct.levels.rows
    index_width = len(ct.levels.columns)
    total_width = len(ct.columns_totals)
    labels = _merge_indexes(ct.df.columns, index_width, total_width)

    for k in sorted(labels.keys()):
        for label in labels[k]:
            r1 = k
            r2 = k
            c1 = label[0] + len(rows)
            c2 = label[1] + len(rows)
            label = label[2]
            yield {'r1': r1, 'r2': r2, 'c1': c1, 'c2': c2, 'label': label}


def _values_labels(ct):
    levels_rows = ct.levels.rows
    levels_columns = ct.levels.columns
    levels_values = ct.levels.values

    for i, cc in enumerate(levels_values):
        r = len(levels_columns)
        c = len(levels_rows) + i
        label = cc
        yield {'r': r, 'c': c, 'label': label}


def _values(ct):
    levels_rows = ct.levels.rows
    levels_columns = ct.levels.columns

    for iv, value in enumerate(ct.df.values):
        for il, label in enumerate(value):
            r = iv + len(levels_columns) + 1
            c = il + len(levels_rows)
            yield {'r': r, 'c': c, 'label': label}
