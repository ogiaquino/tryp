import pandas as pd
from xlwt import Workbook
from xlrd import open_workbook
from style import styles


def to_excel(ct):
    sheetname = ct.excel['sheetname']
    filename = ct.excel['filename']
    template = ct.excel['template']
    wb = Workbook()
    ws = wb.add_sheet(sheetname)

    wbt = open_workbook(template, formatting_info=True)
    wst = wbt.sheet_by_index(0)
    sty = styles(ct, wbt, wst)

    write_axes(ct, ws, sty)
    write_values(ct, ws, sty)
    wb.save(filename)


def write_axes(ct, ws, styles):
    yaxis = ct.visible_yaxis_summary + [ct.yaxis[-1]] * \
        (len(ct.yaxis) - len(ct.visible_yaxis_summary))
    for idx in index(ct, styles):
        _write_yaxis(ct, ws, idx, yaxis)

    if ct.xaxis:
        for idx in columns(ct, styles):
            _write_xaxis(ct, ws, idx, ct.xaxis,)


def _write_yaxis(ct, ws, idx, axis):
    r1 = idx['r1']
    r2 = idx['r2']
    c1 = idx['c1']
    c2 = idx['c2']
    label = idx['label']
    style = idx['style']

    if idx['coordinate'] in axis[idx['axis']] and idx['coordinate'] != '':
        ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"), style)
    else:
        # GRAND TOTAL/SUBTOTAL
        try:
            ws.write_merge(r1, r2, c1, len(ct.yaxis) - 1,
                           label.decode("utf-8"), style)
        except:
            pass


def _write_xaxis(ct, ws, idx, axis):
    r1 = idx['r1']
    r2 = idx['r2']
    c1 = idx['c1']
    c2 = idx['c2']
    label = idx['label']
    style = idx['style']

    if idx['coordinate'] in axis[idx['axis']] and idx['coordinate'] != '':
        ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"), style)
    else:
        # GRAND TOTAL/SUBTOTAL
        try:
            ws.write_merge(r1, len(ct.xaxis) - 1, c1, c2,
                           label.decode("utf-8"), style)
        except:
            pass


def write_values(ct, ws, styles):
    for idx in values_labels(ct, styles):
        _write_values_labels(ct, ws, idx)

    for idx in values(ct, styles):
        _write_values(ct, ws, idx)


def _write_values(ct, ws, idx):
    r = idx['r']
    c = idx['c']
    label = idx['label']
    style = idx['style']
    ws.write(r, c, label, style)


def _write_values_labels(ct, ws, idx):
    r = idx['r']
    c = idx['c']
    label = idx['label']
    style = idx['style']
    ws.write(r, c, label, style)


def merge_indexes(indexes, index_width, total_width):
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


def index(ct, styles):
    columns = ct.xaxis
    index_width = len(ct.yaxis)
    total_width = len(ct.visible_yaxis_summary)
    labels = merge_indexes(ct.dataframe.index, index_width, total_width)

    styles = styles['index']
    for k in sorted(labels.keys()):
        for i, label in enumerate(labels[k]):
            coordinate = ct.coordinates['y'][label[0]]
            style = styles[(coordinate, k)]
            r1 = label[0] + len(columns) + 1
            r2 = label[1] + len(columns) + 1
            c1 = k
            c2 = k
            label = label[2]
            yield {'r1': r1, 'r2': r2, 'c1': c1, 'c2': c2, 'label': label,
                   'style': style, 'coordinate': coordinate, 'axis': k}


def columns(ct, styles):
    index = ct.yaxis
    columns_width = len(ct.xaxis)
    total_width = len(ct.visible_xaxis_summary) + 1
    labels = merge_indexes(ct.dataframe.columns, columns_width, total_width)

    styles = styles['column']
    for k in sorted(labels.keys()):
        for i, label in enumerate(labels[k]):
            coordinate = ct.coordinates['x'][label[0]]
            style = styles[(k,
                            ct.coordinates['x'][label[0]],
                            ct.coordinates['z'][label[0]])]
            r1 = k
            r2 = k
            c1 = label[0] + len(index)
            c2 = label[1] + len(index)
            label = label[2]
            yield {'r1': r1, 'r2': r2, 'c1': c1, 'c2': c2, 'label': label,
                   'style': style, 'coordinate': coordinate, 'axis': k}


def values_labels(ct, styles):
    levels_index = ct.yaxis
    levels_columns = ct.xaxis
    levels_values = ct.values_labels

    styles = styles['values_labels']
    for i, cc in enumerate(levels_values):
        style = styles[ct.coordinates['z'][i]]
        r = len(levels_columns)
        c = len(levels_index) + i
        label = cc
        yield {'r': r, 'c': c, 'label': label, 'style': style}


def values(ct, styles):
    levels_index = ct.yaxis
    levels_columns = ct.xaxis

    styles = styles['values']
    for iv, value in enumerate(ct.dataframe.values):
        for il, label in enumerate(value):
            style = styles[(ct.coordinates['y'][iv],
                            ct.coordinates['x'][il],
                            ct.coordinates['z'][il])]
            r = iv + len(levels_columns) + 1
            c = il + len(levels_index)
            yield {'r': r, 'c': c, 'label': label, 'style': style}
