import pandas as pd
import numpy as np
from xlwt import Workbook
from template import Template


def to_excel(ct):
    sheetname = ct.excel['sheetname']
    filename = ct.excel['filename']
    template = ct.excel['template']
    wb = Workbook()
    ws = wb.add_sheet(sheetname)
    tmpl = Template(ct)
    write_axes(ct, ws, tmpl)
    write_values(ct, ws, tmpl)
    write_corner(ct, ws, tmpl)
    write_header(ct, ws, tmpl)
    freeze_panes(ws, tmpl)
    wb.save(filename)


def freeze_panes(ws, tmpl):
    ws.set_panes_frozen(tmpl.ws.panes_are_frozen)
    ws.set_horz_split_pos(tmpl.ws.horz_split_pos)
    ws.set_vert_split_pos(tmpl.ws.vert_split_pos)
    ws.show_grid = False


def write_header(ct, ws, tmpl):
    for h in tmpl.styles['header']:
        row = h[0]
        col = h[1]
        try:
            label = h[2] % ct.labels
        except:
            label = h[2]
        style = h[3]
        ws.write(h[0], h[1], label, h[3])


def write_corner(ct, ws, tmpl):
    r1 = tmpl.crosstab_loc[0]
    r2 = tmpl.crosstab_loc[0] + len(ct.xaxis)
    c1 = tmpl.crosstab_loc[1]
    c2 = tmpl.crosstab_loc[1] + len(ct.yaxis) - 1
    ws.write_merge(r1, r2, c1, c2, '', tmpl.styles['corner'])


def write_axes(ct, ws, tmpl):
    yaxis = ct.visible_yaxis_summary + [ct.yaxis[-1]] * \
        (len(ct.yaxis) - len(ct.visible_yaxis_summary))
    for idx in index(ct, tmpl):
        _write_yaxis(ct, ws, idx, yaxis, tmpl)

    if ct.xaxis:
        for idx in columns(ct, tmpl):
            _write_xaxis(ct, ws, idx, ct.xaxis, tmpl)


def _write_yaxis(ct, ws, idx, axis, tmpl):
    crosstab_row, crosstab_col = tmpl.crosstab_loc
    r1 = idx['r1'] + crosstab_row
    r2 = idx['r2'] + crosstab_row
    c1 = idx['c1'] + crosstab_col
    c2 = idx['c2'] + crosstab_col
    label = idx['label']
    style = idx['style']

    if idx['coordinate'] in axis[idx['axis']] and idx['coordinate'] != '':
        ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"), style)
    else:
        # GRAND TOTAL/SUBTOTAL
        try:
            c2 = len(ct.yaxis) - 1 + crosstab_col
            if style.label:
                label = label + style.label
            ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"), style)
        except:
            pass


def _write_xaxis(ct, ws, idx, axis, tmpl):
    crosstab_row, crosstab_col = tmpl.crosstab_loc
    r1 = idx['r1'] + crosstab_row
    r2 = idx['r2'] + crosstab_row
    c1 = idx['c1'] + crosstab_col
    c2 = idx['c2'] + crosstab_col
    style = idx['style']
    label = idx['label']

    if idx['coordinate'] in axis[idx['axis']] and idx['coordinate'] != '':
        ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"), style)
    else:
        # GRAND TOTAL/SUBTOTAL
        try:
            r2 = len(ct.xaxis) - 1 + crosstab_row
            if style.label:
                label = label + style.label
            ws.write_merge(r1, r2, c1, c2, label.decode("utf-8"), style)
        except:
            pass


def write_values(ct, ws, tmpl):
    for idx in values_labels(ct, tmpl):
        _write_values_labels(ct, ws, idx, tmpl)

    for idx in values(ct, tmpl):
        _write_values(ct, ws, idx, tmpl)


def _write_values(ct, ws, idx, tmpl):
    crosstab_row, crosstab_col = tmpl.crosstab_loc
    r = idx['r'] + crosstab_row
    c = idx['c'] + crosstab_col
    label = idx['label']
    style = idx['style']
    ws.write(r, c, label, style)


def _write_values_labels(ct, ws, idx, tmpl):
    crosstab_row, crosstab_col = tmpl.crosstab_loc
    r = idx['r'] + crosstab_row
    c = idx['c'] + crosstab_col
    style = idx['style']
    label = style.label or idx['label']
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


def index(ct, tmpl):
    columns = ct.xaxis
    index_width = len(ct.yaxis)
    total_width = len(ct.visible_yaxis_summary)
    labels = merge_indexes(ct.dataframe.index, index_width, total_width)

    styles = tmpl.styles['index']
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


def columns(ct, tmpl):
    index = ct.yaxis
    columns_width = len(ct.xaxis)
    total_width = len(ct.visible_xaxis_summary) + 1
    labels = merge_indexes(ct.dataframe.columns, columns_width, total_width)

    styles = tmpl.styles['column']
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


def values_labels(ct, tmpl):
    levels_index = ct.yaxis
    levels_columns = ct.xaxis
    levels_values = ct.values_labels

    styles = tmpl.styles['values_labels']
    for i, cc in enumerate(levels_values):
        z = ct.zaxis[i] if 'z' not in ct.coordinates else \
            ct.coordinates['z'][i]
        style = styles[z]
        r = len(levels_columns)
        c = len(levels_index) + i
        label = cc
        yield {'r': r, 'c': c, 'label': label, 'style': style}


def values(ct, tmpl):
    levels_index = ct.yaxis
    levels_columns = ct.xaxis

    styles = tmpl.styles['values']
    for iv, value in enumerate(ct.dataframe.values):
        for il, label in enumerate(value):
            y = ct.coordinates['y'][iv]
            x = '' if 'x' not in ct.coordinates else ct.coordinates['x'][il]
            z = ct.zaxis[il] if 'z' not in ct.coordinates else \
                ct.coordinates['z'][il]

            style = styles[(y, x, z)]
            r = iv + len(levels_columns) + 1
            c = il + len(levels_index)
            if np.isnan(label):
                label = '-'
            yield {'r': r, 'c': c, 'label': label, 'style': style}
