import pandas as pd
import numpy as np
import xlsxwriter
from template import Template


def to_excel(ct):
    sheetname = ct.excel['sheetname']
    filename = ct.excel['filename']
    template = ct.excel['template']
    wb = xlsxwriter.Workbook(filename)
    ws = wb.add_worksheet(sheetname)
    if not ct.dataframe.empty:
        tmpl = Template(ct, wb)
        write_header(ct, ws, tmpl)
        write_axes(ct, wb, ws, tmpl)
        write_values(ct, wb, ws, tmpl)
        write_corner(ct, ws, tmpl)
        freeze_panes(ws, tmpl)
        borderize_floor(ct, ws, tmpl)
    else:
        write_merge(ws, 0, 0, 0, 2, 'Your report returned an empty dataset.')
        ws.hide_gridlines(2)
    wb.close()


def write_merge(ws, r1, c1, r2, c2, label, style=None):
    if r1 != r2 or c1 != c2:
        ws.merge_range(r1, c1, r2, c2, label, style)
    else:
        ws.write(r1, c1, label, style)


def borderize_floor(ct, ws, tmpl):
    row = tmpl.ws.nrows - tmpl.crosstab_loc[0] - 1
    style = tmpl.get_styles(row, 0)
    for c in range(len(ct.dataframe.columns) + len(ct.yaxis)):
        r = len(ct.dataframe.index) + tmpl.crosstab_loc[0] + len(ct.xaxis) + 1
        ws.write(r, c, '', style)


def freeze_panes(ws, tmpl):
    ws.hide_gridlines((tmpl.ws.show_grid_lines ^ 1) + 1)
    ws.freeze_panes(tmpl.ws.horz_split_pos, tmpl.ws.vert_split_pos)


def write_header(ct, ws, tmpl):
    for h in tmpl.styles['header']:
        row = h[0]
        col = h[1]
        try:
            label = h[2] % ct.datasets
        except:
            label = h[2]
        style = h[3]
        ws.set_row(row, style.row_height)
        ws.write(row, col, label, style)


def write_corner(ct, ws, tmpl):
    r1 = tmpl.crosstab_loc[0]
    r2 = tmpl.crosstab_loc[0] + len(ct.xaxis)
    c1 = tmpl.crosstab_loc[1]
    c2 = tmpl.crosstab_loc[1] + len(ct.yaxis) - 1
    write_merge(ws, r1, c1, r2, c2, "", tmpl.styles['corner'])


def write_axes(ct, wb, ws, tmpl):
    yaxis = ct.visible_yaxis_summary + [ct.yaxis[-1]] * \
        (len(ct.yaxis) - len(ct.visible_yaxis_summary))
    for idx in index(ct, tmpl, wb):
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
    ws.set_column(c1, c1, style.column_width)

    if idx['coordinate'] in axis[idx['axis']] and idx['coordinate'] != '':
        write_merge(ws, r1, c1, r2, c2, label.decode("utf-8"), style)
    else:
        # GRAND TOTAL/SUBTOTAL
        if idx['c1'] == len(ct.visible_yaxis_summary):
            c2 = len(ct.yaxis) - 1 + crosstab_col
            if style.label:
                label = label + style.label
            if idx['coordinate']:
                write_merge(ws, r1, ct.yaxis.index(idx['coordinate']) + 1,
                            r2, c2, label.decode("utf-8"), style)
            else:
                write_merge(ws, r1, 0, r2, c2, label.decode("utf-8"), style)


def _write_xaxis(ct, ws, idx, axis, tmpl):
    crosstab_row, crosstab_col = tmpl.crosstab_loc
    r1 = idx['r1'] + crosstab_row
    r2 = idx['r2'] + crosstab_row
    c1 = idx['c1'] + crosstab_col
    c2 = idx['c2'] + crosstab_col
    style = idx['style']
    label = idx['label']
    ws.set_row(r1, style.row_height)

    if idx['coordinate'] in axis[idx['axis']] and idx['coordinate'] != '':
        write_merge(ws, r1, c1, r2, c2, label.decode("utf-8"), style)
    else:
        # GRAND TOTAL/SUBTOTAL
        if idx['r1'] == len(ct.visible_xaxis_summary):
            r2 = len(ct.xaxis) - 1 + crosstab_row
            if style.label:
                label = label + style.label
            if idx['coordinate']:
                write_merge(ws, ct.xaxis.index(idx['coordinate']) +
                            crosstab_row + 1, c1, r2, c2,
                            label.decode("utf-8"), style)
            else:
                write_merge(ws, crosstab_row, c1, r2, c2,
                            label.decode("utf-8"), style)


def write_values(ct, wb, ws, tmpl):
    for idx in values_labels(ct, tmpl):
        _write_values_labels(ct, ws, idx, tmpl)

    for idx in values(ct, tmpl, wb):
        _write_values(ct, ws, idx, tmpl)


def _write_values(ct, ws, idx, tmpl):
    crosstab_row, crosstab_col = tmpl.crosstab_loc
    r = idx['r'] + crosstab_row
    c = idx['c'] + crosstab_col
    label = idx['label']
    style = idx['style']
    ws.write(r, c, label, style)
    ws.set_row(r, style.row_height)
    ws.set_column(c, c, style.column_width)


def _write_values_labels(ct, ws, idx, tmpl):
    crosstab_row, crosstab_col = tmpl.crosstab_loc
    r = idx['r'] + crosstab_row
    c = idx['c'] + crosstab_col
    style = idx['style']
    label = style.label or idx['label']
    ws.write(r, c, label, style)


def merge_indexes(indexes, index_width, total_width):
    labels = {}

    def __index(l):
        series = []
        index = [0]
        for i, v in enumerate(l):
            try:
                if l[i] != l[i + 1]:
                    series.append(v)
                    index.append(i + 1)
            except:
                series.append(v)
        return (series, index)

    def __labels(k, series):
        labels[k] = []
        lseries = series[1]
        lseries.append(len(indexes))
        for il, idx in enumerate(lseries[:-1]):
            labels[k].append((idx, lseries[il + 1] - 1, series[0][il]))

    for ir in range(total_width):
        series = pd.Series(zip(*indexes)[ir])
        series = __index(series)
        __labels(ir, series)

    for ir in range(total_width, index_width):
        series = pd.Series(zip(*indexes)[ir])
        __labels(ir, (series, [x for x in range(len(series))]))

    return labels


def index(ct, tmpl, wb):
    columns = ct.xaxis
    index_width = len(ct.yaxis)
    total_width = len(ct.visible_yaxis_summary)
    labels = merge_indexes(ct.dataframe.index, index_width, total_width)

    styles = tmpl.styles['index']
    for k in sorted(labels.keys()):
        for i, label in enumerate(labels[k]):
            coordinate = ct.coordinates['y'][label[0]]
            r1 = label[0] + len(columns) + 1
            r2 = label[1] + len(columns) + 1
            c1 = k
            c2 = k

            style = styles[(coordinate, k)]
            if ct.conditional_style:
                conditional_style = ct.conditional_style(wb, label, coordinate,
                                                     style)
                label = conditional_style['label']
                style = conditional_style['style']
            else:
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


def values(ct, tmpl, wb):
    levels_index = ct.yaxis
    levels_columns = ct.xaxis

    styles = tmpl.styles['values']
    for iv, value in enumerate(ct.dataframe.values):
        for il, label in enumerate(value):
            r = iv + len(levels_columns) + 1
            c = il + len(levels_index)

            y = ct.coordinates['y'][iv]
            x = '' if 'x' not in ct.coordinates else ct.coordinates['x'][il]
            z = ct.zaxis[il] if 'z' not in ct.coordinates else \
                ct.coordinates['z'][il]
            style = styles[(y, x, z)]
            if ct.conditional_style:
                conditional_style = ct.conditional_style(wb, label, z, style)
                style = conditional_style['style']

            if np.isnan(label):
                label = '-'
            yield {'r': r, 'c': c, 'label': label, 'style': style}
