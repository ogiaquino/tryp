import itertools
import numpy as np
import styles
from xlwt import easyxf, Workbook


def to_excel(tryp):
    rmodule = __import__(tryp.reportname, globals(), locals(), ['styles'], -1)
    plus_row = rmodule.styles.plus_row
    rows = tryp.rows
    column_counter_limit = len(tryp.values) + len(tryp.computed_values)  -1
    columns = tryp.columns
    values = tryp.values
    labels = tryp.labels
    crosstab = tryp.crosstab
    connection = tryp.connection
    sheetname = tryp.sheetname
    filename = tryp.filename
    wb = Workbook()
    ws = wb.add_sheet(sheetname)

    if tryp.computed_values:
        module = __import__('%s.computed_values' % tryp.reportname,
                            fromlist=['computed_values'])
        crosstab = getattr(module, 'computed_values')(tryp)

    header_info = rmodule.styles.headers(ws, tryp.connection, crosstab)
    if hasattr(rmodule.styles, 'conditional_rows_label'):
        conditional_rows_label = rmodule.styles.conditional_rows_label(tryp.connection)
    else:
        conditional_rows_label = None
    write_rows_labels(tryp.reportname, rows, columns, crosstab, ws, plus_row, conditional_rows_label)

    write_columns_labels(tryp.reportname, rows, columns, crosstab, ws, plus_row)
    write_values_labels(tryp.reportname, rows, columns, labels, crosstab,
                        ws, plus_row, column_counter_limit)
    write_columns_totals_labels(tryp.reportname, rows, columns, values, crosstab,
                                ws, plus_row, tryp.computed_values)
    write_rows_totals_labels(tryp.reportname, rows, columns, crosstab, ws, plus_row)

    if hasattr(rmodule.styles, 'conditional_formatting'):
        conditional_formatting = rmodule.styles.conditional_formatting
    else:
        conditional_formatting = None
    write_values(tryp.reportname, rows, columns, crosstab, ws, plus_row, conditional_formatting, header_info)

    #merge the corner
    style = easyxf('borders: top medium;')
    ws.write_merge(0 + plus_row, len(columns)+plus_row, 0,
                   len(rows)-1, '', style)

    #borderize thick the last row
    for i in range(len(rows) + len(crosstab.values[0])):
        ws.write(len(crosstab.values)+ len(columns) + plus_row + 1, i, '', style)

    wb.save(filename)


def write_rows_labels(reportname, rows, columns, crosstab, ws, plus_row, conditional_rows_labels=None):
    xf = styles.get_styles(reportname, 'rows_labels')
    for i in range(len(rows)):
        sn = i
        ci = [x[i] for x in crosstab.index]
        label_group = [list(g) for k, g in itertools.groupby(ci)]
        count = -1
        index = []
        for group in label_group:
            index.append([])
            for g in group:
                count = count + 1
                index[-1].append(count)

        for x in index:
            r1 = x[0] + len(columns) + 1 + plus_row
            r2 = x[-1] + len(columns) + 1 + plus_row
            c1 = i
            c2 = i
            label = ci[x[0]]
            label = str(label).decode("utf-8")
            if '!' not in label:
                if conditional_rows_labels:
                    if label in conditional_rows_labels:
                        label = conditional_rows_labels[label]
                if sn in xf:
                    ws.write_merge(r1, r2, c1, c2, label, xf[sn])
                else:
                    ws.write_merge(r1, r2, c1, c2, label)


def write_columns_labels(reportname, rows, columns, crosstab, ws, plus_row):
    xf = styles.get_styles(reportname, 'columns_labels')
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
            r1 = i + plus_row
            r2 = i + plus_row
            c1 = x[0] + len(rows)
            c2 = x[-1] + len(rows)
            label = cc[x[0]]
            label = str(label).decode("utf-8")
            if '!' not in label:
                if sn in xf:
                    ws.write_merge(r1, r2, c1, c2, label, xf[sn])
                else:
                    ws.write_merge(r1, r2, c1, c2, label)


def write_values_labels(reportname, rows, columns, labels, crosstab, ws,
                        plus_row, column_counter_limit):
    column_counter = -1
    xf = styles.get_styles(reportname, 'values_labels')
    for i, cc in enumerate(crosstab.columns):
        column_counter = column_counter + 1
        r = len(columns) + plus_row
        c = len(rows) + i
        if isinstance(cc,  basestring):
            label = cc
        else:
            label = cc[-1]
        if labels:
            label = labels['values'][label]
        label = str(label).decode("utf-8")
        if column_counter in xf:
            ws.write(r, c, label, xf[column_counter])
        else:
            ws.write(r, c, label)
        if column_counter == column_counter_limit:
            column_counter = -1


def get_rows_total_style_name(crosstab, row):
    ci = [x[0] for x in crosstab.index[row] if '!' in x]
    ci = ''.join(ci)
    return ci


def get_columns_total_style_name(mck):
    ci = [x[0] for x in mck if '!' in x]
    return ''.join(ci)


def get_values_style_name(crosstab, row, column):
    ci = [x[0] for x in crosstab.index[row] if '!' in x]
    ci = ''.join(ci)
    cc_col = crosstab.columns[column]
    if not isinstance(cc_col,  basestring):
        cc_col = cc_col[-1]

    cc = [x[0] for x in crosstab.columns[column]
            if '!' in x] + [cc_col]
    cc = ''.join(cc)
    return ci + '%' + cc


def write_values(reportname, rows, columns, crosstab, ws, plus_row, conditional_formatting=None, header_info=None):
    xf = styles.get_styles(reportname, 'values')
    for iv, value in enumerate(crosstab.values):
        for il, label in enumerate(value):
            r = iv + len(columns) + 1 + plus_row
            c = il + len(rows)
            if np.isnan(label):
                label = '-'
            sn = get_values_style_name(crosstab, iv, il)
            if sn in xf:
                if conditional_formatting:
                    xf[sn] = conditional_formatting(xf[sn], header_info, label)
                ws.write(r, c, label, xf[sn])
            else:
                ws.write(r, c, label)


def write_columns_totals_labels(reportname, rows, columns, values, crosstab,
                                ws, plus_row, computed_values):
    xf = styles.get_styles(reportname, 'columns_total')
    merge_columns = {}

    for i, cc in enumerate(crosstab.columns):
        if cc[:-1] not in merge_columns and [c for c in cc if '!' in c]:
            merge_columns[cc[:-1]] = i

    for mck in merge_columns.keys():
        sn = get_columns_total_style_name(mck)
        for i, k in enumerate(mck):
            if '!' in k:
                r1 = i + plus_row
                r2 = len(columns) - 1 + plus_row
                c1 = merge_columns[mck] + len(rows)
                c2 = merge_columns[mck] + len(rows) + len(values) - 1
                if computed_values:
                    c2 = c2 + len(computed_values)
                if mck[-1] == '!':
                    label = 'Total Result'
                else:
                    label = mck[-1].replace('!', '')
                label = str(label).decode("utf-8")
                if sn in xf:
                    slabel = styles.get_labels(reportname, 'columns_total_labels')
                    if sn in slabel:
                        label = label.replace(' Result', slabel[sn])
                    ws.write_merge(r1, r2, c1, c2, label, xf[sn])
                else:
                    ws.write_merge(r1, r2, c1, c2, label)
                break


def write_rows_totals_labels(reportname, rows, columns, crosstab, ws, plus_row):
    xf = styles.get_styles(reportname, 'rows_total')
    merge_rows = {}
    for i, ci in enumerate(crosstab.index):
        if ci not in merge_rows and [c for c in ci if '!' in c]:
            merge_rows[ci] = i
    for mrk in sorted(merge_rows.keys()):
        for i, k in enumerate(mrk):
            if '!' in k:
                r1 = merge_rows[mrk] + len(columns) + 1 + plus_row
                r2 = merge_rows[mrk] + len(columns) + 1 + plus_row
                c1 = i
                c2 = len(rows) - 1
                if mrk[-1] == '!':
                    label = 'Grand Total'
                else:
                    label = mrk[-1].replace('!', '')
                    index = ((i + 1) * 10 + (i + 1)) * -1
                label = str(label).decode("utf-8")
                sn = get_rows_total_style_name(crosstab, i)
                if sn in xf:
                    ws.write_merge(r1, r2, c1, c2, label, xf[sn])
                else:
                    ws.write_merge(r1, r2, c1, c2, label)
                break
