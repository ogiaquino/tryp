import itertools
import numpy as np
import styles
from xlwt import easyxf, Workbook, Pattern, Font, Borders


def to_excel(tryp):
    sheetname = tryp.sheetname
    filename = tryp.filename
    wb = Workbook()
    ws = wb.add_sheet(sheetname)
    write_rows_labels(ws, tryp)
    write_columns_labels(ws, tryp)
    write_values_labels(ws, tryp)
    write_columns_totals_labels(ws, tryp)
    write_rows_totals_labels(ws, tryp)
    write_values(ws, tryp)
    wb.save(filename)


def write_rows_labels(ws, tryp):
    rows = tryp.rows
    columns = tryp.columns
    crosstab = tryp.crosstab
    plus_row = tryp.plus_row
    rmodule = tryp.rmodule
    xf = styles.get_styles(tryp.reportname, 'rows_labels')
    labels_rows = []

    font_red = Font()
    font_red.colour_index = 0x02
    font_red.bold = True
    font_red.name = 'sans-serif'
    font_red.height = 160
    borders_red = Borders()
    borders_red.top = 0x01
    borders_red.bottom= 0x01
    #borders_red.left = 0x01
    #borders_red.right = 0x02

    font_white = Font()
    font_white.colour_index = 0x08
    font_white.bold = True
    font_white.name = 'sans-serif'
    font_white.height = 160
    borders_white = Borders()
    borders_white.top = 0x01
    borders_white.bottom= 0x01
    #borders_white.left = 0x01
    #borders_white.right = 0x02

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
            r1 = x[0] + len(columns) + 1 + plus_row
            r2 = x[-1] + len(columns) + 1 + plus_row
            c1 = i
            c2 = i
            label = ci[x[0]]
            label = str(label).decode("utf-8")
            if '!' not in label:
                if sn in xf:
                    #FIXED THIS SHOULDNT BE LIKE THIS
                    if hasattr(rmodule.styles, 'conditional_rows_label'):
                        conditional_rows_labels = rmodule.styles.conditional_rows_label(tryp.connection, xf[sn])
                        if label in conditional_rows_labels['labels']:
                            label = conditional_rows_labels['labels'][label]
                            tet = conditional_rows_labels['xf']
                            labels_rows.append(r1)
                            ws.write_merge(r1, r2, c1, c2, label, tet)
                        else:
                            #FIXED THIS SHOULDNT BE LIKE THIS
                            if r1 in labels_rows:
                                xf[sn].font = font_red
                                #xf[sn].borders = borders_red
                            else:
                                xf[sn].font = font_white
                                #xf[sn].borders = borders_white
                            ws.write_merge(r1, r2, c1, c2, label, xf[sn])
                    else:
                        ws.write_merge(r1, r2, c1, c2, label, xf[sn])
                else:
                    ws.write_merge(r1, r2, c1, c2, label)


def write_columns_labels(ws, tryp):
    reportname = tryp.reportname
    rows = tryp.rows
    columns = tryp.columns
    crosstab = tryp.crosstab
    plus_row = tryp.plus_row
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


def write_values_labels(ws, tryp):
    reportname = tryp.reportname
    rows = tryp.rows
    columns = tryp.columns
    labels = tryp.labels
    crosstab = tryp.crosstab
    plus_row = tryp.plus_row
    column_counter_limit = tryp.column_counter_limit
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


def write_values(ws, tryp):
    reportname = tryp.reportname
    rows = tryp.rows
    columns = tryp.columns
    crosstab = tryp.crosstab
    plus_row = tryp.plus_row
    rmodule = tryp.rmodule
    header_info = rmodule.styles.headers(ws, tryp)
    if hasattr(rmodule.styles, 'conditional_formatting'):
        conditional_formatting = rmodule.styles.conditional_formatting
    else:
        conditional_formatting = None

    xf = styles.get_styles(reportname, 'values')
    #pat1 = Pattern()
    #pat1.pattern = Pattern.SOLID_PATTERN
    #pat1.pattern_fore_colour = 0x02
    #pat2 = Pattern()
    #pat2.pattern = Pattern.SOLID_PATTERN
    #pat2.pattern_fore_colour = 0x01
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
                #if labels_rows:
                #    newxf = xf[sn]
                #    if r in labels_rows:
                #        newxf.pattern = pat1
                #    else:
                #        newxf.pattern = pat2
                #    ws.write(r, c, label, newxf)
                #else:
                ws.write(r, c, label, xf[sn])
            else:
                ws.write(r, c, label)


def write_columns_totals_labels(ws, tryp):
    reportname = tryp.reportname
    rows = tryp.rows
    columns = tryp.columns
    values = tryp.values
    crosstab = tryp.crosstab
    plus_row = tryp.plus_row
    computed_values = tryp.computed_values
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


def write_rows_totals_labels(ws, tryp):
    reportname = tryp.reportname
    rows = tryp.rows
    columns = tryp.columns
    crosstab = tryp.crosstab
    plus_row = tryp.plus_row
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
