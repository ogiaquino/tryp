import xlrd

template = "report/inventory_by_sku_templates.xls"


def get_values_styles(ct):
    wb = xlrd.open_workbook(template, formatting_info=True)
    ws = wb.sheet_by_index(0)

    yaxis = [''] + ct.visible_yaxis_summary + [ct.yaxis[-1]]
    xaxis = [''] + ct.xaxis
    for i, y in enumerate(yaxis):
        col = -1
        for x in xaxis:
            for z in ct.zaxis:
                col = col + 1
                print y, x, z, ws.cell(i + len(ct.xaxis) + 1,
                                       col + len(ct.yaxis)).value
