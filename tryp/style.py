import xlrd

template = "report/inventory_by_sku_templates.xls"

def get_values_styles(ct):
    wb = xlrd.open_workbook(template, formatting_info=True)
    ws = wb.sheet_by_index(0)
    row = 4
    col = 4
    print ws.cell(4,4).value
