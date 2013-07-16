import xlrd
from xlwt import Style

rd_xls_file = "reports/dsr_excel_template.xls"
xlrd_wb = xlrd.open_workbook(rd_xls_file, formatting_info=True)
rd_ws = xlrd_wb.sheet_by_index(0)
row = 2
col = 7
rd_xf = xlrd_wb.xf_list[rd_ws.cell_xf_index(row,col)]

def style_values(ct, idx):
    return Style.default_style
    #return Style.default_style
    #ilevels[len(set(ct.df.index[x]))]
    #clevels[len(set(ct.df.columns[y])) - 1]
    #vlevels[(y % len(ct.levels.values) + 1) or len(ct.levels.values)]
