#import xlrd
#from xlutils.styles import Styles
#workbook = xlrd.open_workbook('dsr_excel_template.xls', formatting_info=True)
#worksheet = workbook.sheet_by_name('dsr_excel')
#print workbook.font_list[0].bold
#print workbook.font_list[0].struck_out
#print dir(workbook.font_list[0])


import xlrd
rd_xls_file = "reports/dsr_excel_template.xls"
xlrd_wb = xlrd.open_workbook(rd_xls_file, formatting_info=True)
 
# As an example: first sheet, first cell
rd_ws = xlrd_wb.sheet_by_index(0)
row = 2
col = 7
 
print dir(rd_ws)
print rd_ws.cell(2,7).value
print dir(xlrd_wb.xf_list[0])
import ipdb;ipdb.set_trace()
print xlrd_wb.xf_list[1].xf_index
rd_xf = xlrd_wb.xf_list[rd_ws.cell_xf_index(row,col)]
print xlrd_wb.font_list[rd_xf.font_index].bold
print xlrd_wb.font_list[rd_xf.font_index].name
