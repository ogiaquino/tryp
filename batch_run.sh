pyclean .
rm *.xls
trypgen -f reports/inventory_by_sku.tryp -d reports/inventory_by_sku.csv -o inventory_by_sku.xls
trypgen -f reports/dsr_excel.tryp -d reports/dsr_excel.csv -o dsr_excel.xls 
trypgen -f reports/bsr_excel.tryp -d reports/bsr_excel.csv -o bsr_excel.xls 
trypgen -f reports/inventory_by_dist.tryp -d reports/inventory_by_dist.csv -o inventory_by_dist.xls
