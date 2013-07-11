pyclean .
rm *.xls
trypgen -f inventory_by_sku.tryp -d inventory_by_sku.csv -o inventory_by_sku.xls
trypgen -f dsr_excel.tryp -d dsr_excel.csv -o dsr_excel.xls 
trypgen -f bsr_excel.tryp -d bsr_excel.csv -o bsr_excel.xls 
trypgen -f inventory_by_dist.tryp -d inventory_by_dist.csv -o inventory_by_dist.xls
