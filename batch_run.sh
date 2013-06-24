pyclean .
rm *.xls
python tryp.py --reportname=bsr_excel
python tryp.py --reportname=dsr_excel
python tryp.py --reportname=inventory_by_dist
python tryp.py --reportname=inventory_by_sku
