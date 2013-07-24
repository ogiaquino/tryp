pyclean .
rm -f *.xls

for report in `ls report/*.tryp`
do
  base_file=`basename $report`
  trygen -f $report -d ${report%.*}.csv -o ${base_file%.*}.xls 
done
