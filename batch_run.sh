pyclean .
rm -f *.xls

for report in `ls report/*.tryp`
do
  base_file=`basename $report`
  trypgen -f $report -d ${report%.*}.csv -o ${base_file%.*}.xls -t ${report%.*}_template.xls
done
