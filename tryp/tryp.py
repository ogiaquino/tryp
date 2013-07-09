from pandas.io.parsers import read_csv

from excel import to_excel
from dataset import Dataset
from parser import parse


class Tryp(object):
    def __init__(self, reportname, sheetname, filename, dftype, parameters):
        self.reportname = reportname
        self.report = parse('%s/%s.tryp' % (self.reportname, self.reportname))

        self.df = self.data_frame()

        self.rows = self.report['rows']
        self.columns = self.report['columns']
        self.values = self.report['values']
        self.labels = self.report['labels']
        self.rows_totals = self.report['rows_totals']

        self.sheetname = sheetname
        self.filename = filename

        self.crosstab = Dataset(self.df, self.rows, self.columns, self.values,
                                self.rows_totals).crosstab

    def data_frame(self):
        df = read_csv('csv/%s.%s' % (self.reportname, 'csv'))
        return df


def main():
        import argparse
        parser = argparse.ArgumentParser(description='Generate Report.')
        parser.add_argument('--reportname')
        parser.add_argument('--sheetname')
        parser.add_argument('--filename')
        parser.add_argument('--dftype', default='csv')
        parser.add_argument('-p', action='append')

        args = parser.parse_args()
        reportname = args.reportname
        sheetname = args.sheetname or reportname
        filename = args.filename or reportname + '.xls'
        dftype = args.dftype
        parameters = dict([p.split('=') for p in args.p or []])
        tryp = Tryp(reportname, sheetname, filename, dftype, parameters)

        to_excel(tryp)

if __name__ == '__main__':
    main()
