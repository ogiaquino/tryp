import psycopg2
import pandas.io.sql as psql
from pandas.io.parsers import read_csv

from excel import to_excel
from crosstab import crosstab as ct
from parser import parse


class Tryp(object):
    def __init__(self, reportname, sheetname, filename, dftype, result_level):
        self.reportname = reportname
        self.dftype = dftype
        self.result_level = result_level
        self.report = parse('%s/%s.tryp' % (self.reportname, self.reportname))

        self.connection = self.data_connection(self.report['conn_str'])
        self.df = self.data_frame(self.report['query'],
                                  self.connection, dftype)

        self.rows = self.report['rows']
        self.columns = self.report['columns']
        self.values = self.report['values']
        self.labels = self.report['labels']
        self.computed_values = self.report['computed_values'] or []

        self.sheetname = sheetname
        self.filename = filename

        self.crosstab = ct(self)

    def data_connection(self, conn_string):
        if self.dftype == 'db':
            try:
                conn = psycopg2.connect(conn_string)
                return conn
            except Exception, e:
                return None
        else:
            return None

    def data_frame(self, query, connection, dftype=None):
        if dftype == 'csv':
            df = read_csv('csv/%s.%s' % (self.reportname, 'csv'))
        if dftype == 'db':
            df = psql.frame_query(query, con=connection)
        return df


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Generate Report.')
    parser.add_argument('--reportname')
    parser.add_argument('--sheetname')
    parser.add_argument('--filename')
    parser.add_argument('--dftype', default='csv')
    parser.add_argument('--level', default=2)

    args = parser.parse_args()
    reportname = args.reportname
    sheetname = args.sheetname or reportname
    filename = args.filename or reportname + '.xls'
    dftype = args.dftype

    tryp = Tryp(reportname, sheetname, filename, dftype, args.level)

    to_excel(tryp)
