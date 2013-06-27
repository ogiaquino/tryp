#! virtualenv/bin/python
import psycopg2
import pandas.io.sql as psql
from pandas.io.parsers import read_csv

from excel import to_excel
from dataset import Dataset
from parser import parse


class Tryp(object):
    def __init__(self, reportname, sheetname, filename, dftype, parameters):
        self.reportname = reportname
        self.dftype = dftype
        self.report = parse('%s/%s.tryp' % (self.reportname, self.reportname))

        self.connection = self.data_connection(self.report['conn_str'],
                                               parameters)
        self.df = self.data_frame(self.report['query'],
                                  self.connection, dftype, parameters)

        self.rows = self.report['rows']
        self.columns = self.report['columns']
        self.values = self.report['values']
        self.labels = self.report['labels']
        self.computed_values = self.report['computed_values'] or []
        self.rows_results = self.report['rows_results']

        self.sheetname = sheetname
        self.filename = filename

        self.crosstab = Dataset(self.df, self.rows, self.columns, self.values,
                                self.rows_results).crosstab
        if self.computed_values:
            module = __import__('%s.computed_values' % self.reportname,
                                fromlist=['computed_values'])
            self.crosstab = getattr(module, 'computed_values')(self)
        self.rmodule = __import__(self.reportname, globals(), locals(),
                                  ['styles'], - 1)
        self.plus_row = self.rmodule.styles.plus_row
        self.column_counter_limit = len(self.values) + len(self.computed_values) - 1

    def data_connection(self, conn_string, parameters):
        if self.dftype == 'db':
            try:
                if 'tryp_db' in parameters:
                    conn = {}
                    trypdb = parameters['tryp_db'].split(':')
                    conn['user'] = trypdb[0]
                    conn['password'] = trypdb[1]
                    conn['host'] = trypdb[2]
                    conn['port'] = trypdb[3]
                    conn['dbname'] = trypdb[4]
                    conn_string = "host='%(host)s' port='%(port)s' dbname='%(dbname)s' user='%(user)s' password='%(password)s'" % conn
                conn = psycopg2.connect(conn_string)
                return conn
            except Exception, e:
                return None
        else:
            return None

    def data_frame(self, query, connection, dftype, parameters):
        if dftype == 'csv':
            df = read_csv('csv/%s.%s' % (self.reportname, 'csv'))
        if dftype == 'db':
            query = query % parameters
            df = psql.frame_query(query, con=connection)
        return df


if __name__ == '__main__':
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
