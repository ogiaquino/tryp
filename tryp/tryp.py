"""Usage:
    tryp.py -f tryp_file -d csv_file -o output_file  -t template_file
    tryp.py -f tryp_file -o output_file -t template_file \
 [--dbhost=<arg>] [--dbport=<arg>] [--dbname=<arg>] [--dbuser=<arg>] \
 [--dbpassword=<arg>] [--sqlfile=<arg>] [--sqlparams=<arg>]...

   Options:
          -f tryp_file.
          -d csv_file.
          -o out_putfile
          -t template_file
          --connstring connstring
          --dbhost database hostname
          --dbport database port
          --dbname database name
          --dbuser database username
          --dbpassword database password
          --sqlfile sql_file
          --sqlparams sqlparams
"""
import os
import psycopg2
from pandas.io.parsers import read_csv
from pandas.io import sql as psql
from parser import parse
from crosstab import Crosstab


class CrosstabMetaData(object):
    def __init__(self, tryp_file, template_file, output_file, csv_file,
                 dbhost, dbport, dbname, dbuser, dbpassword,
                 sql_file, sqlparams):
        self.report = parse(tryp_file, 'crosstab')
        self.source_dataframe = self.data_frame(csv_file, dbhost, dbport,
                                                dbname, dbuser, dbpassword,
                                                sql_file, sqlparams)
        self.xaxis = self.report['xaxis']
        self.yaxis = self.report['yaxis']
        self.zaxis = self.report['zaxis']
        self.visible_yaxis_summary = self.report['visible_yaxis_summary']
        self.visible_xaxis_summary = self.report['visible_xaxis_summary']
        self.excel = {}
        self.excel['filename'] = output_file
        self.excel['template'] = template_file
        self.excel['sheetname'] = self.report['sheetname']
        self.extmodule = self.is_extmodule_exist(tryp_file)

    def is_extmodule_exist(self, tryp_file):
        tryp_path = os.path.abspath(tryp_file)
        tryp_path, tryp_file = os.path.split(tryp_file)
        tryp_filename = os.path.splitext(tryp_file)[0]
        extmodule = os.path.join(tryp_path, tryp_filename + '.py')
        if os.path.exists(extmodule):
            return (tryp_filename, extmodule)

    def data_frame(self, csv_file, dbhost, dbport, dbname, dbuser, dbpassword,
                   sql_file, sqlparams):
        if csv_file:
            df = read_csv(csv_file)
            self.connection = None
        else:
            query = open(sql_file).read()
            params = dict([params.split('=') for params in sqlparams])
            connstring = "host='%(host)s' port='%(port)s'" \
                         "dbname='%(dbname)s' user='%(user)s'" \
                         "password='%(password)s'"
            connstring = connstring % {'host': dbhost,
                                       'port': dbport,
                                       'dbname': dbname,
                                       'user': dbuser,
                                       'password': dbpassword}
            conn = psycopg2.connect(connstring or self.report['connstring'])
            self.connection = conn
            print conn
            f = open('/tmp/query.sql', 'w')
            f.write(query % params)
            f.close()
            df = psql.frame_query(query % params, con=conn)
        return df


class Tryp(object):
    def __init__(self, tryp_file, template_file, output_file, csv_file,
                 dbhost, dbport, dbname, dbuser, dbpassword, sql_file,
                 sqlparams):
        ctmeta = CrosstabMetaData(tryp_file, template_file, output_file,
                                  csv_file, dbhost, dbport, dbname, dbuser,
                                  dbpassword, sql_file, sqlparams)
        self.crosstab = Crosstab(ctmeta)


def main():
    from docopt import docopt
    args = docopt(__doc__, version='Tryp 0.0')
    tryp_file = args['-f']
    template_file = args['-t']
    output_file = args['-o']
    csv_file = args['-d']
    sql_file = args['--sqlfile']
    sqlparams = args['--sqlparams']
    dbhost = args['--dbhost']
    dbport = args['--dbport']
    dbname = args['--dbname']
    dbuser = args['--dbuser']
    dbpassword = args['--dbpassword']
    Tryp(tryp_file, template_file, output_file, csv_file, dbhost, dbport,
         dbname, dbuser, dbpassword, sql_file, sqlparams).crosstab.to_excel()

if __name__ == '__main__':
    main()
