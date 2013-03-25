import psycopg2
import numpy as np
import pandas.io.sql as psql
from pandas import *

class ReportEngine(object):
    def pandas_df(self, datasource, query):
        conn_string = "host='localhost' dbname='pandas' user='postgres' password='data01'"
        conn = psycopg2.connect(conn_string)
        df = psql.frame_query(query, con=conn)
        return df.to_html()


#return df.pivot_table(rows=['regional', 'region', 'distributor', 'SR Code'], cols=['category'], values=['Sell Out Actual'], aggfunc=np.sum, margins=True).to_html()
