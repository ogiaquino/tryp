import psycopg2
import numpy as np
import pandas.io.sql as psql
import json
from pandas import *


def data_connection(conn_string):
    conn = psycopg2.connect(conn_string)
    return conn

def data_frame(query, connection):
    df = psql.frame_query(query, con=connection)
    return df

def generate_report(data_frame, format):
    if format=='html':
        return data_frame.to_html()



if __name__ == '__main__':
    import xml.etree.ElementTree as ET
    tree = ET.parse('daily_sales.xml')
    root = tree.getroot()
    for child in root:
        print child.tag, child.attrib
    #f = open('daily_sales.tryp')
    #report = json.loads(f.read())
    #conn = data_connection(report['connection'])
    #df = data_frame(report['query'], conn) 

    
    #print generate_report(df, 'html')
    


#print df.pivot_table(rows=['regional', 'region', 'distributor', 'SR Code'], cols=['category'], values=['Sell Out Actual'], aggfunc=np.sum, margins=True).to_html()
