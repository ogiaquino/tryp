import json
import psycopg2
import xml.etree.ElementTree as ET

import pandas.io.sql as psql
import numpy as np

from pandas import *
from xlwt import easyxf, Borders, Workbook, Pattern, Style

def data_connection(conn_string):
    conn = psycopg2.connect(conn_string)
    return conn

def data_frame(query, connection):
    df = psql.frame_query(query, con=connection)
    return df

def generate_report(data_frame, format='crosstab', output='html'):
    if output=='html':
        return data_frame.to_html()

def parse_report(tryp_file):
    tryp_reports = []
    tree = ET.parse(tryp_file)
    root = tree.getroot()
    reports = root.find("reports")
    for report in reports:
        ds_name = report.find("dataset").text
        dataset = root.find(".//datasets/dataset/[@name='%s']" % ds_name)
        query = dataset.find('query').text
        conn_name =  dataset.find('connection').text
        connection = root.find(".//connections/connection/[@name='%s']" % conn_name)
        host = connection.find("host").text
        port = connection.find("port").text
        database = connection.find("database").text
        user = connection.find("user").text
        password = connection.find("password").text
        conn_str = "host='%s' port='%s' dbname='%s' user='%s' password='%s'" % (host, port, database, user, password)
        tryp_reports.append({"query": query, "conn_str": conn_str})
    return tryp_reports
    


if __name__ == '__main__':
    reports = parse_report("daily_sales.tryp")
    for report in reports:
        conn = data_connection(report["conn_str"])
        df = data_frame(report["query"], conn)
            
        df_raw = pivot_table(df, rows=['region','area','distributor','salesrep_name'], aggfunc={'net_sales':np.sum,'sales_order_id_count': np.sum})

        wb = Workbook()
        ws = wb.add_sheet('SHIT')

        #for i in range(3):
        #    tmp = ""
        #    for ir, row in enumerate(df_raw.T.iteritems()):
        #        if tmp == row[0][0 + i]:
        #            ws.write(ir + i, 0 + i, "")
        #        else:
        #            ws.write(ir + i, 0+i, row[0][0+i])
        #        tmp = row[0][0+i]


        #tmp = ""
        #for i, x in enumerate(df_raw.T.iteritems()):
        #    if tmp == x[0][0]:
        #        ws.write(i, 0, "")
        #    else:
        #        ws.write(i, 0, x[0][0])
        #    tmp = x[0][0]
        #tmp = ""
        #for i, x in enumerate(df_raw.T.iteritems()):
        #    if tmp == x[0][1]:
        #        ws.write(i+1, 1, "")
        #    else:
        #        ws.write(i+1, 1, x[0][1])
        #    tmp = x[0][1]
        #tmp = ""
        #for i, x in enumerate(df_raw.T.iteritems()):
        #    if tmp == x[0][2]:
        #        ws.write(i+2, 3, "")
        #    else:
        #        ws.write(i+2, 3, x[0][2])
        #    tmp = x[0][2]
        #tmp = ""
        #for i, x in enumerate(df_raw.T.iteritems()):
        #    if tmp == x[0][3]:
        #        ws.write(i+3, 5, "")
        #    else:
        #        ws.write(i+3, 5, x[0][3])
        #    tmp = x[0][3]
        tmp = ("","","")
        for i, row in enumerate(df_raw.to_records()):
            if tmp[0] != row[0][0]:
                tmp = row[0]
                row = list(row)
                row = list(row[0]) + row[1:]
                for j, c in enumerate(row):
                    ws.write(i+3, j, c)
                continue
            if tmp[1] != row[0][1]:
                tmp = row[0]
                row = list(row)
                row = list(row[0]) + row[1:]
                for j, c in enumerate(row):
                    ws.write(i+2, j, c)
                continue
            #if tmp[2] != row[0][2]:
            #    tmp = row[0]
            #    row = list(row)
            #    row = list(row[0]) + row[1:]
            #    for j, c in enumerate(row):
            #        ws.write(i+1, j, c)
            #    continue
        wb.save('SHIT.xls')
