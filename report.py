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

        def fuck(s):
            return sum([5])

        rows = ['region', 'area', 'distributor', 'salesrep_name']
        aggs = ['net_sales', 'sales_order_id_count']
        agg = {'net_sales':np.sum,'sales_order_id_count': fuck}

        df_raw = pivot_table(df, rows=rows, aggfunc=agg)

        wb = Workbook()
        ws = wb.add_sheet('SHIT')

        tmp = ["" for x in range(len(rows))]
        row_len = len(rows) - 1
        row_index = 0
        for row in df_raw.to_records():
            for i in range(row_len):
                if tmp[i] != row[0][i]:
                    row_index = row_index + 1
                    ws.write(row_index, i, row[0][i])
                    ws.write(row_index, i+1, row[0][i])
                    for j in range(len(aggs)):
                        h = df.groupby(rows[i]).agg(agg[aggs[j]])[aggs[j]][row[0][i]]
                        ws.write(row_index, len(rows)+j, h)

            tmp = row[0]
            row = list(row)
            row = ["" for x in range(row_len)] + [row[0][-1]] + row[1:]

            row_index = row_index + 1
            for i, col in enumerate(row):
                ws.write(row_index, i, col)
        wb.save('SHIT.xls')

import numpy as np
import pandas as pd


groups = ['region', 'area', 'distributor', 'salesrep_name']
agg = ['net_sales', 'sales_order_id_count']

total = pd.DataFrame({'region': ['Grand Total'],
                      'sales_order_id_count': df['sales_order_id_count'].sum(),
                      'net_sales': df['net_sales'].sum()})
total['total_rank'] = 1
def label(ser):
    return '{s} Total'.format(s=ser)

dfs = [total]
for i in range(1,len(groups)):
    parent_group = groups[:i]
    child_group = groups[i]
    group_total = df.groupby(parent_group, as_index=False).sum()
    group_total[child_group] = group_total[parent_group[-1]].apply(label)
    group_total[parent_group[-1] + '_rank'] = 1
    dfs.append(group_total)

dfs.append(df.groupby(groups, as_index=False).sum())

# UNION the DataFrames into one DataFrame
result = pd.concat(dfs)
# Replace NaNs with empty strings
result.fillna(dict([(x,'') for x in groups]), inplace=True)

groups_sort = []
for x in reversed(groups[:-1]):
    groups_sort.append(result[x + '_rank'].rank())
    groups_sort.append(result[x].rank())
groups_sort.append(result['total_rank'].rank())

# Reorder the rows
sorter = np.lexsort((groups_sort))
result = result.take(sorter)
result = result.reindex(columns=groups + aggs)
print result['net_sales']
result1 = result.to_dict()
result.to_csv('/tmp/test1.csv')
result.to_excel('SHIT1.xls')
#print(result.to_string(index=False))



    
    
#total = pd.DataFrame({'region': ['Grand Total'],
#                      'sales_order_id_count': df['sales_order_id_count'].sum(),
#                      'net_sales': df['net_sales'].sum()})
#
#region_total = df.groupby(['region'], as_index=False).sum()
#region_total['area'] = region_total['region'].apply(label)
#
#area_total = df.groupby(['region', 'area'], as_index=False).sum()
#area_total['distributor'] = area_total['area'].apply(label)
#
#dist_total = df.groupby( ['region', 'area', 'distributor'], as_index=False).sum()
#dist_total['salesrep_name'] = dist_total['distributor'].apply(label)
#
#rep_total = df.groupby(['region', 'area', 'distributor', 'salesrep_name'], as_index=False).sum()
#
#total['total_rank'] = 1
#region_total['region_rank'] = 1
#area_total['area_rank'] = 1
#
## UNION the DataFrames into one DataFrame
#result = pd.concat([total, region_total, area_total, dist_total, rep_total])
#
## Replace NaNs with empty strings
#result.fillna({'region': '', 'area': '', 'distributor': '', 'salesrep_name': ''}, inplace=True)
## Reorder the rows
#sorter = np.lexsort((
#        result['distributor'].rank(),
#        result['area_rank'].rank(),
#        result['area'].rank(),
#        result['region_rank'].rank(),
#        result['region'].rank(),
#        result['total_rank'].rank()))
#result = result.take(sorter)
#result = result.reindex(columns=['region', 'area', 'distributor', 'salesrep_name', 'net_sales', 'sales_order_id_count'])
##print(result.to_string(index=False))
#result2 = result.to_dict()
#result.to_csv('/tmp/test2.csv')
#print result1==result2
