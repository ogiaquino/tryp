import psycopg2
import xml.etree.ElementTree as ET

import pandas.io.sql as psql
import numpy as np
import pandas as pd

from xlwt import easyxf, Borders, Workbook, Pattern, Style


def data_connection(conn_string):
    conn = psycopg2.connect(conn_string)
    return conn


def data_frame(query, connection):
    df = psql.frame_query(query, con=connection)
    return df


def crosstabulate(data_frame, groups, aggregates):
    agg = {groups[0]: ['Grand Total']}
    for a in aggregates:
        agg[a[0]] = getattr(df[a[0]], a[1])()
    total = pd.DataFrame(agg)
    total['total_rank'] = 1
    dfs = [total]

    for i in range(1, len(groups)):
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
    result.fillna(dict([(x, '') for x in groups]), inplace=True)
    groups_sort = []
    for x in reversed(groups[:-1]):
        groups_sort.append(result[x + '_rank'].rank())
        groups_sort.append(result[x].rank())
    groups_sort.append(result['total_rank'].rank())

    # Reorder the rows
    sorter = np.lexsort((groups_sort))
    result = result.take(sorter)
    result = result.reindex(columns=groups + [x[0] for x in aggregates])
    return result


def label(ser):
    return '{s} Total'.format(s=ser)


def get_dataset(root, report):
    ds_name = report.find("dataset").text
    dataset = root.find(".//datasets/dataset/[@name='%s']" % ds_name)
    return dataset


def get_query(dataset):
    query = dataset.find('query').text
    return query


def get_conn_str(root, dataset):
    conn_name = dataset.find('connection').text
    connection = root.find(".//connections/connection/[@name='%s']" %
                           conn_name)
    host = connection.find("host").text
    port = connection.find("port").text
    database = connection.find("database").text
    user = connection.find("user").text
    password = connection.find("password").text
    conn_str = "host='%s' port='%s' dbname='%s' user='%s' password='%s'" % \
               (host, port, database, user, password)
    return conn_str


def get_groups(report):
    groups = report.find("groups").text.split(',')
    return groups


def get_aggregates(report):
    aggregates = [(x.find("column").text,
                  x.find("measure").text)
                  for x in report.findall("aggregate")]
    return aggregates


def parse_tryp(tryp_file):
    tryp_reports = []
    tree = ET.parse(tryp_file)
    root = tree.getroot()
    reports = root.find("reports")

    for report in reports:
        dataset = get_dataset(root, report)
        query = get_query(dataset)
        conn_str = get_conn_str(root, dataset)

        groups = get_groups(report)
        aggregates = get_aggregates(report)

        tryp_reports.append({"query": query,
                             "conn_str": conn_str,
                             "groups": groups,
                             "aggregates": aggregates})
    return tryp_reports

if __name__ == '__main__':
    reports = parse_tryp("daily_sales.tryp")
    for report in reports:
        conn = data_connection(report["conn_str"])
        df = data_frame(report["query"], conn)
        ct = crosstabulate(df, report['groups'], report['aggregates'])
        print ct.to_string(index=False)
