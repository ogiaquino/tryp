import psycopg2

import pandas.io.sql as psql
import numpy as np
import pandas as pd
import parser as pr

#from xlwt import easyxf, Borders, Workbook, Pattern, Style


def data_connection(conn_string):
    conn = psycopg2.connect(conn_string)
    return conn


def data_frame(query, connection):
    df = psql.frame_query(query, con=connection)
    return df


def crosstabulate(report):
    conn = data_connection(report["conn_str"])
    df = data_frame(report["query"], conn)

    groups = report['groups']
    aggregates = report['aggregates']
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


if __name__ == '__main__':
    reports = pr.parse_tryp("daily_sales.tryp")
    for report in reports:
        ct = crosstabulate(report)
        print ct.to_string(index=False)
