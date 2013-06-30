import json
import xml.etree.ElementTree as ET


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


def get_columns(report):
    try:
        columns = report.find("columns").text.split(',')
        return columns
    except:
        return []


def get_rows(report):
    rows = report.find("rows").text.split(',')
    return rows


def get_values(report):
    values = report.find("values").text.split(',')
    return values


def get_rows_results(report):
    rows = report.find("rows_results").text.split(',')
    return rows


def get_labels(report):
    labels = {}
    for child in report.find("labels").getchildren():
        text = {}
        for val in child.getchildren():
            label = val.getchildren()
            text[label[0].text] = label[1].text
        labels[child.tag] = text
    return labels


def get_computed_values(report):
    computed_values = []
    if report.find("computed_values") is not None:
        for child in report.find("computed_values").getchildren():
            text = {}
            text = child.text
            computed_values.append((child.tag, text))

        return computed_values
    else:
        return None


def get_report(root):
    return root.find("report")


def parse_tryp(tryp_file):
    # Try json format
    from pprint import pprint
    with open(tryp_file.split('/')[0] + '/dsr_excel.json') as data_file:
        data = json.load(data_file)

    tryp_reports = []
    tree = ET.parse(tryp_file)
    root = tree.getroot()

    report = get_report(root)
    dataset = get_dataset(root, report)
    query = get_query(dataset)
    conn_str = get_conn_str(root, dataset)
    print conn_str == _get_connection_str(data['connections'], 'olap')

    columns = get_columns(report)
    rows = get_rows(report)
    values = get_values(report)
    labels = get_labels(report)
    computed_values = get_computed_values(report)
    rows_results = get_rows_results(report)

    return {
        "query": query,
        "conn_str": conn_str,
        "columns": columns,
        "rows": rows,
        "values": values,
        "labels": labels,
        "computed_values": computed_values,
        "rows_results": rows_results,
    }

parse = parse_tryp


def _get_connection_str(connections, name):
    conn_str = "host='%(host)s' port='%(port)s' dbname='%(dbname)s' " \
               "user='%(user)s' password='%(password)s'" % \
               connections[name]
    return conn_str
