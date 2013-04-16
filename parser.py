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
