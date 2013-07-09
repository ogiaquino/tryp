import json


def parse_tryp(tryp_file):
    with open(tryp_file) as data_file:
        data = json.load(data_file)
        return data["crosstab"]

parse = parse_tryp
