#!/usr/bin/env python
# -*- coding: utf-8 -*-

import json


def parse_tryp(tryp_file, reportobj):
    with open(tryp_file) as data_file:
        data = json.load(data_file)
        return data[reportobj]

parse = parse_tryp
