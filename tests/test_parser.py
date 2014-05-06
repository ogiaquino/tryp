#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import tempfile
import unittest

from tryp.parser import parse_tryp


class TestParser(unittest.TestCase):
    def test_parse_tryp(self):
        tryp_file = tempfile.NamedTemporaryFile(delete=False)
        tryp_file.write("""
            {
             "crosstab": {
              "columns": ["Category"],
              "index": ["Regional","Region","Distributor","SR Code","SR Name"],
              "values": ["Target","Sell Out Actual"],
              "index_totals": ["Regional","Region","Distributor"],
              "columns_totals": ["Category"],
              "labels": {"values": {
                          "Ach": "% Ach",
                          "Target": "Target",
                          "Sell Out Actual": "Sell Out Actual"
                        }
              }
             }
            }
            """)
        tryp_file.close()
        json = parse_tryp(tryp_file.name, "crosstab")

        assert len(json) == 6
        assert "columns" in json
        assert "index" in json
        assert "values" in json
        assert "index_totals" in json
        assert "columns_totals" in json
        assert "labels" in json
        assert isinstance(json['columns'], list)
        assert isinstance(json['index'], list)
        assert isinstance(json['values'], list)
        assert isinstance(json['index_totals'], list)
        assert isinstance(json['columns_totals'], list)
        assert isinstance(json['labels'], dict)
        os.remove(tryp_file.name)
