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
              "rows": ["Regional","Region","Distributor","SR Code","SR Name"],
              "values": ["Target","Sell Out Actual"],
              "rows_totals": ["Regional","Region","Distributor"],
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
        json = parse_tryp(tryp_file.name)

        assert len(json) == 5
        assert "columns" in json
        assert "rows" in json
        assert "values" in json
        assert "rows_totals" in json
        assert "labels" in json
        assert isinstance(json['columns'], list)
        assert isinstance(json['rows'], list)
        assert isinstance(json['values'], list)
        assert isinstance(json['rows_totals'], list)
        assert isinstance(json['labels'], dict)
        os.remove(tryp_file.name)
