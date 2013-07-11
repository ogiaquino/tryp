import os
import unittest
import pandas as pd
from tryp.excel import *
from tryp.dataset import Dataset

data_loc = os.path.dirname(os.path.abspath(__file__)) + '/data'


class TestExcel(unittest.TestCase):
    def test_merge_labels(self):
        df = pd.read_csv('%s/fixture.csv' % data_loc)
        rows = ['region', 'area', 'distributor']
        columns = ['salesrep', 'retailer']
        values = ['sales', 'invoice_count']
        rows_total = ['region', 'area']

        ct = Dataset(df, rows, columns, values, rows_total).crosstab

        labels = merge_labels(ct.index, 3, 2)
        assert len(labels) == len(rows)
        for i in labels:
            assert len(ct.index) - 1 == labels[i][-1][1]
