import unittest
import pandas as pd
from tryp.dataset import Dataset


class TestDataset(unittest.TestCase):
    def test_crosstab(self):
        df = pd.read_csv('./tests/data/fixture.csv')
        rows = ['region', 'area', 'distributor']
        columns = ['salesrep', 'retailer']
        values = ['sales', 'invoice_count']
        rows_total = ['region','area','distributor']

        ct = Dataset(df, rows, columns, values, rows_total).crosstab
        expected_df = pd.read_csv('./tests/data/crosstab.csv')

        for i in range(len(ct.to_records())):
            ct_val = ct.fillna(0.0).to_records()[i]
            ct_val = tuple(ct_val)
            expected_val = list(expected_df.fillna(0.0).to_records()[i])[1:]
            expected_val = tuple(expected_val)
            assert ct_val == expected_val
