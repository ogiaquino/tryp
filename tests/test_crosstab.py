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
        ct.fillna(0.0, inplace=True)
        expected_df.fillna(0.0, inplace=True)

        for i in range(len(ct.to_records())):
            ct_val = ct.to_records()[i]
            ct_val = ct.to_records()[i]
            ct_val = tuple(ct_val)
            expected_val = list(expected_df.to_records()[i])[1:]

            # Need to convert to ''
            # since read_csv convert empty labels to NaN
            # need to fix this later
            if i == 0:
                expected_val[0:3] = ['','','']
            expected_val = tuple(expected_val)
            assert expected_val == ct_val
