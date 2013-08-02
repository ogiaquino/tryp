import os
import unittest
import pandas as pd
from tryp.crosstab import Crosstab

data_loc = os.path.dirname(os.path.abspath(__file__)) + '/data'


class TestCrosstab(unittest.TestCase):
    def test_crosstab(self):
        df = pd.read_csv('%s/fixture.csv' % data_loc)
        index = ['region', 'area', 'distributor']
        columns = ['salesrep', 'retailer']
        values = ['sales', 'invoice_count']
        index_totals = ['region', 'area']
        columns_totals = ['region', 'area', 'distributor']

        excel = {}
        excel['filename'] = 'filename'
        excel['sheetname'] = 'Sheet1'

        trypobj = type('tryp', (object,),
                       {'source_dataframe': df,
                        'yaxis': index,
                        'xaxis': columns,
                        'zaxis': values,
                        'visible_yaxis_summary': index_totals,
                        'visible_xaxis_summary': columns_totals,
                        'extmodule': None,
                        'excel': excel
                        })()

        ct = Crosstab(trypobj)
        expected_df = pd.read_csv('%s/crosstab.csv' % data_loc)
        ct.dataframe.fillna(0.0, inplace=True)
        expected_df.fillna(0.0, inplace=True)

        for i in range(len(ct.dataframe.to_records())):
            ct_val = ct.dataframe.to_records()[i]
            ct_val = ct.dataframe.to_records()[i]
            ct_val = tuple(ct_val)
            expected_val = list(expected_df.to_records()[i])[1:]

            # Need to convert to ''
            # since read_csv convert empty labels to NaN
            # need to fix this later
            if i == 0:
                expected_val[0:3] = ['', '', '']
            expected_val = tuple(expected_val)
            assert expected_val == ct_val
