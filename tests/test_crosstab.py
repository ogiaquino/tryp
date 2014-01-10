import os
import unittest
import pandas as pd
from pandas.util.testing import assert_frame_equal

from tryp.crosstab import Crosstab

data_loc = os.path.dirname(os.path.abspath(__file__)) + '/data'


class TestCrosstab(unittest.TestCase):
    def test_crosstab(self):
        df = pd.read_csv('%s/fixture.csv' % data_loc)
        index = ['region', 'area', 'distributor']
        columns = ['salesrep', 'retailer']
        values = ['sales', 'invoice_count']
        index_totals = ['region', 'area']
        columns_totals = ['salesrep']
        ct = Crosstab(columns, index, values,
                      columns_totals, index_totals, df)

        expected_dataframe = pd.read_pickle('%s/crosstab.df' % data_loc)
        assert_frame_equal(ct.dataframe, expected_dataframe)
        assert ct.dataframe.index.tolist() == \
            expected_dataframe.index.tolist()
        assert ct.dataframe.columns.tolist() == \
            expected_dataframe.columns.tolist()
        assert ct.dataframe.fillna(0.0).values.tolist() == \
            expected_dataframe.fillna(0.0).values.tolist()
