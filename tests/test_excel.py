import os
import unittest
import pandas as pd
from tryp.excel import merge_indexes
from tryp.crosstab import Crosstab

data_loc = os.path.dirname(os.path.abspath(__file__)) + '/data'


class TestExcel(unittest.TestCase):
    def test_merge_indexes(self):
        df = pd.read_csv('%s/fixture.csv' % data_loc)
        index = ['region', 'area', 'distributor']
        columns = ['salesrep', 'retailer']
        values = ['sales', 'invoice_count']
        index_totals = ['region', 'area']
        columns_totals = ['region', 'area', 'distributor']
        ct = Crosstab(columns, index, values,
                      columns_totals, index_totals, df)
        merge_indexes_expected_result = {0: [(0, 0, u''),
                                             (1, 7, u'Bulacan'),
                                             (8, 14, u'Rizal')],

                                         1: [(0, 0, u''),
                                             (1, 1, u'Bulacan'),
                                             (2, 4, u'Baliuag'),
                                             (5, 7, u'Calumpit'),
                                             (8, 8, u'Rizal'),
                                             (9, 11, u'Binangonan'),
                                             (12, 14, u'Carmona')],

                                         2: [(0, 0, u''),
                                             (1, 1, u'Bulacan'),
                                             (2, 2, u'Baliuag'),
                                             (3, 3, u'HI'),
                                             (4, 4, u'KIT'),
                                             (5, 5, u'Calumpit'),
                                             (6, 6, u'CORN'),
                                             (7, 7, u'STI'),
                                             (8, 8, u'Rizal'),
                                             (9, 9, u'Binangonan'),
                                             (10, 10, u'LIB'),
                                             (11, 11, u'WAL'),
                                             (12, 12, u'Carmona'),
                                             (13, 13, u'HERB'),
                                             (14, 14, u'PAU')]}

        indexes = merge_indexes(ct.dataframe.index, 3, 2)
        assert indexes == merge_indexes_expected_result
        assert len(indexes) == len(index)
        for i in indexes:
            assert len(ct.dataframe.index) - 1 == indexes[i][-1][1]
