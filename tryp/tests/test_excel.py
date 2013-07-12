import os
import unittest
import pandas as pd
from tryp.excel import Excel
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

        merge_labels_expected_result = {0: [(0, 0, u''),
                                            (1, 7, u'Central'),
                                            (8, 14, u'East')],

                                        1: [(0, 0, u''),
                                            (1, 1, u'Central'),
                                            (2, 4, u'Butterworth'),
                                            (5, 7, u'Ipoh'),
                                            (8, 8, u'East'),
                                            (9, 11, u'JB'),
                                            (12, 14, u'PJ')],

                                        2: [(0, 0, u''),
                                            (1, 1, u'Central'),
                                            (2, 2, u'Butterworth'),
                                            (3, 3, u'HINMARKETING'),
                                            (4, 4, u'KWANGHENG'),
                                            (5, 5, u'Ipoh'),
                                            (6, 6, u'CORESYN'),
                                            (7, 7, u'SGHEDERAN'),
                                            (8, 8, u'East'),
                                            (9, 9, u'JB'),
                                            (10, 10, u'LEIWAH'),
                                            (11, 11, u'WONDERF&B'),
                                            (12, 12, u'PJ'),
                                            (13, 13, u'HEBAT'),
                                            (14, 14, u'PENGEDAR')]}

        excel = Excel()
        labels = excel.merge_labels(ct.index, 3, 2)
        assert labels == merge_labels_expected_result
        assert len(labels) == len(rows)
        for i in labels:
            assert len(ct.index) - 1 == labels[i][-1][1]
