import unittest
import pandas as pd
import tryp.crosstab as ct


class TestCrosstab(unittest.TestCase):
    def test_group_all(self):
        rows = ['region', 'area']
        columns = ['distributor', 'salesrep']
        values = ['sales', 'invoice_count']
        df = pd.read_csv('tests/data/crosstab.csv', delimiter=',')
        group_all = ct.group_all(df, rows, columns, values)

        expected_columns = [
            ('sales', 'HINMARKETING', 'OSE'),
            ('sales', 'HINMARKETING', 'TLS'),
            ('sales', 'KWANGHENG', 'LBH'),
            ('sales', 'KWANGHENG', 'TCS'),
            ('sales', 'CORESYN', 'LILIAN'),
            ('sales', 'CORESYN', 'TEOH'),
            ('sales', 'SGHEDERAN', 'CHAN'),
            ('sales', 'SGHEDERAN', 'KAMACHI'),
            ('sales', 'LEIWAH', 'NF05'),
            ('sales', 'LEIWAH', 'NF06'),
            ('sales', 'WONDERF&B', 'MONC'),
            ('sales', 'WONDERF&B', 'SEREN'),
            ('sales', 'HEBAT', 'MIGI'),
            ('sales', 'HEBAT', 'OGI'),
            ('sales', 'PENGEDAR', 'NORM'),
            ('sales', 'PENGEDAR', 'SIMON'),
            ('invoice_count', 'HINMARKETING', 'OSE'),
            ('invoice_count', 'HINMARKETING', 'TLS'),
            ('invoice_count', 'KWANGHENG', 'LBH'),
            ('invoice_count', 'KWANGHENG', 'TCS'),
            ('invoice_count', 'CORESYN', 'LILIAN'),
            ('invoice_count', 'CORESYN', 'TEOH'),
            ('invoice_count', 'SGHEDERAN', 'CHAN'),
            ('invoice_count', 'SGHEDERAN', 'KAMACHI'),
            ('invoice_count', 'LEIWAH', 'NF05'),
            ('invoice_count', 'LEIWAH', 'NF06'),
            ('invoice_count', 'WONDERF&B', 'MONC'),
            ('invoice_count', 'WONDERF&B', 'SEREN'),
            ('invoice_count', 'HEBAT', 'MIGI'),
            ('invoice_count', 'HEBAT', 'OGI'),
            ('invoice_count', 'PENGEDAR', 'NORM'),
            ('invoice_count', 'PENGEDAR', 'SIMON')]
        assert expected_columns == [c for c in group_all.columns]

        expected_index = [
            ('Central', 'Butterworth'),
            ('Central', 'Ipoh'),
            ('East', 'JB'),
            ('East', 'PJ')]
        assert expected_index == [i for i in group_all.index]

        expected_values = [
            [1000.0, 1000.0, 1000.0, 1000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 50.0, 50.0, 50.0, 50.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
            [0.0, 0.0, 0.0, 0.0, 1000.0, 1000.0, 1000.0, 1000.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 50.0, 50.0, 50.0,
             50.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
            [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1000.0, 1000.0, 1000.0,
             1000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 50.0, 50.0, 50.0, 50.0, 0.0, 0.0, 0.0, 0.0],
            [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             1000.0, 1000.0, 1000.0, 1000.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 50.0, 50.0, 50.0, 50.0]]
        assert expected_values == group_all.fillna(0).values.tolist()

    def test_columns_totals(self):
        rows = ['region', 'area']
        columns = ['distributor', 'salesrep']
        values = ['sales', 'invoice_count']
        df = pd.read_csv('tests/data/crosstab.csv', delimiter=',')
        group_all = ct.group_all(df, rows, columns, values)
        crosstab = ct.columns_totals(rows, columns, values, group_all, df)

        expected_values = [
            [4000.0, 200.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 2000.0,
             100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
            [4000.0, 200.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0,
             50.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
            [4000.0, 200.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0],
            [4000.0, 200.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 2000.0, 100.0,
             1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]]
        assert expected_values == crosstab.fillna(0).values.tolist()

        expected_columns = [
            ('!', '!', 'sales'),
            ('!', '!', 'invoice_count'),
            ('CORESYN', '!CORESYN', 'sales'),
            ('CORESYN', '!CORESYN', 'invoice_count'),
            ('CORESYN', 'LILIAN', 'sales'),
            ('CORESYN', 'LILIAN', 'invoice_count'),
            ('CORESYN', 'TEOH', 'sales'),
            ('CORESYN', 'TEOH', 'invoice_count'),
            ('HEBAT', '!HEBAT', 'sales'),
            ('HEBAT', '!HEBAT', 'invoice_count'),
            ('HEBAT', 'MIGI', 'sales'),
            ('HEBAT', 'MIGI', 'invoice_count'),
            ('HEBAT', 'OGI', 'sales'),
            ('HEBAT', 'OGI', 'invoice_count'),
            ('HINMARKETING', '!HINMARKETING', 'sales'),
            ('HINMARKETING', '!HINMARKETING', 'invoice_count'),
            ('HINMARKETING', 'OSE', 'sales'),
            ('HINMARKETING', 'OSE', 'invoice_count'),
            ('HINMARKETING', 'TLS', 'sales'),
            ('HINMARKETING', 'TLS', 'invoice_count'),
            ('KWANGHENG', '!KWANGHENG', 'sales'),
            ('KWANGHENG', '!KWANGHENG', 'invoice_count'),
            ('KWANGHENG', 'LBH', 'sales'),
            ('KWANGHENG', 'LBH', 'invoice_count'),
            ('KWANGHENG', 'TCS', 'sales'),
            ('KWANGHENG', 'TCS', 'invoice_count'),
            ('LEIWAH', '!LEIWAH', 'sales'),
            ('LEIWAH', '!LEIWAH', 'invoice_count'),
            ('LEIWAH', 'NF05', 'sales'),
            ('LEIWAH', 'NF05', 'invoice_count'),
            ('LEIWAH', 'NF06', 'sales'),
            ('LEIWAH', 'NF06', 'invoice_count'),
            ('PENGEDAR', '!PENGEDAR', 'sales'),
            ('PENGEDAR', '!PENGEDAR', 'invoice_count'),
            ('PENGEDAR', 'NORM', 'sales'),
            ('PENGEDAR', 'NORM', 'invoice_count'),
            ('PENGEDAR', 'SIMON', 'sales'),
            ('PENGEDAR', 'SIMON', 'invoice_count'),
            ('SGHEDERAN', '!SGHEDERAN', 'sales'),
            ('SGHEDERAN', '!SGHEDERAN', 'invoice_count'),
            ('SGHEDERAN', 'CHAN', 'sales'),
            ('SGHEDERAN', 'CHAN', 'invoice_count'),
            ('SGHEDERAN', 'KAMACHI', 'sales'),
            ('SGHEDERAN', 'KAMACHI', 'invoice_count'),
            ('WONDERF&B', '!WONDERF&B', 'sales'),
            ('WONDERF&B', '!WONDERF&B', 'invoice_count'),
            ('WONDERF&B', 'MONC', 'sales'),
            ('WONDERF&B', 'MONC', 'invoice_count'),
            ('WONDERF&B', 'SEREN', 'sales'),
            ('WONDERF&B', 'SEREN', 'invoice_count')]
        assert expected_columns == [c for c in crosstab.columns]

    def test_rows_totals(self):
        rows = ['region', 'area']
        columns = ['distributor', 'salesrep']
        values = ['sales', 'invoice_count']
        rows_results = ['region']
        df = pd.read_csv('tests/data/crosstab.csv', delimiter=',')
        group_all = ct.group_all(df, rows, columns, values)
        crosstab = ct.columns_totals(rows, columns, values, group_all, df)
        crosstab = ct.rows_totals(rows, rows_results, crosstab)

        expected_values = [
            [16000.0, 800.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 2000.0,
             100.0, 1000.0, 50.0, 1000.0, 50.0, 2000.0, 100.0, 1000.0, 50.0,
             1000.0, 50.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 2000.0,
             100.0, 1000.0, 50.0, 1000.0, 50.0, 2000.0, 100.0, 1000.0, 50.0,
             1000.0, 50.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 2000.0,
             100.0, 1000.0, 50.0, 1000.0, 50.0],
            [8000.0, 400.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0,
             50.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 2000.0, 100.0,
             1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
            [4000.0, 200.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 2000.0,
             100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0],
            [4000.0, 200.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0],
            [8000.0, 400.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 2000.0, 100.0,
             1000.0, 50.0, 1000.0, 50.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0,
             50.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 2000.0, 100.0, 1000.0, 50.0,
             1000.0, 50.0],
            [4000.0, 200.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 2000.0, 100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 2000.0, 100.0,
             1000.0, 50.0, 1000.0, 50.0],
            [4000.0, 200.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 2000.0, 100.0,
             1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 2000.0,
             100.0, 1000.0, 50.0, 1000.0, 50.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
             0.0, 0.0, 0.0, 0.0, 0.0, 0.0]]

        assert expected_values == crosstab.fillna(0).values.tolist()
        crosstab.fillna(0).to_csv('ogi.csv')
