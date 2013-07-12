import imp
import pandas as pd
import numpy as np
from pandas.core.index import MultiIndex

foo = imp.load_source('dsr_excel.computed_values', '')


class Dataset(object):
    def __init__(self, df, rows, columns, values, rows_total):
        self.crosstab = self._crosstab(df, rows, columns, values, rows_total)
        self.crosstab.values_labels = self._values_labels(self.crosstab)

    def _crosstab(self, df, rows, columns, values, rows_total):
        ct = df.groupby(rows + columns).sum()[values].unstack(columns)
        if columns:
            ct = self._columns_totals(df, rows, columns, values, ct)
        ct = self._rows_totals(rows, rows_total, ct)
        return self._rename(ct)

    def _values_labels(self, ct):
        if isinstance(ct.columns, MultiIndex):
            return map(lambda x: x[-1], ct.columns)
        return ct.columns

    def _columns_totals(self, df, rows, columns, values, ct):
        ## CREATE SUBTOTALS FOR EACH COLUMNS
        for i in range(1, len(columns)):
            subtotal = df.groupby(rows + columns[:i]).sum()[values]. \
                unstack(columns[:i])

            for col in subtotal.columns:
                ext = []
                for _ in range(len(col), len(columns) + 1):
                    ext.append('!' + col[-1])
                ct[col + tuple(ext)] = subtotal[col]
        ## END

        ## CREATE GRAND TOTAL
        for value in values:
            total = df.groupby(rows + columns[-1:]).sum()
            key = (value,) + tuple(['!'] * len(columns))
            ct[key] = total[value].unstack(columns[-1:]).sum(axis=1)
        ## END

        ## REORDER AXIS 1 SO THAT AGGREGATES IS THE LAST LEVEL
        order = range(1, len(columns) + 1)
        order.append(0)
        ct = ct.reorder_levels(order, axis=1).sort_index(axis=1)
        ## END

        sorted_columns = []
        for key in sorted(set([cc[:-1] for cc in ct.columns])):
            for val in values:
                sorted_columns.append(key + (val,))

        ct = pd.DataFrame(ct, columns=sorted_columns)
        return ct

    def _rows_totals(self, rows, rows_total, ct):
        ct_row_subtotals = []
        if len(rows) > 1:
            ct_rows = []
            for row in set([x[:-1] for x in ct.index]):
                for i in range(len(rows)):
                    if row[:i+1] not in ct_rows:
                        ct_rows.append(row[:i+1])

            for row in ct_rows:
                result = tuple(['!' + row[-1]
                               for x in range(len(rows) - len(row))])
                rows_dict = dict([(r, len(rows) - i - 1) for i, r in
                                  enumerate(rows)])

                if len(result) in [rows_dict[r] for r in rows_total]:
                    index = row + result
                    row_df = pd.DataFrame({index: ct.ix[row].sum()}).T
                    ct_row_subtotals.append(row_df)

            total = {tuple(['!'] * len(rows)): ct.ix[:].sum()}
            total_df = pd.DataFrame(total).T
            ct_row_subtotals.append(total_df)
        else:
            total_df = pd.DataFrame({'!': ct.ix[:].sum()}).T
            ct_row_subtotals.append(total_df)

        ct = pd.concat([ct] + ct_row_subtotals)
        ct = ct.sort_index(axis=0)
        return ct

    def _rename(self, ct):
        if isinstance(ct.columns, pd.MultiIndex):
            col = map(lambda column: tuple([c.replace('!', '')
                                            for c in column]),
                      ct.columns)
            col = pd.MultiIndex.from_tuples(col, names=ct.columns.names)
        elif isinstance(ct.columns, pd.Index):
            col = [c.replace('!', '') for c in ct.columns]
            col = pd.Index(col)

        if isinstance(ct.index, pd.MultiIndex):
            idx = map(lambda index: tuple([i.replace('!', '') for i in index]),
                      ct.index)
            idx = pd.MultiIndex.from_tuples(idx, names=ct.index.names)
        elif isintance(ct.index, pd.Index):
            idx = [c.replace('!', '') for c in ct.index]
            idx = pd.Index(col)

        renamed = pd.DataFrame(ct.values, index=idx, columns=col)
        return renamed
