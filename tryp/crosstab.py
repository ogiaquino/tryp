import imp
import pandas as pd
import numpy as np

from common import roundrobin
from excel import to_excel as to_excel


class Levels(object):
    pass


class Crosstab(object):
    def __init__(self, meta):

        self.levels = Levels()
        self.levels.index = meta.index
        self.levels.columns = meta.columns
        self.levels.values = meta.values
        self.index_totals = meta.index_totals
        self.columns_totals = meta.columns_totals
        self.excel = meta.excel
        self.df = self._crosstab(meta.df,
                                 meta.index,
                                 meta.columns,
                                 meta.values)
        self._extend(meta.extmodule)

    def to_excel(self):
        to_excel(self)

    def _extend(self, extmodule):
        if extmodule:
            extmodule = imp.load_source(extmodule[0], extmodule[1])
            extmodule.extend(self)
        self.values_labels = self._values_labels(self.df)

    def _crosstab(self, df, index, columns, values):
        ctdf = df.groupby(index + columns).sum()[values].unstack(columns)
        if columns:
            ctdf = self._columns_totals(df, index, columns, values, ctdf)
        return self._index_totals(index, ctdf)

    def _values_labels(self, ct):
        if isinstance(ct.columns, pd.MultiIndex):
            return map(lambda x: x[-1], ct.columns)
        return ct.columns

    def _columns_totals(self, df, index, columns, values, ct):
        sorter = []
        nans = (np.NaN,) * len(self.columns_totals)
        sorter = [[x for x in roundrobin(col[1:], nans)] for col in ct.columns]

        ## CREATE SUBTOTALS FOR EACH COLUMNS
        for i in range(1, len(columns)):
            subtotal = df.groupby(index + columns[:i]).sum()[values]. \
                unstack(columns[:i])

            for col in subtotal.columns:
                ext = []
                for _ in range(len(col), len(columns) + 1):
                    ext.append('' + col[-1])
                ct[col + tuple(ext)] = subtotal[col]
                rank = [np.NaN] * len(self.columns_totals)
                rank[len(col[1:]) - 1] = 1
                rank = [x for x in roundrobin(col[1:] + tuple(ext), rank)]
                sorter.append(rank)
        ## END

        ## CREATE COLUMNS GRAND TOTAL
        for value in values:
            total = df.groupby(index + columns[-1:]).sum()
            label = tuple([''] * len(columns))
            key = (value,) + label
            ct[key] = total[value].unstack(columns[-1:]).sum(axis=1)
            rank = [1] * len(self.columns_totals)
            rank = [x for x in roundrobin(label, tuple(rank))]
            sorter.append(rank)
        ## END

        ## REORDER AXIS 1 SO THAT AGGREGATES ARE THE LAST LEVEL
        order = range(1, len(columns) + 1)
        order.append(0)
        ct = ct.reorder_levels(order, axis=1)
        ## END

        return self._sorter(ct, sorter, ct.columns, 1)

    def _index_totals(self, index, df):
        sorter = []
        nans = (np.NaN,) * len(self.index_totals)
        sorter = [[x for x in roundrobin(idx, nans)] for idx in df.index]

        ## CREATE SUBTOTALS FOR EACH INDEX
        subtotals = []
        for i in range(len(self.index_totals)):
            for idx in set([x[:i+1] for x in df.index]):
                sindex = idx + (idx[-1],) * (len(index) - len(idx))
                stotal = pd.DataFrame({sindex: df.ix[idx].sum()}).T
                subtotals.append(stotal)
                rank = [np.NaN] * len(self.index_totals)
                rank[len(idx) - 1] = 1
                sorter.append([x for x in roundrobin(sindex, rank)])
        ## END

        ## CREATE INDEX GRAND TOTAL
        gindex = tuple([''] * len(index))
        gtotal = pd.DataFrame({gindex: df.ix[:].sum()}).T
        subtotals.append(gtotal)
        rank = [1] * len(self.index_totals)
        sorter.append([x for x in roundrobin(gindex, tuple(rank))])
        ## END

        df = pd.concat([df] + subtotals)
        return self._sorter(df, sorter, df.index, 0)

    def _sorter(self, df, sorter, index, axis):
        sorter = zip(*sorter)
        lexsort = np.lexsort([x for x in reversed(sorter)])
        sorted_index = []
        for lx in lexsort:
            idx = zip(*index)
            lex = tuple([idx[x][lx] for x in range(len(idx))])
            sorted_index.append(lex)
        return df.reindex_axis(axis=axis, labels=sorted_index)
