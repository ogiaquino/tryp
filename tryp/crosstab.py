import imp
import pandas as pd
import numpy as np

from common import roundrobin
from excel import to_excel as to_excel


class Crosstab(object):
    def __init__(self, metadata):

        self.xaxis = metadata.xaxis
        self.yaxis = metadata.yaxis
        self.zaxis = metadata.zaxis
        self.visible_xaxis_summary = metadata.visible_xaxis_summary
        self.visible_yaxis_summary = metadata.visible_yaxis_summary
        self.source_dataframe = metadata.source_dataframe
        self.excel = metadata.excel
        self.ctdataframe = self._crosstab(self.source_dataframe,
                                          self.yaxis,
                                          self.xaxis,
                                          self.zaxis)
        self._extend(metadata.extmodule)

    def to_excel(self):
        to_excel(self)

    def _extend(self, extmodule):
        if extmodule:
            extmodule = imp.load_source(extmodule[0], extmodule[1])
            extmodule.extend(self)
        self.values_labels = self._values_labels(self.ctdataframe)

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
        nans = (np.NaN,) * len(self.visible_xaxis_summary)
        sorter = [[x for x in roundrobin(col[1:], nans)] for col in ct.columns]

        ## CREATE SUBTOTALS FOR EACH COLUMNS
        for i in range(1, len(columns)):
            subtotal = df.groupby(index + columns[:i]).sum()[values]. \
                unstack(columns[:i])

            for col in subtotal.columns:
                ext = ['' + col[-1]] * (len(columns) - len(col) + 1)
                ct[col + tuple(ext)] = subtotal[col]
                rank = [np.NaN] * len(self.visible_xaxis_summary)
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
            rank = [np.NaN] * len(columns)
            rank[0] = 1
            sorter.append([x for x in roundrobin(label, rank)])
        ## END

        ## REORDER AXIS 1 SO THAT AGGREGATES ARE THE LAST LEVEL
        order = range(1, len(columns) + 1)
        order.append(0)
        ct = ct.reorder_levels(order, axis=1)
        ## END

        return self._sorter(ct, sorter, ct.columns, 1)

    def _index_totals(self, index, df):
        sorter = []
        nans = (np.NaN,) * len(self.visible_yaxis_summary)
        sorter = [[x for x in roundrobin(idx, nans)] for idx in df.index]

        ## CREATE SUBTOTALS FOR EACH INDEX
        subtotals = []
        for i in range(len(self.visible_yaxis_summary)):
            for idx in set([x[:i+1] for x in df.index]):
                sindex = idx + (idx[-1],) * (len(index) - len(idx))
                stotal = pd.DataFrame({sindex: df.ix[idx].sum()}).T
                subtotals.append(stotal)
                rank = [np.NaN] * len(self.visible_yaxis_summary)
                rank[len(idx) - 1] = 1
                sorter.append([x for x in roundrobin(sindex, rank)])
        ## END

        ## CREATE INDEX GRAND TOTAL
        gindex = tuple([''] * len(index))
        gtotal = pd.DataFrame({gindex: df.ix[:].sum()}).T
        subtotals.append(gtotal)
        rank = [np.NaN] * len(self.visible_yaxis_summary)
        rank[0] = 1
        sorter.append([x for x in roundrobin(gindex, rank)])
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
