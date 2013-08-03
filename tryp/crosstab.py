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
        self.excel = metadata.excel
        self.dataframe = self._crosstab(metadata.source_dataframe,
                                        self.xaxis,
                                        self.yaxis,
                                        self.zaxis)
        self._extend(metadata.extmodule)

    def to_excel(self):
        to_excel(self)

    def _extend(self, extmodule):
        if extmodule:
            extmodule = imp.load_source(extmodule[0], extmodule[1])
            extmodule.extend(self)
        self.values_labels = self._values_labels(self.dataframe)

    def _crosstab(self, source_dataframe, xaxis, yaxis, zaxis):
        df = source_dataframe.groupby(xaxis + yaxis).sum()
        df = df[zaxis].unstack(xaxis)
        if xaxis:
            df = self._xaxis_summary(source_dataframe,
                                     xaxis,
                                     yaxis,
                                     zaxis,
                                     df)
        return self._yaxis_summary(yaxis, df)

    def _values_labels(self, ct):
        if isinstance(ct.columns, pd.MultiIndex):
            return map(lambda x: x[-1], ct.columns)
        return ct.columns

    def _xaxis_summary(self, source_dataframe, xaxis, yaxis, zaxis, ctdf):
        sorter = self._init_sorter(self.visible_xaxis_summary,
                                   [col[1:] for col in ctdf.columns])

        ## CREATE SUBTOTALS FOR EACH COLUMNS
        for i in range(1, len(xaxis)):
            subtotal = source_dataframe.groupby(xaxis[:i] + yaxis).sum()[zaxis]
            subtotal = subtotal[zaxis].unstack(xaxis[:i])

            for col in subtotal.columns:
                scolumns = col + (col[-1],) * (len(xaxis) - len(col) + 1)
                ctdf[scolumns] = subtotal[col]
                sorter.append(self._rank(scolumns[1:],
                                         len(col[1:]) - 1,
                                         self.visible_xaxis_summary))
        ## END

        ## CREATE COLUMNS GRAND TOTAL
        for value in zaxis:
            total = source_dataframe.groupby(yaxis + xaxis[-1:]).sum()[value]
            keys = tuple([''] * len(xaxis))
            ctdf[(value,) + keys] = total.unstack(xaxis[-1:]).sum(axis=1)
            sorter.append(self._rank(keys, 0, xaxis))
        ## END

        ## REORDER AXIS 1 SO THAT AGGREGATES ARE THE LAST LEVEL
        order = range(1, len(xaxis) + 1) + [0]
        ct = ctdf.reorder_levels(order, axis=1)
        ## END

        return self._sorter(ct, sorter, ct.columns, 1)

    def _yaxis_summary(self, yaxis, df):
        sorter = self._init_sorter(self.visible_yaxis_summary, df.index)

        ## CREATE SUBTOTALS FOR EACH INDEX
        subtotals = []
        for i in range(len(self.visible_yaxis_summary)):
            for idx in set([x[:i+1] for x in df.index]):
                sindex = idx + (idx[-1],) * (len(yaxis) - len(idx))
                stotal = pd.DataFrame({sindex: df.ix[idx].sum()}).T
                subtotals.append(stotal)
                rank = self._rank(sindex,
                                  len(idx) - 1,
                                  self.visible_yaxis_summary)
                sorter.append(rank)
        ## END

        ## CREATE INDEX GRAND TOTAL
        gindex = tuple([''] * len(yaxis))
        gtotal = pd.DataFrame({gindex: df.ix[:].sum()}).T
        subtotals.append(gtotal)
        sorter.append(self._rank(gindex, 0, yaxis))
        ## END

        df = pd.concat([df] + subtotals)
        return self._sorter(df, sorter, df.index, 0)

    def _rank(self, keys, ranking, axis):
        rank = [np.NaN] * len(axis)
        rank[ranking] = 1
        rank = [x for x in roundrobin(keys, rank)]
        return rank

    def _init_sorter(self, visible_axis_summary, axis_labels):
        sorter = []
        nans = (np.NaN,) * len(visible_axis_summary)
        sorter = [[x for x in roundrobin(al, nans)] for al in axis_labels]
        return sorter

    def _sorter(self, df, sorter, index, axis):
        sorter = zip(*sorter)
        lexsort = np.lexsort([x for x in reversed(sorter)])
        sorted_index = []
        for lx in lexsort:
            idx = zip(*index)
            lex = tuple([idx[x][lx] for x in range(len(idx))])
            sorted_index.append(lex)
        return df.reindex_axis(axis=axis, labels=sorted_index)
