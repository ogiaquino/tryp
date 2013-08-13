import imp
import pandas as pd
import numpy as np

from common import roundrobin
from excel import to_excel as to_excel


class Crosstab(object):
    def __init__(self, metadata):
        self.xcoord, self.ycoord, self.zcoord = [], [], []
        self.xaxis = metadata.xaxis
        self.yaxis = metadata.yaxis
        self.zaxis = metadata.zaxis
        self.visible_xaxis_summary = metadata.visible_xaxis_summary
        self.visible_yaxis_summary = metadata.visible_yaxis_summary
        self.excel = metadata.excel
        self.dataframe = self.__crosstab(metadata.source_dataframe,
                                         self.xaxis,
                                         self.yaxis,
                                         self.zaxis)
        self.__extend(metadata.extmodule)

    def to_excel(self):
        to_excel(self)

    def __extend(self, extmodule):
        if extmodule:
            extmodule = imp.load_source(extmodule[0], extmodule[1])
            extmodule.extend(self)
        self.values_labels = self.__values_labels(self.dataframe)

    def __crosstab(self, source_dataframe, xaxis, yaxis, zaxis):
        df = source_dataframe.groupby(xaxis + yaxis).sum()
        df = df[zaxis].unstack(xaxis)
        if xaxis:
            df = self.__xaxis_summary(source_dataframe,
                                      xaxis,
                                      yaxis,
                                      zaxis,
                                      df)
        return self.__yaxis_summary(yaxis, df)

    def __values_labels(self, ct):
        if isinstance(ct.columns, pd.MultiIndex):
            return map(lambda x: x[-1], ct.columns)
        return ct.columns

    def __xaxis_summary(self, source_dataframe, xaxis, yaxis, zaxis, ctdf):
        for idx in ctdf.columns:
            self.xcoord.append(self.xaxis[-1])

        ## CREATE SUBTOTALS FOR EACH COLUMNS
        for i in range(0, len(self.visible_xaxis_summary)):
            subtotal = source_dataframe.groupby(xaxis[:i+1] + yaxis)
            subtotal = subtotal.sum()[zaxis].unstack(xaxis[:i+1])

            for col in subtotal.columns:
                scolumns = col + (col[-1],) * (len(xaxis) - len(col) + 1)
                ctdf[scolumns] = subtotal[col]
                self.xcoord.append(self.visible_xaxis_summary[i])
        ## END

        ## CREATE COLUMNS GRAND TOTAL
        for value in zaxis:
            total = source_dataframe.groupby(yaxis + xaxis[-1:]).sum()[value]
            keys = tuple([''] * len(xaxis))
            ctdf[(value,) + keys] = total.unstack(xaxis[-1:]).sum(axis=1)
            self.xcoord.append('')
        ## END

        ## REORDER AXIS 1 SO THAT AGGREGATES ARE THE LAST LEVEL
        order = range(1, len(xaxis) + 1) + [0]
        ct = ctdf.reorder_levels(order, axis=1)
        ## END

        ## SORT COLUMNS AND RETURN SORTED DF
        sorted_columns = self.__sort_axis(ct.columns,
                                          self.visible_xaxis_summary,
                                          self.xcoord)
        zas = self.zaxis * (len(sorted_columns) / len(self.zaxis))
        sorted_columns = map(lambda x: x[0][:-1] + (x[1],),
                             zip(sorted_columns, zas))
        return ct.reindex_axis(axis=1, labels=sorted_columns)
        ## END

    def __yaxis_summary(self, yaxis, df):
        for idx in df.index:
            self.ycoord.append(self.yaxis[-1])

        ## CREATE SUBTOTALS FOR EACH INDEX
        subtotals = []
        for i in range(len(self.visible_yaxis_summary)):
            for idx in set([x[:i+1] for x in df.index]):
                sindex = idx + (idx[-1],) * (len(yaxis) - len(idx))
                stotal = pd.DataFrame({sindex: df.ix[idx].sum()}).T
                subtotals.append(stotal)
                self.ycoord.append(self.visible_yaxis_summary[i])
        ## END

        ## CREATE INDEX GRAND TOTAL
        gindex = tuple([''] * len(yaxis))
        gtotal = pd.DataFrame({gindex: df.ix[:].sum()}).T
        subtotals.append(gtotal)
        self.ycoord.append('')
        ## END

        ## SORT THE INDEX AND RETURN SORTED DF
        df = pd.concat([df] + subtotals)
        sorted_index = self.__sort_axis(df.index,
                                        self.visible_yaxis_summary,
                                        self.ycoord)
        return df.reindex_axis(axis=0, labels=sorted_index)
        ## END

    def __sort_axis(self, axis, visible_axis, coord):
        sorter = []
        for i, idx in enumerate(axis):
            if coord[i] in visible_axis:
                nans = [np.NaN, ] * (visible_axis.index(coord[i]) + 1)
                nans[visible_axis.index(coord[i])] = 1
                sorter.append([x for x in roundrobin(idx, nans)])
            else:
                nans = (np.NaN,) * len(visible_axis)
                sorter.append([x for x in roundrobin(idx, nans)])

        lexsort = np.lexsort([x for x in reversed(zip(*sorter))])
        sorted_index = []
        for lx in lexsort:
            idx = zip(*axis)
            lex = tuple([idx[x][lx] for x in range(len(idx))])
            sorted_index.append(lex)
        return sorted_index
