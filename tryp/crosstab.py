import pandas as pd
import numpy as np
from util import roundrobin


class Crosstab(object):
    """Crosstab object crosstabulate a pandas dataframe object.

    Parameters:
    xaxis : list of dataframe columns to be referred as xaxis
    yaxis : list of dataframe index to be referred as yaxis
    zaxis : list of dataframe values to be referred as zaxis
    xaxis_total : list of dataframe columns that should be
                  totaled
    yaxis_total : list of dataframe index that should be
                  totaled
    source_dataframe : pandas dataframe object to be
                       crosstabulated

    For a more detailed explanation please refer to crosstab.ipynb
    inside notebooks folder.
    """
    def __init__(self, xaxis, yaxis, zaxis, xaxis_total, yaxis_total,
                 source_dataframe):
        self.coordinates = {}
        self.xaxis = xaxis
        self.yaxis = yaxis
        self.zaxis = zaxis
        self.xaxis_total = xaxis_total
        self.yaxis_total = yaxis_total
        df = source_dataframe.groupby(xaxis + yaxis).sum()
        df = df[zaxis].unstack(xaxis)
        if xaxis:
            df = self._xaxis_total(source_dataframe, xaxis, yaxis, zaxis, df)
        self.dataframe = self._yaxis_total(yaxis, df)

    def _xaxis_total(self, source_dataframe, xaxis, yaxis, zaxis, ctdf):
        coordinates = []
        for _ in ctdf.columns:
            coordinates.append(self.xaxis[-1])

        ## CREATE SUBTOTALS FOR EACH COLUMNS
        for i in range(0, len(self.xaxis_total)):
            subtotal = source_dataframe.groupby(xaxis[:i+1] + yaxis)
            subtotal = subtotal.sum()[zaxis].unstack(xaxis[:i+1])

            for col in subtotal.columns:
                scolumns = col + (col[-1],) * (len(xaxis) - len(col) + 1)
                ctdf[scolumns] = subtotal[col]
                coordinates.append(self.xaxis_total[i])
        ## END

        ## CREATE COLUMNS GRAND TOTAL
        for value in zaxis:
            total = source_dataframe.groupby(yaxis + xaxis[-1:]).sum()[value]
            keys = tuple([''] * len(xaxis))
            ctdf[(value,) + keys] = total.unstack(xaxis[-1:]).sum(axis=1)
            coordinates.append('')
        ## END

        ## REORDER AXIS 1 SO THAT AGGREGATES ARE THE LAST LEVEL
        order = range(1, len(xaxis) + 1) + [0]
        ct = ctdf.reorder_levels(order, axis=1)
        ## END

        ## SORT COLUMNS AND RETURN SORTED DF
        sorted_columns = self._sort_axis(ct.columns,
                                         self.xaxis_total,
                                         coordinates, 'x')
        zas = self.zaxis * (len(sorted_columns) / len(self.zaxis))
        self.coordinates['z'] = zas
        sorted_columns = map(lambda x: x[0][:-1] + (x[1],),
                             zip(sorted_columns, zas))
        return ct.reindex_axis(axis=1, labels=sorted_columns)
        ## END

    def _yaxis_total(self, yaxis, df):
        coordinates = []
        for _ in df.index:
            coordinates.append(self.yaxis[-1])

        ## CREATE SUBTOTALS FOR EACH INDEX
        subtotals = []
        for i in range(len(self.yaxis_total)):
            for idx in set([x[:i+1] for x in df.index]):
                sindex = idx + (idx[-1],) * (len(yaxis) - len(idx))
                stotal = pd.DataFrame({sindex: df.ix[idx].sum()}).T
                subtotals.append(stotal)
                coordinates.append(self.yaxis_total[i])
        ## END

        ## CREATE INDEX GRAND TOTAL
        gindex = tuple([''] * len(yaxis))
        gtotal = pd.DataFrame({gindex: df.ix[:].sum()}).T
        subtotals.append(gtotal)
        coordinates.append('')
        ## END

        ## SORT THE INDEX AND RETURN SORTED DF
        df = pd.concat([df] + subtotals)
        sorted_index = self._sort_axis(df.index,
                                       self.yaxis_total,
                                       coordinates, 'y')
        return df.reindex_axis(axis=0, labels=sorted_index)
        ## END

    def _sort_axis(self, labels, visible_axis, coordinates, axis):
        sorter = []
        for i, l in enumerate(labels):
            if coordinates[i] in visible_axis:
                nans = [np.NaN, ] * (visible_axis.index(coordinates[i]) + 1)
                nans[visible_axis.index(coordinates[i])] = 1
                sorter.append([x for x in roundrobin(l, nans)])
            else:
                nans = (np.NaN,) * len(visible_axis)
                sorter.append([x for x in roundrobin(l, nans)])

        sorted_index = []
        labels = zip(*labels)
        lexsort = np.lexsort([x for x in reversed(zip(*sorter))])
        self.coordinates[axis] = []
        for lx in lexsort:
            lex = tuple([labels[x][lx] for x in range(len(labels))])
            sorted_index.append(lex)
            self.coordinates[axis].append(coordinates[lx])
        return sorted_index
