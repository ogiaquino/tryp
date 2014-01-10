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
    dataframe : pandas dataframe object to be
                crosstabulated
    """
    def __init__(self, xaxis, yaxis, zaxis, xaxis_total, yaxis_total,
                 dataframe):
        self.coordinates = {}
        self.xaxis = xaxis
        self.yaxis = yaxis
        self.zaxis = zaxis
        self.xaxis_total = xaxis_total
        self.yaxis_total = yaxis_total
        self.dataframe = dataframe

        if xaxis:
            df = self._xaxis_total()

        # At this stage dataframe is crosstabulated
        self.dataframe = self._yaxis_total(df)

    def _xaxis_total(self):
        coordinates = []
        xaxis = self.xaxis
        yaxis = self.yaxis
        zaxis = self.zaxis
        xaxis_total = self.xaxis_total
        yaxis_total = self.yaxis_total
        dataframe = self.dataframe

        # UNSTACKED DATAFRAME
        udf = dataframe.groupby(xaxis + yaxis).sum() \
                [zaxis].unstack(xaxis)

        for _ in udf.columns:
            coordinates.append(xaxis[-1])

        # CREATE SUBTOTALS FOR EACH COLUMNS
        for i in range(0, len(xaxis_total)):
            subtotal = dataframe.groupby(xaxis[:i+1] + yaxis)
            subtotal = subtotal.sum()[zaxis].unstack(xaxis[:i+1])

            for col in subtotal.columns:
                scolumns = col + (col[-1],) * (len(xaxis) - len(col) + 1)
                udf[scolumns] = subtotal[col]
                coordinates.append(xaxis_total[i])
        # END

        # CREATE COLUMNS GRAND TOTAL
        for value in zaxis:
            total = dataframe.groupby(yaxis + xaxis[-1:]).sum()[value]
            keys = tuple([''] * len(xaxis))
            udf[(value,) + keys] = total.unstack(xaxis[-1:]).sum(axis=1)
            coordinates.append('')
        # END

        # REORDER AXIS 1 SO THAT AGGREGATES ARE THE LAST LEVEL
        order = range(1, len(xaxis) + 1) + [0]
        ct = udf.reorder_levels(order, axis=1)
        # END

        # SORT COLUMNS AND RETURN SORTED DF
        sorted_columns = self._sort_axis(ct.columns,
                                         xaxis_total,
                                         coordinates, 'x')
        zas = zaxis * (len(sorted_columns) / len(zaxis))
        self.coordinates['z'] = zas
        sorted_columns = map(lambda x: x[0][:-1] + (x[1],),
                             zip(sorted_columns, zas))
        return ct.reindex_axis(axis=1, labels=sorted_columns)
        # END

    def _yaxis_total(self, df):
        coordinates = []
        yaxis = self.yaxis
        for _ in df.index:
            coordinates.append(self.yaxis[-1])

        # CREATE SUBTOTALS FOR EACH INDEX
        subtotals = []
        for i in range(len(self.yaxis_total)):
            for idx in set([x[:i+1] for x in df.index]):
                sindex = idx + (idx[-1],) * (len(yaxis) - len(idx))
                stotal = pd.DataFrame({sindex: df.ix[idx].sum()}).T
                subtotals.append(stotal)
                coordinates.append(self.yaxis_total[i])
        # END

        # CREATE INDEX GRAND TOTAL
        gindex = tuple([''] * len(yaxis))
        gtotal = pd.DataFrame({gindex: df.ix[:].sum()}).T
        subtotals.append(gtotal)
        coordinates.append('')
        # END

        # SORT THE INDEX AND RETURN SORTED DF
        df = pd.concat([df] + subtotals)
        sorted_index = self._sort_axis(df.index,
                                       self.yaxis_total,
                                       coordinates, 'y')
        return df.reindex_axis(axis=0, labels=sorted_index)
        # END

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
