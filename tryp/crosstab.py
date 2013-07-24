import imp
import pandas as pd
import numpy as np

from itertools import cycle, islice
from collections import OrderedDict
from excel import to_excel as to_excel


def roundrobin(*iterables):
    "roundrobin('ABC', 'D', 'EF') --> A D E B F C"
    # Recipe credited to George Sakkis
    pending = len(iterables)
    nexts = cycle(iter(it).next for it in iterables)
    while pending:
        try:
            for next in nexts:
                yield next()
        except StopIteration:
            pending -= 1
            nexts = cycle(islice(nexts, pending))

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
                                 meta.values,
                                )
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
        ctdf = self._index_totals(index, ctdf)
        return self._rename(ctdf)

    def _values_labels(self, ct):
        if isinstance(ct.columns, pd.MultiIndex):
            return map(lambda x: x[-1], ct.columns)
        return ct.columns

    def _columns_totals(self, df, index, columns, values, ct):
        ## CREATE SUBTOTALS FOR EACH COLUMNS
        for i in range(1, len(columns)):
            subtotal = df.groupby(index + columns[:i]).sum()[values]. \
                unstack(columns[:i])

            for col in subtotal.columns:
                ext = []
                for _ in range(len(col), len(columns) + 1):
                    ext.append('!' + col[-1])
                ct[col + tuple(ext)] = subtotal[col]
        ## END

        ## CREATE GRAND TOTAL
        for value in values:
            total = df.groupby(index + columns[-1:]).sum()
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

    def _index_totals(self, index, df):
        sorter = []
        subtotals = []

        nans = (np.NaN,) * len(self.index_totals)
        sorter = [[x for x in roundrobin(idx, nans)]  for idx in df.index]

        for i in range(len(self.index_totals)):
            for idx in set([x[:i+1] for x in df.index]):
                sindex = idx + (idx[-1],) * (len(index) - len(idx))
                stotal = pd.DataFrame({sindex: df.ix[idx].sum()}).T
                subtotals.append(stotal)

                rank = [np.NaN] * len(self.index_totals)                        
                rank[len(idx) - 1] = 1                                     
                sorter.append([x for x in roundrobin(sindex, rank)])

        gindex = tuple([''] * len(index))
        gtotal = pd.DataFrame({gindex: df.ix[:].sum()}).T
        subtotals.append(gtotal)

        rank = [1] * len(self.index_totals)                        
        sorter.append([x for x in roundrobin(gindex, tuple(rank))])

        df = pd.concat([df] + subtotals)

        sorter = zip(*sorter)
        lexsort = np.lexsort([x for x in reversed(sorter)])

        sorted_index=[]
        for lx in lexsort:
            index = zip(*df.index)
            lex = tuple([index[x][lx] for x in range(len(index))])
            sorted_index.append(lex)

        df = pd.DataFrame(df, index=sorted_index)
        return df

    def _rename(self, df):
        if isinstance(df.columns, pd.MultiIndex):
            col = map(lambda column: tuple([c.replace('!', '')
                                            for c in column]),
                      df.columns)
            col = pd.MultiIndex.from_tuples(col, names=df.columns.names)
        elif isinstance(df.columns, pd.Index):
            col = [c.replace('!', '') for c in df.columns]
            col = pd.Index(col)

        if isinstance(df.index, pd.MultiIndex):
            idx = map(lambda index: tuple([i.replace('!', '') for i in index]),
                      df.index)
            idx = pd.MultiIndex.from_tuples(idx, names=df.index.names)
        elif isintance(df.index, pd.Index):
            idx = [c.replace('!', '') for c in df.index]
            idx = pd.Index(col)

        renamed = pd.DataFrame(df.values, index=idx, columns=col)
        return renamed

    def get_index_axis(self, df):
        subs = []
        for idx in df.index:
            if subtotal:
                sub = self.levels.index[:-subtotal]
                sub = sub + sub[-1:] * (len(self.levels.index) - len(sub))
                subs.append(sub)
            else:
                subs.append(self.levels.index)
        return subs

    def get_columns_axis(self, df):
        d = OrderedDict()

        for idx in map(lambda x: x[:-1], df.columns):
            if subtotal:
                sub = self.levels.columns[:-subtotal]
                sub = sub + sub[-1:] * (len(self.levels.columns) - len(sub))
                d[tuple([x.replace('!','') for x in idx])] = sub
            else:
                d[idx] = self.levels.columns
        return d
