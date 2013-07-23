import imp
import pandas as pd
import numpy as np

from collections import OrderedDict
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
                                 meta.values,
                                 meta.index_totals)
        self._extend(meta.extmodule)

    def to_excel(self):
        to_excel(self)

    def _extend(self, extmodule):
        if extmodule:
            extmodule = imp.load_source(extmodule[0], extmodule[1])
            extmodule.extend(self)
        self.values_labels = self._values_labels(self.df)

    def _crosstab(self, df, index, columns, values, index_totals):
        ctdf = df.groupby(index + columns).sum()[values].unstack(columns)
        if columns:
            ctdf = self._columns_totals(df, index, columns, values, ctdf)
        ctdf = self._index_totals(index, index_totals, ctdf)
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

    def _index_totals(self, index, index_totals, df):
        df_index_subtotals = []
        if len(index) > 1:
            df_index = []
            for idx in set([x[:-1] for x in df.index]):
                for i in range(len(index)):
                    if idx[:i+1] not in df_index:
                        df_index.append(idx[:i+1])

            for idx in df_index:
                result = tuple(['!' + idx[-1]
                               for x in range(len(index) - len(idx))])
                index_dict = dict([(r, len(index) - i - 1) for i, r in
                                  enumerate(index)])

                if len(result) in [index_dict[r] for r in index_totals]:
                    idxr = idx + result
                    index_df = pd.DataFrame({idxr: df.ix[idx].sum()}).T
                    df_index_subtotals.append(index_df)

            total = {tuple(['!'] * len(index)): df.ix[:].sum()}
            total_df = pd.DataFrame(total).T
            df_index_subtotals.append(total_df)
        else:
            total_df = pd.DataFrame({'!': ct.ix[:].sum()}).T
            ct_index_subtotals.append(total_df)

        df = pd.concat([df] + df_index_subtotals)
        df = df.sort_index(axis=0)

        self.get_columns_axis(df)
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
