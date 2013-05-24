import pandas as pd
import numpy as np


def crosstab(rows, columns, values, df, computed_values=None, result_level=2):
    df_groupby_sum = df.groupby(rows + columns).sum()
    crosstab = df_groupby_sum[values].unstack(columns)

    for i in range(1, len(columns)):
        total_groups = rows + columns[:i]
        crosstab_totals = df.groupby(total_groups).sum()[values]. \
            unstack(columns[:i])

        for col in crosstab_totals.columns:
            label_length = (len(columns) - len(col)) + 1
            total_index = ['!' + col[-1] + ' Result'
                           for x in range(label_length)]
            crosstab[col + tuple(total_index)] = crosstab_totals[col]

    if columns:
        for value in values:
            df_groupby_sum = df.groupby(rows + columns[-1:]).sum()
            key = (value,) + tuple(['!' for x in range(len(columns))])
            crosstab[key] = df_groupby_sum[value].unstack(columns[-1:]).sum(axis=1)

        order = [x for x in range(1, len(columns)+1)] + [0]
        crosstab = crosstab.reorder_levels(order, axis=1).sort_index(axis=1)

        crosstab_columns = []
        for key in sorted(set([k[:-1] for k in crosstab.keys()])):
            for val in values:
                crosstab_columns.append(key + (val,))

        crosstab = pd.DataFrame(crosstab, columns=crosstab_columns)
    crosstab_row_subtotals = []
    if len(rows) > 1:
        crosstab_rows = []
        # FOR KRAFT AND OTHER THAT FOLLOWS THEIR REPORT
        # STANDART x[:-2] SHOULD BE USED HERE:
        # for row in set([x[:-2] for x in crosstab.index]):
        # TO HAVE THE LAST TWO ROW LEVEL ON THE SAME ROW AND NOT SHOW
        # ANY AGGREGRATION NORMALLY THESE ARE SR CODE/NAME
        # OR SKU/PRODUCT DESCRIPTION. NEED TO FIND OF BETTER
        # WAY TO HANDLE THIS.
        for row in set([x[:-int(result_level)] for x in crosstab.index]):
            for i in range(len(rows)):
                if row[:i+1] not in crosstab_rows:
                    crosstab_rows.append(row[:i+1])

        for row in crosstab_rows:
            result = tuple(['!' + row[-1] + ' Result'
                           for x in range(len(rows) - len(row))])
            index = row + result
            row_df = pd.DataFrame({index: crosstab.ix[row].sum()}).T
            crosstab_row_subtotals.append(row_df)
        total = {tuple(['!' for x in range(len(rows))]): crosstab.ix[:].sum()}
        total_df = pd.DataFrame(total).T
        crosstab_row_subtotals.append(total_df)
    else:
        total_df = pd.DataFrame({'!': crosstab.ix[:].sum()}).T
        crosstab_row_subtotals.append(total_df)

    crosstab = pd.concat([crosstab] + crosstab_row_subtotals)
    crosstab = crosstab.sort_index(axis=0)
    return crosstab
