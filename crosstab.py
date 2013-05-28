import pandas as pd
import numpy as np


def crosstab(tryp):
    rows = tryp.rows
    columns = tryp.columns
    values = tryp.values
    computed_values = tryp.computed_values
    rows_results = tryp.rows_results
    df = tryp.df
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
        for row in set([x[:-1] for x in crosstab.index]):
            for i in range(len(rows)):
                if row[:i+1] not in crosstab_rows:
                    crosstab_rows.append(row[:i+1])

        for row in crosstab_rows:
            result = tuple(['!' + row[-1] + ' Result'
                           for x in range(len(rows) - len(row))])
            rows_dict = dict([(r, len(tryp.rows) - i - 1) for i, r in enumerate(tryp.rows)])
            
            if len(result) in [rows_dict[r] for r in rows_results]: 
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
