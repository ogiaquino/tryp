import pandas as pd
import numpy as np


def extend(dataset):
    crosstab = dataset.crosstab
    new_columns = []
    for cc in crosstab.columns:
        if cc[:-1] not in new_columns:
            new_columns.append(cc[:-1])

    for nc in new_columns:
        soa = crosstab[nc + ('Sell Out Actual',)]
        target = crosstab[nc + ('Target',)]
        ach = soa / target * 100
        ach = ach.replace(np.inf, 0.00)
        ach = ach.replace(-np.inf, 0.00)
        crosstab[nc + ('Ach',)] = \
            pd.DataFrame(ach.round(1))

    sorted_keys = ['',
                   'BISCUITS',
                   'SNACKS',
                   'CHEESE',
                   'GROCERY',
                   'CONFECTIONERY']

    crosstab_columns = []
    crosstab_series = []

    for key in sorted_keys:
        for val in dataset.values:
            s = crosstab[(key, val)]
            crosstab_series.append(s)
            crosstab_columns.append((key, val))

    index = crosstab.index
    columns = pd.MultiIndex.from_tuples(crosstab_columns)
    crosstab = pd.DataFrame(zip(*crosstab_series),
                            index=index,
                            columns=columns)

    return crosstab
