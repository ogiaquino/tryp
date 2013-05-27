import pandas as pd
import numpy as np


def computed_values(tryp):
    crosstab = tryp.crosstab
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

    sorted_keys = ['!',
                   'BISCUITS',
                   'SNACKS',
                   'CHEESE',
                   'GROCERY',
                   'CONFECTIONERY']
    values = ('Target', 'Sell Out Actual', 'Ach')

    sorted_crosstab = pd.DataFrame(index=crosstab.index)
    for key in sorted_keys:
        for val in values:
            s = crosstab[(key,) + (val,)]
            sorted_crosstab[(key,) + (val,)] = s

    return sorted_crosstab
