import pandas as pd
import numpy as np
from crosstab import crosstab as ct


def computed_values(crosstab, rows=None, columns=None, df=None, tryp=None):
    gross = ct(rows, columns, ['Gross Sales'], df)
    new_columns = []
    for cc in crosstab.columns:
        if cc[:-1] not in new_columns:
            new_columns.append(cc[:-1])
    for nc in new_columns:
        ach = crosstab[nc + ('BSR',)] / gross[nc + ('Gross Sales',)] * 100
        ach = ach.replace(np.inf, 0.00)
        ach = ach.replace(-np.inf, 0.00)
        crosstab[nc + ('percentage',)] = pd.DataFrame(ach.round(1))

    sorted_keys = ['!',
                   'BISCUITS',
                   'SNACKS',
                   'CHEESE',
                   'GROCERY',
                   'CONFECTIONERY']
    values = ('BSR', 'percentage')

    sorted_crosstab = pd.DataFrame(index=crosstab.index)

    for key in sorted_keys:
        s = crosstab[(key,) + ('BSR',)]
        sorted_crosstab[(key,) + ('BSR',)] = s
        s = crosstab[(key,) + ('percentage',)]
        sorted_crosstab[(key,) + ('percentage',)] = s

    return sorted_crosstab
