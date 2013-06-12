import pandas as pd
import numpy as np


def computed_values(tryp):
    crosstab = tryp.crosstab
    new_columns = []
    for cc in crosstab.columns:
        if cc[:-1] not in new_columns:
            new_columns.append(cc[:-1])
    for nc in new_columns:
        doscv = crosstab[nc + ('SOH Value',)] / crosstab[nc + ('Sales Value',)] * 100
        doscv = doscv.replace(np.inf, 0.00)
        doscv = doscv.replace(-np.inf, 0.00)
        crosstab[nc + ('DOSC Value',)] = pd.DataFrame(doscv.round(1))

    new_columns = []
    for cc in crosstab.columns:
        if cc[:-1] not in new_columns:
            new_columns.append(cc[:-1])
    for nc in new_columns:
        doscq = crosstab[nc + ('SOH Qty',)] / crosstab[nc + ('Sales Qty',)] * 100
        doscq = doscq.replace(np.inf, 0.00)
        doscq = doscq.replace(-np.inf, 0.00)
        crosstab[nc + ('DOSC Qty',)] = pd.DataFrame(doscq.round(1))

    new_columns = []
    for cc in crosstab.columns:
        if cc[:-1] not in new_columns:
            new_columns.append(cc[:-1])
    for nc in new_columns:
        doscq = crosstab[nc + ('SOH Volume',)] / crosstab[nc + ('Sales Volume',)] * 100
        doscq = doscq.replace(np.inf, 0.00)
        doscq = doscq.replace(-np.inf, 0.00)
        crosstab[nc + ('DOSC Volume',)] = pd.DataFrame(doscq.round(1))

    values=["Sales Value","SOH Value","DOSC Value", "Sales Qty","SOH Qty","DOSC Qty", "Sales Volume","SOH Volume","DOSC Volume"]
    crosstab_columns = []
    for key in sorted(set([k[:-1] for k in crosstab.keys()])):
        for val in values:
            crosstab_columns.append(key + (val,))

    crosstab = pd.DataFrame(crosstab, columns=crosstab_columns)
    return crosstab
