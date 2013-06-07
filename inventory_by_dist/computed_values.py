import pandas as pd
import numpy as np
import pandas.io.sql as psql

from time import gmtime, strftime

def computed_values(tryp):
    crosstab = tryp.crosstab
    wd = 75
    new_columns = []
    for cc in crosstab.columns:
        if cc not in new_columns:
            new_columns.append(cc)
    for nc in new_columns:
        ave_svrm = crosstab["Sales_rm"] / wd
        ave_svrm = ave_svrm.replace(np.inf, 0.00)
        ave_svrm = ave_svrm.replace(-np.inf, 0.00)
        crosstab['Average Sales RM'] = pd.DataFrame(ave_svrm)

    new_columns = []
    for cc in crosstab.columns:
        if cc not in new_columns:
            new_columns.append(cc)
    for nc in new_columns:
        dosc_rm = crosstab["SOH_rm"] / crosstab["Average Sales RM"] #* wd
        dosc_rm = dosc_rm.replace(np.inf, 0.00)
        dosc_rm = dosc_rm.replace(-np.inf, 0.00)
        crosstab['DOSC RM'] = pd.DataFrame(dosc_rm)

    new_columns = []
    for cc in crosstab.columns:
        if cc not in new_columns:
            new_columns.append(cc)
    for nc in new_columns:
        ave_svctn = crosstab["Sales_ctn"] / wd
        ave_svctn = ave_svctn.replace(np.inf, 0.00)
        ave_svctn= ave_svctn.replace(-np.inf, 0.00)
        crosstab['Average Sales CTN'] = pd.DataFrame(ave_svctn)

    new_columns = []
    for cc in crosstab.columns:
        if cc not in new_columns:
            new_columns.append(cc)
    for nc in new_columns:
        dosc_ctn = crosstab["SOH_ctn"] / crosstab["Average Sales CTN"] #* wd
        dosc_ctn = dosc_ctn.replace(np.inf, 0.00)
        dosc_ctn = dosc_ctn.replace(-np.inf, 0.00)
        crosstab['DOSC CTN'] = pd.DataFrame(dosc_ctn)

    new_columns = []
    for cc in crosstab.columns:
        if cc not in new_columns:
            new_columns.append(cc)
    for nc in new_columns:
        ave_svton = crosstab["Sales_volume"] / wd
        ave_svton = ave_svton.replace(np.inf, 0.00)
        ave_svton= ave_svton.replace(-np.inf, 0.00)
        crosstab['Average Sales TONNES'] = pd.DataFrame(ave_svton)

    new_columns = []
    for cc in crosstab.columns:
        if cc not in new_columns:
            new_columns.append(cc)
    for nc in new_columns:
        dosc_ton = crosstab["SOH_volume"] / crosstab["Average Sales TONNES"] #* wd
        dosc_ton = dosc_ton.replace(np.inf, 0.00)
        dosc_ton = dosc_ton.replace(-np.inf, 0.00)
        crosstab['DOSC TONNES'] = pd.DataFrame(dosc_ton)

    values=["Sales_rm","Average Sales RM","SOH_rm","DOSC RM"]
    values= values + ["Sales_ctn","Average Sales CTN","SOH_ctn","DOSC CTN"]
    values= values + ["Sales_volume","Average Sales TONNES","SOH_volume","DOSC TONNES"]
    crosstab_columns = []
    for val in values:
        crosstab_columns.append(val)
    crosstab = pd.DataFrame(crosstab, columns=crosstab_columns)
    return crosstab
