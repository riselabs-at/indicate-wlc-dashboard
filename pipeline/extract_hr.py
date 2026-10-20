"""Croatia: D7.12, sheet 'Data_overview'. Absolute kgCO2e per building and module, GFA and UFA."""
import pandas as pd
from common import Extract, num, find_source
from config import SOURCES, TYPE_OF, RSP_YEARS

# column positions in 'Data_overview' (header on row 5 of the sheet)
COL = dict(stage=0, typology=1, case=2, total=3, gfa=9, ufa=10)
MOD_COL = {'A1-A3': 12, 'A4': 13, 'A5': 14, 'B2-B3': 15, 'B4-B5': 16, 'B6': 17, 'B7': 18, 'C1-C4': 19, 'D': 20}
NAT = {'SINGLE FAMILY HOUSE': 'Single family house', 'MULTY FAMILY HOUSE': 'Multi family house',
       'OFFICE BUILDING': 'Office building', 'EDUCATION': 'Education', 'HOSPITAL': 'Hospital', 'HOTEL': 'Hotel',
       'SPORT HALL BUILDING': 'Sport hall', 'NEW BUILDING FOR RETAIL AND WHOLESALE': 'Retail and wholesale'}


def extract(folder):
    src = find_source(folder, SOURCES['HR']['pattern'])
    x = Extract('HR', src)
    raw = pd.read_excel(src, sheet_name='Data_overview', header=None)
    df = raw.iloc[5:]
    df = df[df[COL['case']].notna() & df[COL['typology']].notna()]
    x.step('read', "Sheet 'Data_overview': data from row 6; absolute results block (cols M–U)", len(df))

    max_diff = 0
    for _, r in df.iterrows():
        bid = 'HR-' + str(r[COL['case']])
        typ = r[COL['typology']]
        x.buildings.append(dict(
            building_id=bid, country='HR', source_id=str(r[COL['case']]),
            national_typology=NAT[typ], harmonised_type=TYPE_OF[('HR', typ)],
            intervention=str(r[COL['stage']]).capitalize(), renovation_level='',
            gfa_m2=num(r[COL['gfa']]), ufa_m2=num(r[COL['ufa']]), hfa_m2=None,
            rsp_years=RSP_YEARS, results_basis='absolute kgCO2e',
        ))
        s = 0
        for m, c in MOD_COL.items():
            v = num(r[c])
            x.lcm.append(dict(building_id=bid, module=m, source_value=v, source_unit='kgCO2e',
                              value_kgco2e=v, source_columns=str(raw.iat[4, c]).strip()))
            if m != 'D':
                s += v
        max_diff = max(max_diff, abs(s - num(r[COL['total']])))
    x.step('map', 'Modules mapped 1:1 (B3→B2-B3); D kept separately', len(x.lcm))
    x.check('Sum of A–C modules vs source "GWP total" (max abs difference, kgCO2e)', max_diff, 5)

    # The per-GFA block of the source has a running index in its B3 column -> not used.
    bad = raw.iloc[5:5 + len(df), 26]
    if (pd.to_numeric(bad, errors='coerce').diff().dropna() == 1).all():
        x.issue('info', '', 'Per-GFA B3 column in the source holds a running index (0, 1, 2…)',
                'Per-m² values recomputed from the absolute block')
    # design variants: identical floor areas
    areas = df.groupby([COL['typology'], COL['gfa']]).size()
    x.issue('info', '', f'The {len(df)} cases are design variants of {len(areas)} base buildings (identical floor areas)',
            'Kept as separate records; percentiles reflect variant spread')
    return x
