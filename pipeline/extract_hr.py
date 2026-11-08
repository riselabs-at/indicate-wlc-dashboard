"""Croatia – D7.12, sheet 'Data_overview'. Absolute kgCO2e per building and module (GWP total only); GFA and UFA."""
import os
import pandas as pd
from common import Extract, num, find_source
from config import SOURCES, TYPE_OF, TYPOLOGIES, DEPTH, RSP_YEARS

S = SOURCES['HR']
COL = dict(stage=0, typology=1, case=2, total=3, gwp_gfa=4, gwp_gfa_y=5, gfa=9, ufa=10)
# source column -> (template module, covers, flag)
MOD_COL = {12: ('A1-A3', '', ''), 13: ('A4', '', ''), 14: ('A5', '', ''), 15: ('B3', '', ''),
           16: ('B4', 'B4-B5', 'B4 and B5 reported together'), 17: ('B6', '', ''), 18: ('B7', '', ''),
           19: ('C1', 'C1-C4', 'C1-C4 reported together'), 20: ('D1', 'D1-D2', 'D reported as one value')}


def extract(folder):
    src = find_source(folder, S['pattern'])
    x = Extract('HR', src)
    raw = pd.read_excel(src, sheet_name='Data_overview', header=None)
    df = raw.iloc[5:]
    df = df[df[COL['case']].notna() & df[COL['typology']].notna()]
    x.step('read', "Sheet 'Data_overview': data from row 6; absolute results block (columns M–U)", len(df))
    worst, worst_y = 0, 0
    for _, r in df.iterrows():
        sid = str(r[COL['case']])
        bid = 'HR-' + sid
        code = TYPE_OF[('HR', r[COL['typology']])]
        rsp = num(r[COL['gwp_gfa']]) / num(r[COL['gwp_gfa_y']])
        worst_y = max(worst_y, abs(rsp - RSP_YEARS))
        x.projects.append(dict(
            building_id=bid, contact_org=S['org'], project_code=sid,
            project_intervention_type=str(r[COL['stage']]).capitalize(), project_renovation_depth=DEPTH['-'],
            x_intervention_basis='reported', site_country=S['country'], site_country_iso3=S['iso3'],
            bldg_use_subtype=TYPOLOGIES[code], x_typology_national=r[COL['typology']],
            bldg_area_gfa=num(r[COL['gfa']]), bldg_area_nfa=num(r[COL['ufa']]), x_nfa_basis='UFA',
            bldg_area_references='GFA and UFA as reported (column headers)',
            lca_rsp=RSP_YEARS, x_rsp_basis='derived: GWPtot/GFA ÷ GWPtot/GFA/year = 50',
            x_results_basis='absolute kgCO2e', x_source_deliverable=S['deliverable'], x_source_file=os.path.basename(src), x_source_id=sid,
        ))
        s = 0
        for c, (mod, covers, flag) in MOD_COL.items():
            v = num(r[c])
            x.result(bid, 'GWPtotal', mod, v, v, 'kgCO2e', str(raw.iat[4, c]).strip(), covers, flag)
            s += v if mod != 'D1' else 0
        worst = max(worst, abs(s - num(r[COL['total']])))
    x.step('map', 'Absolute block → template modules (B4-B5 → B4, C1-C4 → C1, D → D1, each flagged)', len(x.results))
    x.check('Sum of A–C modules vs source "GWP total" (max abs difference, kgCO2e)', worst, 5)
    x.check('RSP derived from per-year column = 50 years (max abs difference)', worst_y, 0.01)

    bad = raw.iloc[5:5 + len(df), 26]
    if (pd.to_numeric(bad, errors='coerce').diff().dropna() == 1).all():
        x.issue('info', '', 'Per-GFA B3 column of the source holds a running index (0, 1, 2…).', 'Not used; per-m² values are calculated from the absolute block.')
    areas = df.groupby([COL['typology'], COL['gfa']]).size()
    x.issue('info', '', f'The {len(df)} cases are design variants of {len(areas)} base buildings (identical floor areas).',
            'Kept as separate records; percentiles reflect variant spread.', 'Which variant (if any) represents each base building as built?')
    x.issue('info', '', 'B4-B5, C1-C4 and D reported as combined values.', 'Stored in the first module of each range with covers flag.',
            'Can B4 / B5, C1–C4 and D1 / D2 be reported separately?')
    x.issue('confirm', '', 'The D7.12 file has results per module only. D3.2 uses Croatian element and material results (Fig. 10, 11), '
            'material masses (Fig. 4), structural systems (Fig. 3) and floors (mean 2.6).', 'No element, material or floor data for HR.',
            'Please provide the element (EU Tier 2/3) and material results and the building metadata used in D3.2.')
    return x
