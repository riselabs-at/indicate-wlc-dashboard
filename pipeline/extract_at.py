"""Austria – D7.3, sheet 'Data overview'. Absolute kgCO2e per building, module and GWP sub-indicator; UFA only."""
import os
import pandas as pd
from common import Extract, num, find_source
from config import SOURCES, TYPE_OF, TYPOLOGIES, DEPTH

S = SOURCES['AT']
# source part -> (template module, covers, flag)
PARTS = {'A1A3': ('A1-A3', '', ''), 'A4': ('A4', '', ''), 'A5': ('A5', '', ''), 'B4': ('B4', '', 'B5 not reported'),
         'B5': ('B6', '', 'header "B5", description row "B6"; treated as B6 (to confirm)'),
         'C1': ('C1', '', ''), 'C2': ('C2', '', ''), 'C3C4': ('C3', 'C3-C4', 'C3 and C4 reported together')}
IND = {'GWPtotal': 'GWP total', 'GWPfossil': 'GWP fossil', 'GWPbio': 'GWP biogenic', 'GWPluluc': 'GWP LULUC'}


def extract(folder):
    src = find_source(folder, S['pattern'])
    x = Extract('AT', src)
    df = pd.read_excel(src, sheet_name='Data overview', header=0, skiprows=[1])
    lcia = str(pd.read_excel(src, sheet_name='Data overview', header=None, nrows=2).iat[1, 5])
    x.step('read', "Sheet 'Data overview': header row 1, description row 2 (kept for the LCIA method)", len(df))
    df = df[df['name'].notna()]

    for _, r in df.iterrows():
        sid = str(r['name']).zfill(3)
        bid = 'AT-' + sid
        code = TYPE_OF[('AT', r['typology'])]
        x.projects.append(dict(
            building_id=bid, contact_org=S['org'], project_code=sid,
            project_intervention_type='New construction', project_renovation_depth=DEPTH['-'],
            x_intervention_basis='not stated in source; assumed new construction',
            site_country=S['country'], site_country_iso3=S['iso3'],
            bldg_use_subtype=TYPOLOGIES[code], x_typology_national=r['typology'],
            bldg_area_nfa=num(r['UFA']), x_nfa_basis='UFA (source: "UFA Total (m²) with filled gaps")',
            bldg_area_references='NFA = UFA total, with filled gaps (source header)',
            x_storeys_total=num(r['Unnamed: 3']),
            lca_rsp=num(r['Unnamed: 4']), x_rsp_basis='reported',
            lca_lcia_method='EF v3.1 (EN 15804), GWP100' if 'EF v3.1' in lcia else None,
            x_results_basis='absolute kgCO2e', x_source_deliverable=S['deliverable'], x_source_file=os.path.basename(src),
            x_source_id=sid,
        ))
        for ind, prefix in IND.items():
            for part, (mod, covers, flag) in PARTS.items():
                col = f'{prefix} - {part}'
                v = num(r[col])
                x.result(bid, ind, mod, v, v, 'kgCO2e', col, covers, flag)
    x.step('map', 'GWP total / fossil / biogenic / LULUC × 8 source modules → template modules (C3C4 → C3 covering C3-C4; "B5" → B6)', len(x.results))

    t = pd.DataFrame(x.results)
    tot = t[t.indicator == 'GWPtotal'].value_kgco2e.sum()
    parts = t[t.indicator != 'GWPtotal'].value_kgco2e.sum()
    x.check('GWP total = fossil + biogenic + LULUC (relative difference, all values)', (tot - parts) / tot, 0.001)

    x.issue('info', '', 'Column header "GWP total - B5", but the description row says B6.',
            'Stored as B6 (operational energy). Confirmed by D3.2: its AT operational benchmarks (B6-B7) equal these values.',
            'Please correct the column header to B6.')
    x.issue('info', '', 'B4 reported without B5; C3 and C4 reported together.', 'B4 stored as B4; C3+C4 stored in C3 with covers = C3-C4.')
    x.issue('confirm', '', 'No GFA reported; storeys only as a total. D3.2 (Table 1, Fig. 2) reports AT GFA (mean 3,547 m², median 1,699 m²) '
            'and floors above ground (mean 3.4), so these data exist.', 'Austria drops out of per-GFA views; total storeys kept as x_storeys_total.',
            'Please provide GFA and floors above / below ground per building (as used in D3.2).')
    x.issue('confirm', '', 'Mean UFA in this file is 2,470 m²; D3.2 Table 1 gives 2,241 m² for the same 67 buildings '
            '(per-m² results match D3.2 exactly, so the UFA used for normalisation is the same).', 'UFA as in the file ("with filled gaps").',
            'Which UFA values does D3.2 Table 1 use?')
    x.issue('info', '', 'Type of intervention not stated.', 'Treated as new construction.', 'Are all 67 buildings new construction?')
    return x
