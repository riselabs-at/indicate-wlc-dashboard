"""Luxembourg: D7.30. Product-level results per m² in 'Overview of results', metadata in 'A. Metadata'.

The reference area of the per-m² results is not stated. ASSUMPTION: m² GFA (total above + below ground).
Absolute values are therefore source value × GFA; the dashboard can switch the assumption to UFA.
"""
import pandas as pd
from common import Extract, num, find_source
from config import SOURCES, TYPE_OF, RSP_YEARS

COLS = {'A1-A3': ['A1-A3, fossil', 'A1-A3, biog'], 'A4': ['A4, fossil'], 'A5': ['A5, fossil'],
        'B4-B5': ['B4-B5, fossil'], 'C1-C4': ['C2, fossil', 'C3, fossil', 'C4, fossil', 'C3-C4, biog'], 'D': ['D, tot']}
META = ['proj', 'typ', 'structure', 'ufa', 'gfa_below', 'gfa_above', 'floors_below', 'floors_above', 'compactness',
        'wwr', 'gfa', 'capita', 'year', 'gwp_scope', 'energy_source', 'eprim', 'an', 'pv', 'element_scope',
        'heating', 'cooling', 'facade']
UNIT = 'kgCO2e/m² (area not stated; assumed GFA)'


def extract(folder):
    src = find_source(folder, SOURCES['LU']['pattern'])
    x = Extract('LU', src)
    meta = pd.read_excel(src, sheet_name='A. Metadata', header=None).iloc[3:, :22]
    meta.columns = META
    meta = meta[meta.proj.notna()]
    x.step('read', "Sheet 'A. Metadata': one row per project", len(meta))

    res = pd.read_excel(src, sheet_name='Overview of results')
    x.step('read', "Sheet 'Overview of results': one row per product", len(res))
    res = res[res.Project.notna() & (res.Project != 'Projet')]
    x.step('filter', 'Dropped blank rows and repeated French header rows ("Projet")', len(res))
    for m, cs in COLS.items():
        res[m] = res[cs].fillna(0).sum(axis=1)
    res['code'] = res['LEVELS Tier 3'].astype(str)
    res['mat'] = res['INDICATE Materials Tier 2'].fillna('0.0 Undefined materials').astype(str)
    x.step('map', 'Product columns summed into harmonised modules (fossil + biogenic; C1 not reported)', len(res))

    for _, r in meta.iterrows():
        p = r.proj
        sub = res[res.Project == p]
        bid = 'LU-' + p
        if sub.empty:
            x.issue('warning', bid, 'Project in metadata without product results', 'Skipped')
            continue
        gfa, ufa, yr = num(r.gfa), num(r.ufa), num(r.year)
        x.buildings.append(dict(
            building_id=bid, country='LU', source_id=p, national_typology=r.typ,
            harmonised_type=TYPE_OF[('LU', r.typ)], intervention='New construction (assumed)', renovation_level='',
            gfa_m2=gfa, ufa_m2=ufa, hfa_m2=None, rsp_years=RSP_YEARS, results_basis='per m² (assumed GFA)',
            structure=r.structure if isinstance(r.structure, str) else None,
            year=int(yr) if yr and yr > 1900 else None,
            energy_source=r.energy_source if isinstance(r.energy_source, str) else None,
            n_products=len(sub),
        ))
        for m in COLS:
            v = float(sub[m].sum())
            x.lcm.append(dict(building_id=bid, module=m, source_value=v, source_unit=UNIT,
                              value_kgco2e=v * gfa, source_columns=' + '.join(COLS[m])))
        for (code, ), g in sub.groupby(['code']):
            for m in COLS:
                v = float(g[m].sum())
                if v:
                    x.elements.append(dict(building_id=bid, classification='LU LEVELS Tier 3', code=code,
                                           group=code.split('.')[0] + '.' + code.split('.')[1] if code.count('.') else code,
                                           module=m, value_kgco2e_m2=v, area_basis='assumed GFA', quality='original'))
        for (mat, ), g in sub.groupby(['mat']):
            for m in COLS:
                v = float(g[m].sum())
                if v:
                    x.materials.append(dict(building_id=bid, material_group=mat, module=m, value_kgco2e_m2=v))
    x.step('aggregate', 'Products summed per project × module, per element code × module, per material × module',
           f'{len(x.buildings)} buildings / {len(x.elements)} element rows / {len(x.materials)} material rows')

    # elements must add up to the building totals
    el = pd.DataFrame(x.elements).groupby('building_id').value_kgco2e_m2.sum()
    tot = pd.DataFrame(x.lcm).groupby('building_id').source_value.sum()
    x.check('Element rows sum to building totals (max abs difference, kgCO2e/m²)', float((el - tot).abs().max()), 0.01)

    # code descriptions derived from the data: top-3 product categories by embodied GWP
    emb = [m for m in COLS if m != 'D']
    res['emb'] = res[emb].sum(axis=1)
    for code, g in res.groupby('code'):
        top = g.groupby('Category class Mat intern')['emb'].sum().sort_values(ascending=False).head(3)
        x.element_desc[('LU LEVELS Tier 3', code)] = '; '.join(map(str, top.index))

    x.issue('confirm', '', 'Reference area of the per-m² results not stated in the file', 'Assumed GFA total; switchable in the dashboard')
    x.issue('confirm', '', '"House" typology includes larger multi-unit residences (e.g. MI_Con_1, 35 occupants)', 'Mapped to Single-family house')
    x.issue('confirm', '', 'LEVELS Tier 3 codes have no names in the file', 'Group names inferred from product content; codes described by top-3 product categories')
    x.issue('info', '', 'No operational modules (B6/B7) reported', 'Luxembourg drops out of whole-life views')
    x.issue('info', '', 'Type of intervention not stated (construction years 2017–2028)', 'Treated as new construction')
    return x
