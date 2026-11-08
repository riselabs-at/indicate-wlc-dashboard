"""Luxembourg – D7.30. Product-level results per m² ('Overview of results'), project metadata ('A. Metadata').

The file does not state the reference area of the per-m² results. They are per m² UFA: the D3.2 synthesis report
(Table E.1, kgCO2e/m² UFA) reproduces the module statistics of these values exactly (checked 2026-11-06).
Absolute values = source value × UFA.
"""
import os
import pandas as pd
from common import Extract, num, find_source
from config import SOURCES, TYPE_OF, TYPOLOGIES, DEPTH

S = SOURCES['LU']
UNIT = 'kgCO2e/m² UFA (per D3.2; not stated in the file)'
META = ['proj', 'typ', 'structure', 'ufa', 'gfa_below', 'gfa_above', 'floors_below', 'floors_above', 'compactness',
        'wwr', 'gfa', 'capita', 'year', 'gwp_scope', 'energy_source', 'eprim', 'an', 'pv', 'element_scope',
        'heating', 'cooling', 'facade']
# product columns -> (indicator, template module, covers, flag)
COLMAP = [
    ('A1-A3, fossil', 'GWPfossil', 'A1-A3', '', 'fossil incl. LULUC'), ('A1-A3, biog', 'GWPbio', 'A1-A3', '', ''),
    ('A4, fossil', 'GWPfossil', 'A4', '', 'fossil incl. LULUC'), ('A5, fossil', 'GWPfossil', 'A5', '', 'fossil incl. LULUC'),
    ('B4-B5, fossil', 'GWPfossil', 'B4', 'B4-B5', 'B4 and B5 together; fossil incl. LULUC'),
    ('C2, fossil', 'GWPfossil', 'C2', '', 'fossil incl. LULUC'), ('C3, fossil', 'GWPfossil', 'C3', '', 'fossil incl. LULUC'),
    ('C4, fossil', 'GWPfossil', 'C4', '', 'fossil incl. LULUC'), ('C3-C4, biog', 'GWPbio', 'C3', 'C3-C4', 'C3 and C4 together'),
    ('D, tot', 'GWPtotal', 'D1', 'D1-D2', 'D reported as one value'),
]
# GWPtotal per template module = sum of these product columns
TOTAL = {'A1-A3': (['A1-A3, fossil', 'A1-A3, biog'], '', ''), 'A4': (['A4, fossil'], '', ''), 'A5': (['A5, fossil'], '', ''),
         'B4': (['B4-B5, fossil'], 'B4-B5', 'B4 and B5 together'), 'C2': (['C2, fossil'], '', ''),
         'C3': (['C3, fossil', 'C3-C4, biog'], '', 'includes the biogenic part of C3–C4'), 'C4': (['C4, fossil'], '', ''),
         'D1': (['D, tot'], 'D1-D2', 'D reported as one value')}


def extract(folder):
    src = find_source(folder, S['pattern'])
    x = Extract('LU', src)
    meta = pd.read_excel(src, sheet_name='A. Metadata', header=None).iloc[3:, :22]
    meta.columns = META
    meta = meta[meta.proj.notna()]
    x.step('read', "Sheet 'A. Metadata': one row per project", len(meta))
    res = pd.read_excel(src, sheet_name='Overview of results')
    x.step('read', "Sheet 'Overview of results': one row per product", len(res))
    res = res[res.Project.notna() & (res.Project != 'Projet')]
    x.step('filter', 'Dropped blank rows and repeated French header rows ("Projet")', len(res))
    cols = sorted({c for v in TOTAL.values() for c in v[0]})
    res[cols] = res[cols].fillna(0)
    res['code'] = res['LEVELS Tier 3'].astype(str)
    res['mat'] = res['INDICATE Materials Tier 2'].fillna('0.0 Undefined materials').astype(str)

    for _, r in meta.iterrows():
        p = r.proj
        sub = res[res.Project == p]
        bid = 'LU-' + p
        if sub.empty:
            x.issue('warning', bid, 'Project in metadata without product results.', 'Skipped.')
            continue
        gfa, ufa, yr, cap = num(r.gfa), num(r.ufa), num(r.year), num(r.capita)
        users_basis = None
        if cap is not None:
            users_basis = 'reported ("Capita")' + ('; equals UFA ÷ 10, likely estimated' if ufa and abs(cap - ufa / 10) < 0.01 else '')
        tcode = TYPE_OF[('LU', r.typ)]
        txt = lambda v: v if isinstance(v, str) and v.strip() and v.strip().lower() != 'none' else None
        x.projects.append(dict(
            building_id=bid, contact_org=S['org'], project_code=p,
            project_intervention_type='New construction', project_renovation_depth=DEPTH['-'],
            x_intervention_basis='not stated in source (years of construction 2017–2028); assumed new construction',
            site_country=S['country'], site_country_iso3=S['iso3'],
            bldg_use_subtype=TYPOLOGIES[tcode], x_typology_national=r.typ,
            bldg_year_complete=int(yr) if yr and yr > 1900 else None,
            bldg_users_total=cap, x_users_basis=users_basis,
            bldg_area_gfa=gfa, bldg_area_nfa=ufa, x_nfa_basis='UFA',
            bldg_area_references='GFA = below + above ground; NFA = UFA (A. Metadata). Per-m² results are per m² UFA (not stated in the file; confirmed by D3.2 Table E.1).',
            x_gfa_below_m2=num(r.gfa_below), x_gfa_above_m2=num(r.gfa_above),
            bldg_floors_ag=num(r.floors_above), bldg_floors_bg=num(r.floors_below),
            inv_energy_consumption_description=' / '.join(v for v in [txt(r.heating), txt(r.cooling)] if v) or None,
            x_energy_source_national=txt(r.energy_source), x_primary_energy_kwh_m2a=num(r.eprim), x_heated_area_an_m2=num(r.an),
            x_struct_material_national=txt(r.structure), x_compactness=num(r.compactness), x_window_wall_ratio=num(r.wwr),
            x_rsp_basis='not reported (product service lives only); 50 years assumed for per-year values',
            x_results_basis='per m² UFA (per D3.2)',
            x_source_deliverable=S['deliverable'], x_source_file=os.path.basename(src), x_source_id=p,
        ))
        for mod, (cs, covers, flag) in TOTAL.items():
            v = float(sub[cs].sum().sum())
            x.result(bid, 'GWPtotal', mod, v * ufa, v, UNIT, ' + '.join(cs), covers, flag)
        for col, ind, mod, covers, flag in COLMAP:
            if ind == 'GWPtotal':
                continue
            v = float(sub[col].sum())
            x.result(bid, ind, mod, v * ufa, v, UNIT, col, covers, flag)
        for code, g in sub.groupby('code'):
            for mod, (cs, _, _) in TOTAL.items():
                v = float(g[cs].sum().sum())
                if v:
                    x.elements.append(dict(building_id=bid, classification='LU LEVELS Tier 3', national_code=code, module=mod,
                                           value_kgco2e_m2=v, area_basis='UFA', quality='original', note='', source_cell=''))
        for mat, g in sub.groupby('mat'):
            for mod, (cs, _, _) in TOTAL.items():
                v = float(g[cs].sum().sum())
                if v:
                    x.materials.append(dict(building_id=bid, material_group=mat, module=mod, value_kgco2e_m2=v))
    x.step('aggregate', 'Products summed per project × template module, per element code and per material group (INDICATE Materials Tier 2 = template groups)',
           f'{len(x.projects)} buildings / {len(x.elements)} element rows / {len(x.materials)} material rows')
    x.step('convert', 'Absolute kgCO2e = per-m² value × UFA (reference area per D3.2 Table E.1)', '')

    el = pd.DataFrame(x.elements).groupby('building_id').value_kgco2e_m2.sum()
    t = pd.DataFrame(x.results)
    tot = t[t.indicator == 'GWPtotal'].groupby('building_id').source_value.sum()
    x.check('Element rows sum to building totals incl. D (max abs difference, kgCO2e/m²)', float((el - tot).abs().max()), 0.01)
    fb = t[t.indicator != 'GWPtotal'].groupby('building_id').source_value.sum()
    x.check('GWPtotal = fossil (+LULUC) + biogenic, excl. D (max abs difference, kgCO2e/m²)',
            float((t[(t.indicator == 'GWPtotal') & (t.module != 'D1')].groupby('building_id').source_value.sum() - fb).abs().max()), 0.01)

    res['emb'] = res[[c for c in cols if c != 'D, tot']].sum(axis=1)
    for code, g in res.groupby('code'):
        top = g.groupby('Category class Mat intern')['emb'].sum().sort_values(ascending=False).head(3)
        x.element_desc[('LU', code)] = '; '.join(map(str, top.index))

    x.issue('info', '', 'Reference area of the per-m² results not stated in the file.',
            'Treated as per m² UFA: D3.2 Table E.1 (per m² UFA) matches these values exactly for A1-A3, A4-A5, B1-B4 and C1-C4.',
            'Please confirm that the per-m² results are per m² UFA.')
    x.issue('confirm', '', '"House" includes larger multi-unit residences (e.g. MI_Con_1, 35 occupants).', 'Mapped to SFH (as in D3.2).',
            'Should the larger "House" projects be MFH instead of SFH?')
    x.issue('confirm', '', 'LEVELS Tier 3 codes have no names in the file.', 'Mapped to EU Table 4 tiers from the product content (inferred).',
            'Please provide the names of the LEVELS codes or confirm the mapping to EU Table 4.')
    x.issue('confirm', '', 'Structure given as main material only (concrete, wood, mixes, steel).', 'Kept as x_struct_material_national; bldg_struct_type left blank.',
            'Is the structure frame or massive (template list: frame / massive wood, concrete, steel, brick)?')
    x.issue('confirm', '', 'No operational modules (B6/B7) in the export file, but D3.2 reports LU operational carbon for 49 of 52 buildings '
            '(median 373 kgCO2e/m² UFA). Energy given as primary energy demand per heated area only.',
            'Luxembourg drops out of whole-life views; primary energy kept as x_ fields.',
            'Please provide the B6 (and B7) results used in D7.30/D3.2 and the annual energy consumption (kWh/a).')
    x.issue('confirm', '', 'D3.2 embodied carbon for LU (median 1,372 kgCO2e/m² UFA) is higher than the sum of the reported embodied modules '
            '(median 1,324), although each module matches D3.2.', 'Embodied = sum of A1-A3, A4, A5, B4-B5, C1-C4 as reported.',
            'Which additional values does the D3.2 / D7.30 embodied total include?')
    x.issue('info', '', 'Type of intervention not stated (construction years 2017–2028).', 'Treated as new construction.', 'Are all 52 projects new construction?')
    x.issue('info', '', 'Users of office buildings equal UFA ÷ 10.', 'Kept, flagged as likely estimated (x_users_basis).')
    return x
