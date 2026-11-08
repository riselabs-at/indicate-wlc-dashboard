"""Combine the country extracts into the template-based database tables and run cross-checks."""
import datetime
import pandas as pd
from config import (MODULES, INDICATORS, DISPLAY, DISPLAY_ORDER, EMBODIED, WHOLE_LIFE, RSP_YEARS, SOURCES, TYPOLOGY_MAP,
                    TYPOLOGIES, ELEMENT_MAP, EL_MAP, EU_ELEMENTS, EU_NAME, EU_TIER, CODEBOOK, PROJECT_COLUMNS,
                    DATABASE_VERSION, TEMPLATE_VERSION)
import extract_at, extract_hr, extract_it, extract_lu

EXTRACTORS = [extract_at, extract_hr, extract_it, extract_lu]
DISPLAY_OF = {m: d for d, ms in DISPLAY.items() for m in ms}
EMB_NATIVE = [m for d in EMBODIED for m in DISPLAY[d]]

STEPS = [
    ('1', 'Read', 'Each national workbook is read by its own extractor (pipeline/extract_<country>.py). Only result and metadata sheets are used.'),
    ('2', 'Clean', 'Blank rows, repeated header rows and non-building rows are dropped. Numbers are parsed defensively (decimal commas, text → blank).'),
    ('3', 'Describe buildings in template vocabulary', 'One row per building in "projects", using the attribute codes of the INDICATE reporting template v1.1 '
     '(project_code, bldg_area_gfa, lca_rsp …). Information the template has no field for is kept in extension fields (x_…). '
     'Dataset-level information (organisation, country) is repeated per building.'),
    ('4', 'Map typologies', 'National typologies → the 9 template codes (SFH, MFH, ABL, OFF, TRA, EDU, HEA, HOR, OTH); see map_typologies. '
     'Apartment buildings → ABL, multi-family houses → MFH (agreed 2026-10-20).'),
    ('5', 'Map results to template modules', 'Every reported value is stored once in "wlc_results" as building × indicator (GWPtotal, GWPfossil, GWPbio, GWPluluc) '
     '× template module (A1-A3 … D2, incl. B5). A value covering several modules is stored in the first module of its range, with the range in '
     '"covers" (template rule); see map_modules.'),
    ('6', 'Convert units', 'Absolute kgCO2e per building over a 50-year reference study period. AT, HR and IT report absolute values. LU reports per m² '
     'without stating the area; the values are per m² UFA (D3.2 Table E.1 reproduces them exactly), so absolute = per-m² value × UFA.'),
    ('7', 'Areas', 'GFA as reported. NFA = UFA (AT, HR, LU) or HFA (IT; no UFA reported) – agreed 2026-10-20; the basis is in x_nfa_basis.'),
    ('8', 'Elements', 'Element-level results (LU: LEVELS Tier 3 codes; IT: tiers incl. proxy-quality flags) are mapped to the EU Delegated Regulation '
     'Table 4 tiers used in template v1.1; see map_elements. Mappings are inferred from content and marked "to confirm".'),
    ('9', 'Display modules', 'For charts and the results sheet, template modules are grouped into A1-A3, A4, A5, B1-B3, B4-B5, B6, B7, C1-C4 and D. '
     'Embodied = A1-A3 + A4 + A5 + B4-B5 + C1-C4 (reported by all datasets); whole life adds B1-B3, B6, B7.'),
    ('10', 'Check', 'Reconciliation per dataset (module sums vs source totals, absolute vs per m², sub-indicators vs total, element sums vs totals); see qa_checks.'),
    ('11', 'Flag and ask', 'Assumptions and data-quality observations are logged in "issues", with the question to the national team where input is needed.'),
    ('12', 'Export', 'This workbook, CSV copies, the dashboard data file and the pre-filled template overviews per country (pipeline/run_all.py).'),
]


def build(source_folder):
    ex = [m.extract(source_folder) for m in EXTRACTORS]
    log = [r for x in ex for r in x.log]
    issues = [i for x in ex for i in x.issues]
    checks = [c for x in ex for c in x.checks]

    # ---------------- projects
    p = pd.DataFrame([r for x in ex for r in x.projects])
    el_ids = {r['building_id'] for x in ex for r in x.elements}
    p['x_has_elements'] = p.building_id.isin(el_ids).map({True: 'yes', False: 'no'})
    for c in PROJECT_COLUMNS:
        if c not in p:
            p[c] = None
    p = p[PROJECT_COLUMNS]
    assert p.building_id.is_unique
    p['country'] = p.building_id.str[:2]
    code_of = {v: k for k, v in TYPOLOGIES.items()}
    p['typology_code'] = p.bldg_use_subtype.map(code_of)

    # ---------------- results (long, template modules)
    r = pd.DataFrame([x for e in ex for x in e.results])
    r.insert(1, 'country', r.building_id.str[:2])
    r['display_module'] = r.module.map(DISPLAY_OF)
    assert set(r.module) <= set(MODULES) and set(r.indicator) <= set(INDICATORS)
    assert not r.duplicated(['building_id', 'indicator', 'module']).any()

    tot = r[r.indicator == 'GWPtotal']
    wide = tot.pivot_table(index='building_id', columns='display_module', values='value_kgco2e', aggfunc='sum', dropna=False)
    reported = tot.dropna(subset=['value_kgco2e']).groupby(['building_id', 'display_module']).size().unstack()
    wide = wide.where(reported.reindex_like(wide).notna()).reindex(columns=DISPLAY_ORDER)
    wide = p.set_index('building_id')[['country', 'typology_code', 'project_intervention_type', 'project_renovation_depth',
                                       'bldg_area_gfa', 'bldg_area_nfa', 'lca_rsp']].join(wide).reset_index()
    wide['lca_rsp'] = wide.lca_rsp.fillna(RSP_YEARS)

    # ---------------- elements → EU Table 4
    e = pd.DataFrame([x for k in ex for x in k.elements])
    e.insert(1, 'country', e.building_id.str[:2])
    m = e.apply(lambda row: EL_MAP.get((row.country, row.national_code), (None, 'unmapped', '')), axis=1, result_type='expand')
    e['eu_code'], e['mapping_status'], e['mapping_basis'] = m[0], m[1], m[2]
    unm = e[e.eu_code.isna()][['country', 'national_code']].drop_duplicates()
    if len(unm):
        issues.append(('ALL', 'warning', '', f'Element codes without mapping: {unm.values.tolist()}', 'Kept as "4.1 Other or not allocated".', ''))
        e['eu_code'] = e.eu_code.fillna('4.1')
    e['eu_name'] = e.eu_code.map(EU_NAME)
    e['eu_tier'] = e.eu_code.map(EU_TIER)
    e['eu_tier2'] = e.eu_code.str.split('.').str[:2].str.join('.')
    e['eu_tier1'] = e.eu_code.str.split('.').str[0]
    pa = p.set_index('building_id'); gfa = pa.bldg_area_gfa
    ref = e.building_id.map(pa.bldg_area_gfa).where(e.area_basis != 'UFA', e.building_id.map(pa.bldg_area_nfa))
    e['value_kgco2e'] = e.value_kgco2e_m2 * ref                     # LU per m² UFA, IT per m² GFA
    e = e[['building_id', 'country', 'classification', 'national_code', 'eu_code', 'eu_name', 'eu_tier', 'eu_tier2', 'eu_tier1',
           'mapping_status', 'mapping_basis', 'module', 'value_kgco2e_m2', 'area_basis', 'value_kgco2e', 'quality', 'note', 'source_cell']]
    emb = e[(e.module.isin(EMB_NATIVE)) | (e.module == 'A–C (as reported)')]
    tier2 = emb.pivot_table(index='building_id', columns='eu_tier2', values='value_kgco2e_m2', aggfunc='sum').reset_index()
    tier2.insert(1, 'country', tier2.building_id.str[:2])

    mat = pd.DataFrame([x for k in ex for x in k.materials])
    mat_w = mat.pivot_table(index=['building_id', 'material_group'], columns='module', values='value_kgco2e_m2', aggfunc='sum')
    mat_w = mat_w.reindex(columns=[c for c in MODULES if c in mat_w.columns]).reset_index()

    # cross-check: IT tier sums vs LCM embodied (A–C excl. B6) per m² GFA (tiers include proxies → info only)
    it_ids = p[p.country == 'IT'].building_id
    t_sum = e[e.country == 'IT'].groupby('building_id').value_kgco2e_m2.sum()
    lcm = tot[tot.country == 'IT'][~tot.module.isin(['B6', 'B7', 'D1', 'D2'])].groupby('building_id').value_kgco2e.sum() / gfa.reindex(it_ids)
    ratio = (t_sum / lcm).dropna()
    checks.append(('IT', 'Tier sum ÷ LCM A–C per m² GFA: median ratio (information only; tiers include proxies)', round(float(ratio.median()), 3), 1e9, True))
    log.append(('IT', 'compare', 'Tier sums vs life-cycle A–C totals per m² GFA, ratio range', f'{ratio.min():.2f} – {ratio.max():.2f} (median {ratio.median():.2f})'))
    lu_el = e[(e.country == 'LU')].groupby('building_id').value_kgco2e.sum()
    lu_tot = tot[tot.country == 'LU'].groupby('building_id').value_kgco2e.sum()
    checks.append(('LU', 'Elements mapped to EU tiers sum to building totals incl. D (max abs difference, kgCO2e)',
                   float((lu_el - lu_tot).abs().max()), 1, bool((lu_el - lu_tot).abs().max() <= 1)))

    # ---------------- codebook with coverage per country
    cov = []
    for code, label, unit, prio, kind in CODEBOOK:
        row = dict(attribute_code=code, label=label, unit=unit, priority=prio, kind=kind)
        for c in ['AT', 'HR', 'IT', 'LU']:
            s = p.loc[p.country == c, code]
            row[f'coverage_{c}'] = round(float(s.notna().mean()), 3) if len(s) else None
        cov.append(row)
    for mod in MODULES:
        row = dict(attribute_code=f'GWPtotal {mod}', label=f'GWPtotal, module {mod}', unit='kgCO2e', priority='HP', kind='template')
        for c in ['AT', 'HR', 'IT', 'LU']:
            nb = (p.country == c).sum()
            s = tot[(tot.country == c) & (tot.module == mod)].dropna(subset=['value_kgco2e'])
            comb = (s.covers != mod).any()
            row[f'coverage_{c}'] = round(len(s) / nb, 3) if nb else None
            if comb and len(s):
                row['note_' + c] = 'covers ' + ', '.join(sorted(set(s.covers)))
        cov.append(row)
    codebook = pd.DataFrame(cov)

    # ---------------- mapping tables
    map_mod = (r.groupby(['country', 'indicator', 'source_columns', 'module', 'covers', 'flag', 'source_unit'], dropna=False)
                .size().rename('n_buildings').reset_index().sort_values(['country', 'indicator', 'module']))
    map_typ = pd.DataFrame(TYPOLOGY_MAP, columns=['country', 'national_typology', 'template_code', 'status', 'note'])
    map_typ['template_label'] = map_typ.template_code.map(TYPOLOGIES)
    n_typ = p.groupby(['country', 'x_typology_national']).size().rename('n_buildings').reset_index()
    map_typ = map_typ.merge(n_typ, left_on=['country', 'national_typology'], right_on=['country', 'x_typology_national'], how='left').drop(columns='x_typology_national')
    desc = {k: v for x in ex for k, v in x.element_desc.items()}
    map_el = pd.DataFrame(ELEMENT_MAP, columns=['country', 'national_code', 'eu_code', 'status', 'basis'])
    map_el['eu_name'] = map_el.eu_code.map(EU_NAME)
    map_el['national_content (LU: top-3 product categories)'] = map_el.apply(lambda r_: desc.get((r_.country, r_.national_code), ''), axis=1)
    eu = pd.DataFrame(EU_ELEMENTS, columns=['eu_code', 'tier', 'name'])

    sources = pd.DataFrame([(c, s['country'], s['iso3'], s['deliverable'], s['title'], s['org'],
                             next(x.source_file for x in ex if x.country == c).rsplit('/', 1)[-1]) for c, s in SOURCES.items()],
                           columns=['country', 'site_country', 'site_country_iso3', 'deliverable', 'title', 'contact_org', 'file_name'])
    log = pd.DataFrame(log, columns=['country', 'action', 'detail', 'result'])
    issues = pd.DataFrame(issues, columns=['country', 'type', 'building_ids', 'observation', 'handling', 'question_to_national_team'])
    issues.insert(0, 'issue_id', [f'I{i + 1:02d}' for i in range(len(issues))])
    issues['status'] = issues.type.map({'confirm': 'open – partner to confirm', 'warning': 'handled', 'info': 'noted'})
    checks = pd.DataFrame(checks, columns=['country', 'check', 'value', 'tolerance', 'passed'])
    meta = dict(version=DATABASE_VERSION, template=TEMPLATE_VERSION, built=datetime.date.today().isoformat(), rsp=RSP_YEARS,
                embodied=EMBODIED, whole_life=WHOLE_LIFE)
    return dict(meta=meta, sources=sources, steps=pd.DataFrame(STEPS, columns=['step', 'name', 'description']), log=log,
                projects=p, results=r, wide=wide, elements=e, tier2=tier2, materials=mat_w, codebook=codebook,
                map_modules=map_mod, map_typologies=map_typ, map_elements=map_el, eu_elements=eu, issues=issues, checks=checks)
