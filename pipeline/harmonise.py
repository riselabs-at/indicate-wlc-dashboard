"""Combine the country extracts into harmonised tables and run cross-dataset checks."""
import datetime
import pandas as pd
from config import (MODULES, EMBODIED, WHOLE_LIFE, RSP_YEARS, SOURCES, MODULE_MAP, TYPOLOGY_MAP, COUNTRY_NAMES,
                    DATABASE_VERSION)
import extract_at, extract_hr, extract_lu, extract_it

EXTRACTORS = [extract_at, extract_hr, extract_it, extract_lu]

BUILDING_COLUMNS = ['building_id', 'country', 'source_deliverable', 'source_id', 'national_typology', 'harmonised_type',
                    'intervention', 'intervention_class', 'renovation_level', 'gfa_m2', 'ufa_m2', 'hfa_m2',
                    'ufa_equivalent_m2', 'ufa_equivalent_basis', 'rsp_years', 'results_basis', 'has_element_data',
                    'storeys', 'year', 'structure', 'energy_source', 'typology_use', 'n_products',
                    'gwp_fossil_kgco2e', 'gwp_biogenic_kgco2e', 'gwp_luluc_kgco2e']

GENERAL_STEPS = [
    ('1', 'Read', 'Each national workbook is read by its own extractor (pipeline/extract_<country>.py). Only result sheets are used; cover sheets are not.'),
    ('2', 'Clean', 'Blank rows, repeated header rows and non-building rows are dropped. Numbers are parsed defensively (decimal commas, text cells → blank).'),
    ('3', 'Identify', 'Each building gets a database ID "<country>-<source ID>". National typology, intervention type and floor areas (GFA, UFA, HFA) are kept as reported.'),
    ('4', 'Map modules', 'Source columns are mapped to 9 harmonised modules (A1-A3, A4, A5, B2-B3, B4-B5, B6, B7, C1-C4, D); see map_modules. Values are stored as absolute kgCO2e per building over a 50-year reference study period.'),
    ('5', 'Convert units', 'Absolute results (AT, HR, IT) are used as reported. Luxembourg reports per m² without stating the area; absolute = per-m² value × GFA (assumption, see issues).'),
    ('6', 'Map typologies', 'National typologies are mapped to harmonised building types; see map_typologies. Codes without a definition stay as their own type.'),
    ('7', 'Normalise', 'Per-m² results = absolute ÷ area (GFA, or the UFA-equivalent: UFA for AT/HR/LU, HFA for IT). Per year = ÷ 50. These are formulas in the results sheet.'),
    ('8', 'Elements', 'Element-level results are kept in each dataset\'s own classification: LU "LEVELS Tier 3" codes (summed from product rows), IT tiers (from the SumUP Tiers sheet, with a quality flag per value read from the cell colours). They are not merged across countries.'),
    ('9', 'Check', 'Reconciliation checks per dataset (module sums vs source totals, absolute vs per-m², element sums vs totals); see qa_checks.'),
    ('10', 'Flag', 'Assumptions and data-quality observations are logged in issues, marked "confirm" where a partner needs to confirm the choice.'),
    ('11', 'Export', 'This workbook, CSV copies of the tables and the dashboard data file (assets/js/data.js) are written by pipeline/run_all.py.'),
]


def build(source_folder):
    extracts = [m.extract(source_folder) for m in EXTRACTORS]
    log = [(c, *rest) for x in extracts for (c, *rest) in x.log]
    issues = [i for x in extracts for i in x.issues]
    checks = [c for x in extracts for c in x.checks]

    b = pd.DataFrame([r for x in extracts for r in x.buildings])
    b['source_deliverable'] = b.country.map(lambda c: SOURCES[c]['deliverable'])
    b['intervention_class'] = b.intervention.map(lambda s: 'Renovation' if 'renovation' in s.lower() else 'New construction')
    b['ufa_equivalent_m2'] = b.ufa_m2.where(b.country != 'IT', b.hfa_m2)
    b['ufa_equivalent_basis'] = b.country.map(lambda c: 'HFA (heated floor area)' if c == 'IT' else 'UFA')
    el_ids = {r['building_id'] for x in extracts for r in x.elements}
    b['has_element_data'] = b.building_id.isin(el_ids).map({True: 'yes', False: 'no'})
    for c in BUILDING_COLUMNS:
        if c not in b:
            b[c] = None
    b = b[BUILDING_COLUMNS]
    assert b.building_id.is_unique, 'duplicate building IDs'

    lcm = pd.DataFrame([r for x in extracts for r in x.lcm])
    lcm['country'] = lcm.building_id.str[:2]
    lcm = lcm[['building_id', 'country', 'module', 'source_value', 'source_unit', 'source_columns', 'value_kgco2e']]

    wide = lcm.pivot(index='building_id', columns='module', values='value_kgco2e').reindex(columns=MODULES)
    wide = b.set_index('building_id')[['country', 'harmonised_type', 'intervention_class', 'gfa_m2',
                                       'ufa_equivalent_m2', 'rsp_years']].join(wide).reset_index()

    elements = pd.DataFrame([r for x in extracts for r in x.elements])
    lu_el = elements[elements.classification.str.startswith('LU')]
    lu_el = lu_el.pivot_table(index=['building_id', 'code', 'group'], columns='module', values='value_kgco2e_m2',
                              aggfunc='sum').reindex(columns=[m for m in MODULES if m in lu_el.module.unique()]).reset_index()
    it_el = elements[elements.classification.str.startswith('IT')][
        ['building_id', 'code', 'group', 'value_kgco2e_m2', 'quality', 'note', 'source_cell']]
    materials = pd.DataFrame([r for x in extracts for r in x.materials])
    lu_mat = materials.pivot_table(index=['building_id', 'material_group'], columns='module', values='value_kgco2e_m2',
                                   aggfunc='sum').reset_index()

    desc = {}
    for x in extracts:
        desc.update(x.element_desc)
    codes = []
    for code in sorted(lu_el.code.unique(), key=lambda s: [int(p) if p.isdigit() else p for p in s.split('.')]):
        codes.append(('LU LEVELS Tier 3', code, code.rsplit('.', 1)[0], desc.get(('LU LEVELS Tier 3', code), '')))
    for code in sorted(it_el.code.unique(), key=lambda s: [int(p) for p in s.split('.')]):
        g = it_el[it_el.code == code]
        codes.append(('IT tiers (D7.19/D7.20)', code, g.group.iloc[0],
                      'Name not given in the file. ' + ('Tier 2.3 includes 2.4 (source note). ' if code == '2.3' else '')
                      + ('Insufficient data for D7.22 benchmarks (source note). ' if code in ('2.2', '2.5', '2.6', '2.7', '2.8', '2.9') else '')
                      + ('Renovation cases only. ' if code == '4.1' else '')))
    codes = pd.DataFrame(codes, columns=['classification', 'code', 'group', 'description'])

    # cross-check: IT tier sum vs LCM embodied per m² GFA (tiers include proxies, so differences are expected)
    it_new = b[(b.country == 'IT')].set_index('building_id')
    tier_tot = it_el.groupby('building_id').value_kgco2e_m2.sum()
    w = wide.set_index('building_id')
    emb_all = w.loc[it_new.index, ['A1-A3', 'A4', 'A5', 'B2-B3', 'B4-B5', 'C1-C4']].sum(axis=1) / w.loc[it_new.index, 'gfa_m2']
    ratio = (tier_tot / emb_all).dropna()
    checks.append(('IT', 'Tier sum ÷ LCM embodied (A–C excl. B6) per m² GFA: median ratio (info only; proxies included)',
                   round(float(ratio.median()), 3), 1e9, True))
    log.append(('IT', 'compare', 'Tier totals vs LCM embodied per m² GFA, ratio range',
                f'{ratio.min():.2f} – {ratio.max():.2f} (median {ratio.median():.2f})'))

    maps = dict(
        modules=pd.DataFrame(MODULE_MAP, columns=['country', 'source_sheet', 'source_columns', 'harmonised_module', 'note']),
        typologies=pd.DataFrame(TYPOLOGY_MAP, columns=['country', 'national_typology', 'harmonised_type', 'note']),
    )
    sources = pd.DataFrame([(c, COUNTRY_NAMES[c], s['deliverable'], s['title'], s['lead'],
                             next(x.source_file for x in extracts if x.country == c).rsplit('/', 1)[-1])
                            for c, s in SOURCES.items()],
                           columns=['country', 'country_name', 'deliverable', 'title', 'lead_organisation', 'file_name'])
    meta = dict(version=DATABASE_VERSION, built=datetime.date.today().isoformat(), rsp=RSP_YEARS,
                embodied=EMBODIED, whole_life=WHOLE_LIFE, modules=MODULES)
    steps = pd.DataFrame(GENERAL_STEPS, columns=['step', 'name', 'description'])
    log = pd.DataFrame(log, columns=['country', 'action', 'detail', 'result'])
    issues = pd.DataFrame(issues, columns=['country', 'type', 'building_ids', 'observation', 'handling'])
    issues.insert(0, 'issue_id', [f'I{i + 1:02d}' for i in range(len(issues))])
    issues['status'] = issues.type.map({'confirm': 'open – partner to confirm', 'warning': 'handled', 'info': 'noted'})
    checks = pd.DataFrame(checks, columns=['country', 'check', 'value', 'tolerance', 'passed'])
    return dict(meta=meta, sources=sources, steps=steps, log=log, buildings=b, lcm=lcm, wide=wide, lu_elements=lu_el,
                it_elements=it_el, lu_materials=lu_mat, codes=codes, maps=maps, issues=issues, checks=checks,
                extracts=extracts)
