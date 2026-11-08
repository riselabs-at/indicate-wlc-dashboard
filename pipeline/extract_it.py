"""Italy – D7.21 summary workbook.

'SumUP LCM'   – 50 cases: results per life-cycle module (absolute kgCO2e block used), GFA and HFA.
'SumUP Tiers' – the same cases per building-element tier (kgCO2e/m² GFA). Cell colours mark original values and
                several kinds of proxy values (sheet legend); each tier value keeps that quality flag.
"""
import os
import openpyxl
import pandas as pd
from common import Extract, num, find_source
from config import SOURCES, TYPE_OF, TYPOLOGIES, DEPTH, RSP_YEARS, IT_FILL_QUALITY, IT_FONT_BLUE, IT_FONT_GREY

S = SOURCES['IT']
C = dict(code=0, intervention=1, level=2, typology=4, typ_code=5, gwptot=7, gwp_gfa=8, gwp_gfa_y=9, gfa=12, hfa=13)
# absolute block: template module -> (source columns, flag)
ABS = {'A1-A3': ([45], ''), 'A4': ([46], ''),
       'A5': ([47, 48, 49], 'A5 = A5 (GWP-fossil) + A5w_DAB + A5w_BOQ (construction waste)'),
       'B2': ([50], 'reported as GWP-fossil'), 'B3': ([51], 'reported as GWP-fossil'), 'B4': ([52], 'B5 not reported'),
       'B6': ([53], ''), 'C1': ([54], 'reported as GWP-fossil'), 'C2': ([55], ''), 'C3': ([56], ''), 'C4': ([57], '')}
PER_GFA = {'A1-A3': [20], 'A4': [21], 'A5': [22], 'B2': [23], 'B3': [24], 'B4': [25], 'B6': [26], 'C1': [27], 'C2': [28], 'C3': [29], 'C4': [30]}


def colour_key(color):
    if color is None:
        return None
    if color.type == 'rgb':
        return color.rgb
    if color.type == 'theme':
        return ('theme', color.theme, round(color.tint or 0, 1))
    return None


def bid_of(code):
    return code.replace('IT_', 'IT-', 1)


def extract(folder):
    src = find_source(folder, S['pattern'])
    x = Extract('IT', src)
    raw = pd.read_excel(src, sheet_name='SumUP LCM', header=None)
    hdr = raw.iloc[3]
    df = raw.iloc[5:]
    df = df[df[C['code']].astype(str).str.startswith('IT_')]
    x.step('read', "Sheet 'SumUP LCM': header row 4, units row 5, one row per case from row 6", len(df))

    worst_abs = worst_tot = worst_y = 0
    for _, r in df.iterrows():
        code = str(r[C['code']]).strip()
        bid = bid_of(code)
        gfa, hfa = num(r[C['gfa']]), num(r[C['hfa']])
        interv = str(r[C['intervention']]).strip()
        level = str(r[C['level']]).strip()
        tcode = TYPE_OF[('IT', r[C['typology']])]
        x.projects.append(dict(
            building_id=bid, contact_org=S['org'], project_code=code,
            project_intervention_type=interv, project_renovation_depth=DEPTH.get(level, level),
            x_intervention_basis='reported (Intervention, Level: Minor R- / Major R+)',
            site_country=S['country'], site_country_iso3=S['iso3'],
            bldg_use_subtype=TYPOLOGIES[tcode], x_typology_national=r[C['typology']],
            bldg_area_gfa=gfa, bldg_area_nfa=hfa, x_nfa_basis='HFA (heated floor area; no UFA reported)',
            bldg_area_references='GFA and HFA (heated floor area) as reported; NFA = HFA',
            lca_rsp=RSP_YEARS, x_rsp_basis='derived: GWPtot/GFA ÷ GWPtot/GFA/year = 50',
            x_results_basis='absolute kgCO2e (also given per m² GFA and HFA)',
            x_source_deliverable=S['deliverable'], x_source_file=os.path.basename(src), x_source_id=code,
        ))
        s = 0
        for mod, (cols, flag) in ABS.items():
            v = sum(num(r[c]) or 0 for c in cols)
            per = sum(num(r[c]) or 0 for c in PER_GFA[mod])
            worst_abs = max(worst_abs, abs(per * gfa - v) / max(abs(v), 1))
            s += v
            if mod == 'B6' and v == 0:
                x.issue('warning', bid, 'B6 reported as 0 (operational carbon share 0 %).', 'Treated as not reported.',
                        'Is operational carbon (B6) available for this building?')
                v, flag = None, 'reported as 0; treated as not reported'
            x.result(bid, 'GWPtotal', mod, v, v, 'kgCO2e', ' + '.join(str(hdr[c]).strip() for c in cols), '', flag)
        worst_tot = max(worst_tot, abs(s - num(r[C['gwptot']])) / num(r[C['gwptot']]))
        worst_y = max(worst_y, abs(num(r[C['gwp_gfa_y']]) * RSP_YEARS - num(r[C['gwp_gfa']])))
        if hfa is None:
            x.issue('warning', bid, 'No HFA reported.', 'Excluded from per-NFA views.', 'Can the heated (or useful) floor area be provided?')
    x.step('map', 'Absolute block → template modules A1-A3, A4, A5 (incl. waste), B2, B3, B4, B6, C1–C4', len(x.results))
    x.check('Absolute = per-m² GFA × GFA, per module (max relative difference)', worst_abs, 1e-6)
    x.check('Sum of modules = GWPtot (max relative difference)', worst_tot, 1e-6)
    x.check('GWPtot/GFA/year × 50 = GWPtot/GFA (max abs difference) → RSP 50 years', worst_y, 1e-6)

    b6 = df[[C['code'], 53]].copy()
    b6[53] = pd.to_numeric(b6[53], errors='coerce')
    for v, g in b6[b6[53] > 0].groupby(53):
        if len(g) > 1:
            x.issue('confirm', ', '.join(map(bid_of, g[C['code']])), f'Identical B6 value ({v:,.0f} kgCO2e) for different buildings.',
                    'Kept as reported.', 'Is the identical B6 value intended (same energy model) or a copy error?')
    same = df.groupby([C['typology'], C['gfa'], C['hfa']])[C['code']].apply(list)
    for ids in same[same.str.len() > 1]:
        x.issue('info', ', '.join(map(bid_of, ids)), 'Cases share identical GFA and HFA (variants of one building).', 'Kept as separate records.')

    _extract_tiers(x, src)

    x.issue('confirm', '', 'No UFA reported; heated floor area (HFA) given instead.', 'NFA = HFA (agreed 2026-10-20).',
            'Is a useful floor area (UFA) available?')
    x.issue('confirm', '', 'Typology codes HOS and XTR are not defined in the file. D3.2 describes HOS as hospitality '
            '(mainly long-term stay and student rooms, 3 new buildings, grouped with MFH) and XTR as other non-residential (2 new buildings).',
            'HOS mapped to HOR (agreed 2026-11-08), XTR to OTH.',
            'Please confirm: HOS = hospitality (template HOR), XTR = other non-residential (OTH).')
    x.issue('confirm', '', 'D3.2 IT benchmarks (29 new buildings, per m² HFA, proxies excluded) differ from the values in this file '
            '(e.g. whole-life median 2,211 vs 2,315 kgCO2e/m² HFA; this file has no HFA for IT-XTR_07). Operational values match.',
            'Values used as in SumUP LCM.',
            'Which data version is final: does SumUP LCM include the proxy values excluded in D3.2? Please provide the HFA of XTR_07.')
    x.issue('confirm', '', 'Tier names are not given in the file.', 'Tiers 2.2–2.9 read as the EU Table 4 building-services list (agreed 2026-10-20).',
            'Please confirm the tier names (in particular 2.2–2.9 and 4.1).')
    lv = df[C['level']].astype(str).str.strip().value_counts()
    x.issue('info', '', f"{int(lv.get('Minor', 0) + lv.get('Major', 0))} of {len(df)} cases are renovations "
            f"({int(lv.get('Minor', 0))} minor R-, {int(lv.get('Major', 0))} major R+). D3.2 uses the new buildings only.",
            'project_renovation_depth set; renovations benchmarked separately.')
    x.issue('info', '', 'A5 construction part, B2, B3 and C1 are GWP-fossil only; D and B7 not reported.', 'Used as reported, flagged per value.')
    x.issue('confirm', '', 'Element tiers include proxy values (missing values, outlier replacement, cross-typology proxies, +16.67 paint).',
            'Each tier value carries a quality flag from the cell colour; tier sums differ from the life-cycle totals.')
    return x


def _extract_tiers(x, src):
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb['SumUP Tiers']
    val = lambda r, c: ws.cell(r, c).value
    n_new = n_ren = 0
    reno_start = next(r for r in range(1, ws.max_row + 1) if str(val(r, 4) or '').strip() == 'Renovation')

    def add(case, code, c):
        v = num(c.value)
        if v is None:
            return
        q = IT_FILL_QUALITY.get(colour_key(c.fill.fgColor) if c.fill.fill_type else None, 'original')
        notes = []
        fk = colour_key(c.font.color) if c.font and c.font.color is not None else None
        if fk == IT_FONT_BLUE:
            notes.append('includes +16.67 kgCO2e/m² proxy for internal paint')
        if fk == IT_FONT_GREY:
            notes.append('insufficient data for D7.22 benchmarks (grey in source)')
        x.elements.append(dict(building_id=bid_of(case), classification='IT tiers (D7.19/D7.20)', national_code=code,
                               module='A–C (as reported)', value_kgco2e_m2=v, area_basis='GFA', quality=q,
                               note='; '.join(notes), source_cell=f"'SumUP Tiers'!{c.coordinate}"))

    for r in range(1, reno_start):                                   # new construction, GFA half (C..M)
        if str(val(r, 2) or '').strip() != 'tier3':
            continue
        cases = {c: val(r - 1, c) for c in range(3, 14) if str(val(r - 1, c) or '').startswith('IT_')}
        rr = r + 1
        while not str(val(rr, 2) or '').strip().lower().startswith('tot'):
            code = str(val(rr, 2)).strip()
            if code and code != 'None':
                for c, case in cases.items():
                    add(case, code, ws.cell(rr, c))
            rr += 1
        n_new += len(cases)
    for r in range(reno_start, ws.max_row + 1):                      # renovation blocks
        cases = {c: val(r, c) for c in range(2, 15) if str(val(r, c) or '').startswith('IT_')}
        if not cases:
            continue
        rr = r + 3
        while str(val(rr, 1) or '').strip().upper() != 'TOT' and rr <= ws.max_row:
            for c, case in cases.items():
                add(case, str(val(rr, 1)).strip(), ws.cell(rr, c))
            rr += 1
        n_ren += len(cases)
    t = pd.DataFrame(x.elements)
    x.step('read', "Sheet 'SumUP Tiers': GFA-normalised tier values; quality flag from cell fill / font colour",
           f'{n_new} new + {n_ren} renovation cases, {len(t)} tier values')
    x.step('flag', 'Tier values by quality flag', '; '.join(f'{k}: {v}' for k, v in t['quality'].value_counts().items()))
