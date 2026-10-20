"""Italy: D7.21 summary workbook.

'SumUP LCM'   – 50 cases, results per life-cycle module, per m² GFA, per m² HFA and absolute (kgCO2e).
'SumUP Tiers' – the same cases per building-element tier (kgCO2e/m² GFA and HFA). Cell colours mark
                original values and several kinds of proxy values (see the sheet's legend), so tier
                values are not identical to the LCM results.
"""
import openpyxl
import pandas as pd
from common import Extract, num, find_source
from config import SOURCES, TYPE_OF, RSP_YEARS, IT_FILL_QUALITY, IT_FONT_BLUE, IT_FONT_GREY

# 0-based column positions in 'SumUP LCM'
C = dict(code=0, intervention=1, level=2, typology=4, typ_code=5, gwptot=7, gwp_gfa=8, gwp_gfa_y=9, gfa=12, hfa=13,
         ec_gfa=15, oc_gfa=17)
ABS = {'A1-A3': [45], 'A4': [46], 'A5': [47, 48, 49], 'B2-B3': [50, 51], 'B4-B5': [52], 'B6': [53], 'C1-C4': [54, 55, 56, 57]}
PER_GFA = {'A1-A3': [20], 'A4': [21], 'A5': [22], 'B2-B3': [23, 24], 'B4-B5': [25], 'B6': [26], 'C1-C4': [27, 28, 29, 30]}


def colour_key(color):
    if color is None:
        return None
    if color.type == 'rgb':
        return color.rgb
    if color.type == 'theme':
        return ('theme', color.theme, round(color.tint or 0, 1))
    return None


def extract(folder):
    src = find_source(folder, SOURCES['IT']['pattern'])
    x = Extract('IT', src)
    raw = pd.read_excel(src, sheet_name='SumUP LCM', header=None)
    hdr = raw.iloc[3]
    df = raw.iloc[5:]
    df = df[df[C['code']].astype(str).str.startswith('IT_')]
    x.step('read', "Sheet 'SumUP LCM': header row 4, units row 5, one row per case from row 6", len(df))

    worst_abs, worst_tot, worst_y = 0, 0, 0
    for _, r in df.iterrows():
        code = str(r[C['code']]).strip()
        bid = code.replace('IT_', 'IT-', 1)
        gfa, hfa = num(r[C['gfa']]), num(r[C['hfa']])
        interv = str(r[C['intervention']]).strip()
        level = str(r[C['level']]).strip()
        x.buildings.append(dict(
            building_id=bid, country='IT', source_id=code, national_typology=r[C['typology']],
            harmonised_type=TYPE_OF[('IT', r[C['typology']])],
            intervention=interv, renovation_level='' if level == '-' else level,
            gfa_m2=gfa, ufa_m2=None, hfa_m2=hfa, rsp_years=RSP_YEARS, results_basis='absolute kgCO2e',
            typology_use=r[C['typ_code']],
        ))
        s = 0
        for m, cols in ABS.items():
            v = sum(num(r[c]) or 0 for c in cols)
            per = sum(num(r[c]) or 0 for c in PER_GFA[m])
            worst_abs = max(worst_abs, abs(per * gfa - v) / max(abs(v), 1))
            if m == 'B6' and v == 0:
                x.issue('warning', bid, 'B6 reported as 0 (operational carbon share 0 %)', 'Treated as not reported')
                v = None
            x.lcm.append(dict(building_id=bid, module=m, source_value=v, source_unit='kgCO2e',
                              value_kgco2e=v, source_columns=' + '.join(str(hdr[c]).strip() for c in cols)))
            s += v or 0
        worst_tot = max(worst_tot, abs(s - num(r[C['gwptot']])) / num(r[C['gwptot']]))
        worst_y = max(worst_y, abs(num(r[C['gwp_gfa_y']]) * RSP_YEARS - num(r[C['gwp_gfa']])))
        if hfa is None:
            x.issue('warning', bid, 'No HFA reported', 'Excluded from per-HFA/UFA views')
    x.step('map', 'Absolute block mapped: A5 = A5 + A5w_DAB + A5w_BOQ; B2+B3 → B2-B3; B4 → B4-B5; C1..C4 summed', len(x.lcm))
    x.check('Absolute vs per-m² GFA × GFA, per module (max relative difference)', worst_abs, 1e-6)
    x.check('Sum of modules vs GWPtot (max relative difference)', worst_tot, 1e-6)
    x.check('GWPtot/GFA/year × 50 vs GWPtot/GFA (max abs difference) → RSP = 50 years', worst_y, 1e-6)

    # duplicated values worth flagging
    b6 = df[[C['code'], 53]].copy()
    b6[53] = pd.to_numeric(b6[53], errors='coerce')
    for v, g in b6[b6[53] > 0].groupby(53):
        if len(g) > 1:
            x.issue('confirm', ', '.join(c.replace('IT_', 'IT-', 1) for c in g[C['code']]), f'Identical B6 value ({v:,.0f} kgCO2e) for different buildings',
                    'Kept as reported')
    same = df.groupby([C['typology'], C['gfa'], C['hfa']])[C['code']].apply(list)
    for ids in same[same.str.len() > 1]:
        x.issue('info', ', '.join(c.replace('IT_', 'IT-', 1) for c in ids), 'Cases share identical GFA and HFA (variants of one building)', 'Kept as separate records')

    _extract_tiers(x, src)

    x.issue('confirm', '', 'No UFA reported; heated floor area (HFA) given instead', 'HFA used as the UFA-equivalent in per-UFA views (switchable)')
    x.issue('confirm', '', 'Typology codes HOS and XTR are not defined in the file', 'Kept as own harmonised types: "Other residential (IT: HOS)", "Other non-residential (IT: XTR)"')
    x.issue('info', '', '21 of 50 cases are renovations (12 minor R-, 9 major R+)', 'Flagged by intervention; dashboard defaults to new construction')
    x.issue('info', '', 'B2 and B3 are GWP-fossil only; C1 is GWP-fossil only', 'Used as reported')
    x.issue('info', '', 'Module D and B7 not reported', '')
    x.issue('confirm', '', 'Element tiers include proxy values (missing values, outlier replacement, cross-typology proxies, +16.67 paint)', 'Each tier value carries a quality flag from the cell colour; tier totals differ from LCM totals')
    return x


def _extract_tiers(x, src):
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb['SumUP Tiers']
    cell = lambda r, c: ws.cell(r, c)          # 1-based
    val = lambda r, c: ws.cell(r, c).value
    n_new = n_ren = 0
    reno_start = next(r for r in range(1, ws.max_row + 1) if str(val(r, 4) or '').strip() == 'Renovation')

    def add(bid, code, group, c):
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
        x.elements.append(dict(building_id=bid, classification='IT tiers (D7.19/D7.20)', code=code, group=group,
                               module='embodied (as in source)', value_kgco2e_m2=v, area_basis='GFA', quality=q,
                               note='; '.join(notes), source_cell=c.coordinate))

    # new construction blocks (GFA half: columns C..M)
    for r in range(1, reno_start):
        if str(val(r, 2) or '').strip() != 'tier3':
            continue
        cases = {c: val(r - 1, c) for c in range(3, 14) if str(val(r - 1, c) or '').startswith('IT_')}
        rr = r + 1
        while not str(val(rr, 2) or '').strip().lower().startswith('tot'):
            code = str(val(rr, 2)).strip()
            if code and code != 'None':
                for c, case in cases.items():
                    add(case.replace('IT_', 'IT-', 1), code, str(val(rr, 1)).strip(), cell(rr, c))
            rr += 1
        n_new += len(cases)
    # renovation blocks: case IDs in a row, tiers below in column A until 'TOT'
    for r in range(reno_start, ws.max_row + 1):
        cases = {c: val(r, c) for c in range(2, 15) if str(val(r, c) or '').startswith('IT_')}
        if not cases:
            continue
        rr = r + 3
        while str(val(rr, 1) or '').strip().upper() != 'TOT' and rr <= ws.max_row:
            code = str(val(rr, 1)).strip()
            group = code if code.count('.') < 2 else code.rsplit('.', 1)[0]
            for c, case in cases.items():
                add(case.replace('IT_', 'IT-', 1), code, group, cell(rr, c))
            rr += 1
        n_ren += len(cases)
    x.step('read', "Sheet 'SumUP Tiers': GFA-normalised tier values; quality flag from cell fill and font colour",
           f'{n_new} new-construction + {n_ren} renovation cases, {len(x.elements)} tier values')

    t = pd.DataFrame(x.elements)
    x.step('flag', 'Tier values by quality flag', '; '.join(f'{k}: {v}' for k, v in t['quality'].value_counts().items()))
