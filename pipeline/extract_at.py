"""Austria: D7.3, sheet 'Data overview'. Absolute kgCO2e per building and module, UFA only."""
import pandas as pd
from common import Extract, num, find_source
from config import SOURCES, TYPE_OF, RSP_YEARS

PARTS = ['A1A3', 'A4', 'A5', 'B4', 'B5', 'C1', 'C2', 'C3C4']


def extract(folder):
    src = find_source(folder, SOURCES['AT']['pattern'])
    x = Extract('AT', src)
    df = pd.read_excel(src, sheet_name='Data overview', header=0, skiprows=[1])  # row 2 = long descriptions
    x.step('read', "Sheet 'Data overview': header row 1, description row 2 skipped", len(df))
    df = df[df['name'].notna()]
    x.step('filter', 'Rows with a building ID', len(df))

    for _, r in df.iterrows():
        bid = 'AT-' + str(r['name']).zfill(3)
        rsp = num(r['Unnamed: 4'])
        x.buildings.append(dict(
            building_id=bid, country='AT', source_id=str(r['name']).zfill(3),
            national_typology=r['typology'], harmonised_type=TYPE_OF[('AT', r['typology'])],
            intervention='New construction (assumed)', renovation_level='',
            gfa_m2=None, ufa_m2=num(r['UFA']), hfa_m2=None, storeys=num(r['Unnamed: 3']),
            rsp_years=rsp, results_basis='absolute kgCO2e',
            gwp_fossil_kgco2e=sum(num(r[f'GWP fossil - {p}']) for p in PARTS),
            gwp_biogenic_kgco2e=sum(num(r[f'GWP biogenic - {p}']) for p in PARTS),
            gwp_luluc_kgco2e=sum(num(r[f'GWP LULUC - {p}']) for p in PARTS),
        ))
        if rsp != RSP_YEARS:
            x.issue('warning', bid, f'Reference study period {rsp} ≠ {RSP_YEARS}', 'Per-year values use the reported RSP')
        modules = {
            'A1-A3': (['GWP total - A1A3'],), 'A4': (['GWP total - A4'],), 'A5': (['GWP total - A5'],),
            'B4-B5': (['GWP total - B4'],), 'B6': (['GWP total - B5'],),
            'C1-C4': (['GWP total - C1', 'GWP total - C2', 'GWP total - C3C4'],),
        }
        for m, (cols,) in modules.items():
            v = sum(num(r[c]) for c in cols)
            x.lcm.append(dict(building_id=bid, module=m, source_value=v, source_unit='kgCO2e',
                              value_kgco2e=v, source_columns=' + '.join(cols)))
    x.step('map', 'Modules mapped: A1-A3, A4, A5, B4→B4-B5, "B5"→B6, C1+C2+C3C4→C1-C4', len(x.lcm))

    # gas split must add up to GWP total
    tot = sum(l['value_kgco2e'] for l in x.lcm)
    split = sum(b['gwp_fossil_kgco2e'] + b['gwp_biogenic_kgco2e'] + b['gwp_luluc_kgco2e'] for b in x.buildings)
    x.check('GWP total = fossil + biogenic + LULUC (relative difference, all buildings)', (tot - split) / tot, 0.001)

    x.issue('confirm', '', 'Column header "GWP total - B5" but description row says B6', 'Treated as B6 (operational energy)')
    x.issue('info', '', 'B4-B5 contains B4 only', 'Mapped to B4-B5')
    x.issue('info', '', 'No GFA reported', 'Austria drops out of per-GFA views')
    x.issue('info', '', 'Type of intervention not stated', 'Treated as new construction')
    return x
