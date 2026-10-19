"""Harmonise the INDICATE LIFE national WLC datasets (AT, HR, LU) into assets/js/data.js for the dashboard.

Usage (from the repository root):
    pip install pandas openpyxl
    python tools/build_data.py path/to/source-folder

The source folder must contain the three deliverable workbooks (not included in this repository):
    *IND-AT*.xlsx      D7.3  AT national WLC dataset
    *HR_WLC*.xlsx      D7.12 HR national WLC dataset
    *LU_Export*.xlsx   D7.30 LU national WLC dataset
"""
import json, math, glob, sys, os, datetime
import pandas as pd

D = (sys.argv[1] if len(sys.argv) > 1 else 'source').rstrip('/') + '/'
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
AT_F = glob.glob(D + '*IND-AT*')[0]
HR_F = glob.glob(D + '*HR_WLC*')[0]
LU_F = glob.glob(D + '*LU_Export*')[0]

# Harmonised life-cycle modules (order = stacking order)
MODS = ['A1-A3', 'A4', 'A5', 'B3', 'B4-B5', 'B6', 'B7', 'C1-C4', 'D']

HTYPE = {
    # AT
    'Single-family house': 'Single-family house', 'Apartment block': 'Multi-family / apartments',
    'Multi-family house': 'Multi-family / apartments', 'Education building': 'Education',
    'Retail': 'Retail', 'Office building': 'Office', 'Hotel and Restaurant': 'Hotel',
    # HR
    'SINGLE FAMILY HOUSE': 'Single-family house', 'MULTY FAMILY HOUSE': 'Multi-family / apartments',
    'OFFICE BUILDING': 'Office', 'EDUCATION': 'Education', 'HOSPITAL': 'Hospital', 'HOTEL': 'Hotel',
    'SPORT HALL BUILDING': 'Hall / sports', 'NEW BUILDING FOR RETAIL AND WHOLESALE': 'Retail',
    # LU
    'Office': 'Office', 'Hall': 'Hall / sports', 'House': 'Single-family house',
    'Apartments': 'Multi-family / apartments', 'School': 'Education',
}
HR_NAT = {'SINGLE FAMILY HOUSE': 'Single family house', 'MULTY FAMILY HOUSE': 'Multi family house',
          'OFFICE BUILDING': 'Office building', 'EDUCATION': 'Education', 'HOSPITAL': 'Hospital',
          'HOTEL': 'Hotel', 'SPORT HALL BUILDING': 'Sport hall',
          'NEW BUILDING FOR RETAIL AND WHOLESALE': 'Retail and wholesale'}

def f(x):
    try:
        v = float(str(x).replace(',', '.'))
        return None if math.isnan(v) else v
    except Exception:
        return None

buildings = []

# ---------------- AT ----------------
at = pd.read_excel(AT_F, sheet_name='Data overview', header=0, skiprows=[1])
at = at[at['name'].notna()]
for _, r in at.iterrows():
    m = {
        'A1-A3': r['GWP total - A1A3'], 'A4': r['GWP total - A4'], 'A5': r['GWP total - A5'],
        'B4-B5': r['GWP total - B4'],
        'B6': r['GWP total - B5'],  # header says B5, row-2 description says B6 -> operational energy
        'C1-C4': r['GWP total - C1'] + r['GWP total - C2'] + r['GWP total - C3C4'],
    }
    buildings.append(dict(
        id='AT-' + str(r['name']).zfill(3), country='AT', natType=r['typology'], type=HTYPE[r['typology']],
        ufa=f(r['UFA']), gfa=None, storeys=f(r['Unnamed: 3']), rsp=f(r['Unnamed: 4']),
        mods={k: float(v) for k, v in m.items()},
        gas={'fossil': float(sum(r[f'GWP fossil - {k}'] for k in ['A1A3','A4','A5','B4','B5','C1','C2','C3C4'])),
             'biogenic': float(sum(r[f'GWP biogenic - {k}'] for k in ['A1A3','A4','A5','B4','B5','C1','C2','C3C4'])),
             'luluc': float(sum(r[f'GWP LULUC - {k}'] for k in ['A1A3','A4','A5','B4','B5','C1','C2','C3C4']))},
    ))

# ---------------- HR ----------------
hr = pd.read_excel(HR_F, sheet_name='Data_overview', header=None)
hr = hr.iloc[5:]
hr = hr[hr[2].notna() & hr[1].notna()]
for _, r in hr.iterrows():
    m = {'A1-A3': r[12], 'A4': r[13], 'A5': r[14], 'B3': r[15], 'B4-B5': r[16], 'B6': r[17],
         'B7': r[18], 'C1-C4': r[19], 'D': r[20]}
    buildings.append(dict(
        id='HR-' + str(r[2]), country='HR', natType=HR_NAT[r[1]], type=HTYPE[r[1]], stage=str(r[0]).title(),
        ufa=f(r[10]), gfa=f(r[9]), rsp=50.0,
        mods={k: float(v) for k, v in m.items()}, srcTotal=float(r[3]),
    ))

# ---------------- LU ----------------
meta = pd.read_excel(LU_F, sheet_name='A. Metadata', header=None).iloc[3:, :22]
meta.columns = ['proj', 'typ', 'structure', 'ufa', 'gfa_b', 'gfa_a', 'fl_b', 'fl_a', 'compact', 'wwr', 'gfa',
                'capita', 'year', 'scope', 'energy', 'eprim', 'an', 'pv', 'elscope', 'heat', 'cool', 'facade']
meta = meta[meta.proj.notna()]
res = pd.read_excel(LU_F, sheet_name='Overview of results')
res = res[res.Project.notna() & (res.Project != 'Projet')]
COLS = {'A1-A3': ['A1-A3, fossil', 'A1-A3, biog'], 'A4': ['A4, fossil'], 'A5': ['A5, fossil'],
        'B4-B5': ['B4-B5, fossil'], 'C1-C4': ['C2, fossil', 'C3, fossil', 'C4, fossil', 'C3-C4, biog'], 'D': ['D, tot']}
for k, cs in COLS.items():
    res[k] = res[cs].fillna(0).sum(axis=1)
res['el3'] = res['LEVELS Tier 3'].astype(str)
res['mat'] = res['INDICATE Materials Tier 2'].fillna('0.0 Undefined materials').astype(str)

lu_el, lu_mat = {}, {}
for _, r in meta.iterrows():
    p = r.proj
    sub = res[res.Project == p]
    if sub.empty:
        print('LU project without results:', p, file=sys.stderr); continue
    gfa, ufa = f(r.gfa), f(r.ufa)
    # ASSUMPTION: LU results are kgCO2e per m2 GFA (reference area not stated in file) -> absolute = value * GFA
    mods = {k: float(sub[k].sum()) for k in COLS}  # raw kgCO2e/m2 (reference area assumed in dashboard)
    yr = f(r.year)
    buildings.append(dict(
        id='LU-' + p, country='LU', natType=r.typ, type=HTYPE[r.typ], structure=(r.structure if isinstance(r.structure,str) else None),
        ufa=ufa, gfa=gfa, rsp=50.0, year=int(yr) if yr and yr > 1900 else None, heating=(r.energy if isinstance(r.energy,str) else None),
        mods=mods, perM2=True, nProducts=int(len(sub)),
    ))
    el = sub.groupby('el3')[list(COLS)].sum()
    lu_el['LU-' + p] = {c: {k: round(float(v), 5) for k, v in el.loc[c].items() if abs(v) > 0} for c in el.index}
    mt = sub.groupby('mat')[list(COLS)].sum()
    lu_mat['LU-' + p] = {c: {k: round(float(v), 5) for k, v in mt.loc[c].items() if abs(v) > 0} for c in mt.index}

# fill missing modules with None (not assessed) vs 0
for b in buildings:
    b['mods'] = {k: (round(b['mods'][k], 3) if k in b['mods'] else None) for k in MODS}

out = dict(modules=MODS, buildings=buildings, luElements=lu_el, luMaterials=lu_mat)

df = pd.DataFrame([{**{k: v for k, v in b.items() if k != 'mods'}, **b['mods']} for b in buildings])
print(df.groupby(['country', 'type']).size().to_string())
print(len(buildings), 'buildings')
# checks
h = df[df.country == 'HR']
chk = (h[['A1-A3','A4','A5','B3','B4-B5','B6','B7','C1-C4']].sum(axis=1) - h.srcTotal).abs().max()
print('HR max |sum(A-C) - source total|:', chk)

# Tier-3 element descriptions derived from the data: top-3 product categories by embodied GWP
res['emb'] = res[list(k for k in COLS if k != 'D')].sum(axis=1)
desc = {}
for c, g in res.groupby('el3'):
    top = g.groupby('Category class Mat intern')['emb'].sum().sort_values(ascending=False).head(3)
    desc[c] = [str(t) for t in top.index]
out['luTier3Desc'] = desc
dst = os.path.join(ROOT, 'assets', 'js', 'data.js')
with open(dst, 'w', encoding='utf-8') as fh:
    fh.write(f'/* INDICATE WLC Dashboard data, build {datetime.date.today()}. Generated by tools/build_data.py; do not edit by hand. PRELIMINARY: verification pending. */\n')
    fh.write('window.WLC_DATA = ' + json.dumps(out, ensure_ascii=False, separators=(',', ':')) + ';\n')
print('wrote', dst)
