"""Write one pre-filled copy of the INDICATE reporting template v1.1 per country (step 3: data request to national teams).

Each copy has:
  - 'Data request'          – what was pre-filled, the colour code, and what we ask for
  - 'Questions'             – open issues for this country with an answer column
  - 'Element mapping'       – (IT, LU) national element codes → EU Table 4 tiers, to confirm
  - 'Multi-Project_Overview'– one row per building, filled from the database; empty high-priority cells marked
"""
import os
from copy import copy
import openpyxl
import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter as L
from config import SOURCES, MODULES, DATABASE_VERSION

ASK = PatternFill('solid', fgColor='FFF4B6C2')      # please complete (high priority, empty)
CHECK = PatternFill('solid', fgColor='FFFFF2A8')    # pre-filled from an assumption – please check
AUTHOR = 'INDICATE WLC database'
HP_HEADERS = ['Contact email', 'Project code (custom)', 'Intervention type (select)', 'Renovation depth (select)', 'Country name (select)',
              'Country code (select)', 'Building typology (select)', 'Number of users (n)', 'Gross Floor Area (m2)', 'Net Floor Area (m2)',
              'Floors above ground (n)', 'Floors below ground (n)', 'Building height (m)', 'Energy consumption per year (kWh/a)',
              'Structure type (select)', 'Total mass (kg)', 'Reference study period (RSP)']
FIELD = {'Project code (custom)': 'project_code', 'Intervention type (select)': 'project_intervention_type',
         'Renovation depth (select)': 'project_renovation_depth', 'Country name (select)': 'site_country',
         'Country code (select)': 'site_country_iso3', 'Building typology (select)': 'bldg_use_subtype',
         'Number of users (n)': 'bldg_users_total', 'Gross Floor Area (m2)': 'bldg_area_gfa', 'Net Floor Area (m2)': 'bldg_area_nfa',
         'Floors above ground (n)': 'bldg_floors_ag', 'Floors below ground (n)': 'bldg_floors_bg', 'Building height (m)': 'bldg_height',
         'Energy consumption per year (kWh/a)': 'inv_energy_consumption', 'Structure type (select)': 'bldg_struct_type',
         'Total mass (kg)': 'inv_mat_mass_total', 'Reference study period (RSP)': 'lca_rsp'}


def _v(x):
    return None if x is None or (isinstance(x, float) and pd.isna(x)) else (x.item() if hasattr(x, 'item') else x)


def _table(ws, top, head, rows, widths):
    for j, h in enumerate(head, 2):
        c = ws.cell(top, j, h)
        c.font, c.fill = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='FF0D6A5A')
        c.alignment = Alignment(wrap_text=True, vertical='top')
    for i, row in enumerate(rows, top + 1):
        for j, v in enumerate(row, 2):
            c = ws.cell(i, j, _v(v))
            c.alignment = Alignment(wrap_text=True, vertical='top')
            if head[j - 2].startswith('Answer') or head[j - 2].startswith('Confirm'):
                c.fill = PatternFill('solid', fgColor='FFFFCC99')
    for j, w in enumerate(widths, 2):
        ws.column_dimensions[L(j)].width = w


def export(d, template_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    p, r, iss, me = d['projects'], d['results'], d['issues'], d['map_elements']
    tot = r[r.indicator == 'GWPtotal']
    written = []
    for cc, src in SOURCES.items():
        wb = openpyxl.load_workbook(template_path)
        mp = wb['Multi-Project_Overview']
        head = {mp.cell(2, j).value: j for j in range(1, mp.max_column + 1)}
        mod_col = {}
        for j in range(1, mp.max_column + 1):                       # first results block = absolute values
            h = mp.cell(2, j).value
            if h in MODULES and h not in mod_col:
                mod_col[h] = j
        pc = p[p.country == cc].reset_index(drop=True)
        assert len(pc) <= 101, 'more buildings than template rows'
        n_ask = n_check = 0
        for i, b in pc.iterrows():
            row = 3 + i
            notes = []
            mp.row_dimensions[row].height = 15
            for h, f in FIELD.items():
                c = mp.cell(row, head[h])
                c.value = _v(b[f])
            mp.cell(row, head['NFA definition (e.g. UFA, HFA)']).value = 'HFA' if cc == 'IT' else 'UFA'
            # assumptions → yellow + comment
            if 'assumed' in str(b.x_intervention_basis):
                for h in ('Intervention type (select)', 'Renovation depth (select)'):
                    c = mp.cell(row, head[h])
                    c.fill, c.comment = copy(CHECK), Comment(f'Pre-filled: {b.x_intervention_basis}. Please check.', AUTHOR)
                    n_check += 1
            if b.x_typology_national and 'to confirm' in d['map_typologies'].set_index(['country', 'national_typology']).status.get((cc, b.x_typology_national), ''):
                c = mp.cell(row, head['Building typology (select)'])
                c.fill, c.comment = copy(CHECK), Comment(f'National typology "{b.x_typology_national}" mapped to {b.bldg_use_subtype}. Please confirm.', AUTHOR)
                n_check += 1
            if _v(b.x_users_basis) and 'estimated' in b.x_users_basis:
                c = mp.cell(row, head['Number of users (n)'])
                c.fill, c.comment = copy(CHECK), Comment('Equals UFA ÷ 10 – looks estimated. Please check.', AUTHOR)
                n_check += 1
            if cc in ('HR', 'IT') or (cc == 'LU'):
                c = mp.cell(row, head['Reference study period (RSP)'])
                if cc == 'LU':
                    c.value = None
                    c.comment = Comment('Not reported (product service lives only). 50 years assumed in the database.', AUTHOR)
                else:
                    c.comment = Comment(f'{b.x_rsp_basis}', AUTHOR)
            if _v(b.x_struct_material_national):
                mp.cell(row, head['Structure type (select)']).comment = Comment(
                    f'Reported main material: "{b.x_struct_material_national}". Please choose frame / massive + material.', AUTHOR)
            if _v(b.x_primary_energy_kwh_m2a) is not None:
                mp.cell(row, head['Energy consumption per year (kWh/a)']).comment = Comment(
                    f'Reported primary energy demand {b.x_primary_energy_kwh_m2a:.1f} kWh/m²·a per heated area An = {b.x_heated_area_an_m2 or "?"} m². '
                    'Please give final energy consumption of the building in kWh/a.', AUTHOR)
            if _v(b.x_storeys_total) is not None:
                mp.cell(row, head['Floors above ground (n)']).comment = Comment(f'Reported: {b.x_storeys_total:.0f} storeys in total (not split).', AUTHOR)
            # results
            rb = tot[tot.building_id == b.building_id]
            for res in rb.itertuples(index=False):
                if res.value_kgco2e is None or pd.isna(res.value_kgco2e):
                    continue
                c = mp.cell(row, mod_col[res.module])
                c.value = float(res.value_kgco2e)
                if res.covers != res.module or res.flag:
                    c.comment = Comment('; '.join(x for x in [f'covers {res.covers}' if res.covers != res.module else '', res.flag] if x), AUTHOR)
                    notes.append(f'{res.module}: ' + (f'covers {res.covers}' if res.covers != res.module else res.flag))
            a, z = L(mod_col['A1-A3']), L(mod_col['C4'])
            mp.cell(row, head['GWPtotal']).value = f'=IF(COUNT({a}{row}:{z}{row})=0,"",SUM({a}{row}:{z}{row}))'
            if cc == 'LU':
                notes.insert(0, 'Absolute values = per-m² results × UFA (area basis not stated in the file; per m² UFA as in D3.2 – please confirm)')
            mp.cell(row, head['Notes']).value = ' | '.join(dict.fromkeys(notes)) or None
            for h in HP_HEADERS:
                c = mp.cell(row, head[h])
                if c.value in (None, ''):
                    c.fill = copy(ASK)
                    n_ask += 1
        mp.cell(1, head['GWPtotal']).value = 'Results (GWPtotal) for whole building (kgCO2e) (not per m²). GWPtotal = sum of A1-A3 … C4 (formula, excl. D).'

        # ---- data request sheet
        dr = wb.create_sheet('Data request', 0)
        dr['B2'] = f'INDICATE WLC database – data request: {src["country"]} ({src["deliverable"]})'
        dr['B2'].font = Font(bold=True, size=16)
        lines = [
            f'This copy of the INDICATE reporting template v1.1 is pre-filled with the {len(pc)} buildings of deliverable {src["deliverable"]} '
            f'as harmonised in the INDICATE WLC database v{DATABASE_VERSION}. Please review it and complete what is missing.',
            '',
            'Where to look',
            '• Multi-Project_Overview – one row per building. Life-cycle results are the absolute GWPtotal per module (kgCO2e) as reported; per-m² columns calculate automatically.',
            '• Questions – open points for your dataset. Please answer in the orange column.',
            '• Element mapping – (Italy, Luxembourg) how your element codes were mapped to the EU Table 4 tiers. Please confirm or correct.',
            '',
            'Colour code in Multi-Project_Overview',
            '• Pink – high-priority field that is still empty. Please fill in where the information exists.',
            '• Yellow with a comment – pre-filled from an assumption. Please check.',
            '• Cell comments – source notes, e.g. combined modules ("covers B4-B5"), GWP-fossil-only values, reported values that do not fit the template field.',
            '',
            'Template conventions',
            '• A value covering several modules (e.g. B4-B5) is entered in the first module of the range; the range is noted in a comment and the Notes column.',
            '• Net Floor Area = useful floor area (UFA) or, if not available, heated floor area (HFA); state which in "NFA definition".',
            '• Energy consumption per year in kWh/a for the whole building.',
            '',
            f'Summary for this file: {n_ask} empty high-priority cells marked pink, {n_check} pre-filled assumptions marked yellow.',
        ]
        for i, t in enumerate(lines, 4):
            c = dr.cell(i, 2, t)
            c.alignment = Alignment(wrap_text=True, vertical='top')
            if t in ('Where to look', 'Colour code in Multi-Project_Overview', 'Template conventions'):
                c.font = Font(bold=True)
        for i, (fill, label) in enumerate([(ASK, 'Please complete'), (CHECK, 'Please check (assumption)')], len(lines) + 6):
            dr.cell(i, 2, label).fill = copy(fill)
        dr.column_dimensions['A'].width, dr.column_dimensions['B'].width = 3, 140

        # ---- questions
        q = wb.create_sheet('Questions', 1)
        q['B2'] = f'Open questions – {src["country"]}'
        q['B2'].font = Font(bold=True, size=14)
        qi = iss[(iss.country == cc) & (iss.question_to_national_team.fillna('') != '')]
        _table(q, 4, ['ID', 'Buildings', 'Observation', 'Current handling', 'Question', 'Answer (please fill)'],
               qi[['issue_id', 'building_ids', 'observation', 'handling', 'question_to_national_team']].assign(a=None).values.tolist(),
               [7, 26, 50, 40, 50, 50])

        # ---- element mapping
        if cc in ('IT', 'LU'):
            em = wb.create_sheet('Element mapping', 2)
            em['B2'] = f'Element codes of {src["country"]} → EU Table 4 tiers (template tab "WLC-emissions_by-element") – please confirm'
            em['B2'].font = Font(bold=True, size=14)
            rows = me[me.country == cc][['national_code', 'national_content (LU: top-3 product categories)', 'eu_code', 'eu_name', 'basis', 'status']]
            rows = rows.assign(c=None).values.tolist()
            _table(em, 4, ['Your code', 'Content found (LU: largest product categories)', 'EU code', 'EU Table 4 element', 'Basis of mapping',
                           'Status', 'Confirm / correct (please fill)'], rows, [11, 60, 9, 50, 50, 11, 36])

        out = os.path.join(out_dir, f'INDICATE_WLC_data-request_{cc}_{src["deliverable"]}.xlsx')
        wb.save(out)
        written.append((out, len(pc), n_ask, n_check))
    return written
