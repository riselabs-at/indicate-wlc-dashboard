"""Write the harmonised database to an Excel workbook (and CSV copies).

Conventions in the workbook: blue text = values written by the pipeline (data), black text = Excel formulas.
The 'results' and 'stats_by_group' sheets are formula-driven, so they recalculate if a value is corrected.
"""
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.workbook.properties import CalcProperties
from config import MODULES, MODULE_LABELS, COUNTRY_NAMES

FONT = 'Arial'
F_BODY = Font(name=FONT, size=10)
F_DATA = Font(name=FONT, size=10, color='0000FF')
F_HEAD = Font(name=FONT, size=10, bold=True, color='FFFFFF')
F_TITLE = Font(name=FONT, size=14, bold=True)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color='555555')
FILL_HEAD = PatternFill('solid', fgColor='0D6A5A')
FILL_WIP = PatternFill('solid', fgColor='FFF4D6')
THIN = Border(bottom=Side(style='thin', color='CCCCCC'))
NUM0, NUM1, NUM2 = '#,##0;-#,##0;0', '#,##0.0;-#,##0.0;0', '#,##0.00;-#,##0.00;0'


def _sheet(wb, name, df, widths=None, num_formats=None, data_cols=None, note=None):
    ws = wb.create_sheet(name)
    r0 = 1
    if note:
        ws.cell(1, 1, note).font = F_NOTE
        r0 = 3
    for j, c in enumerate(df.columns, 1):
        cell = ws.cell(r0, j, c)
        cell.font, cell.fill = F_HEAD, FILL_HEAD
        cell.alignment = Alignment(wrap_text=True, vertical='top')
    for i, row in enumerate(df.itertuples(index=False), r0 + 1):
        for j, v in enumerate(row, 1):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            cell = ws.cell(i, j, v.item() if hasattr(v, 'item') else v)
            cell.font = F_DATA if (data_cols is None or df.columns[j - 1] in data_cols) else F_BODY
            fmt = (num_formats or {}).get(df.columns[j - 1])
            if fmt:
                cell.number_format = fmt
    ws.freeze_panes = ws.cell(r0 + 1, 2)
    ws.auto_filter.ref = f'A{r0}:{get_column_letter(len(df.columns))}{r0 + len(df)}'
    for j, c in enumerate(df.columns, 1):
        w = (widths or {}).get(c) or min(60, max(10, len(str(c)) + 2,
                                                 int(df[c].astype(str).str.len().quantile(0.9)) + 2 if len(df) else 10))
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[r0].height = 30
    return ws


def _text_sheet(wb, name, rows, title):
    ws = wb.create_sheet(name, 0) if name == 'README' else wb.create_sheet(name)
    ws.cell(1, 1, title).font = F_TITLE
    r = 3
    for kind, a, b in rows:
        if kind == 'h':
            ws.cell(r, 1, a).font = F_BOLD
        elif kind == 'wip':
            c = ws.cell(r, 1, a)
            c.font, c.fill = Font(name=FONT, size=10, bold=True, color='4A3500'), FILL_WIP
            ws.cell(r, 2).fill = FILL_WIP
        else:
            ws.cell(r, 1, a).font = F_BODY
            ws.cell(r, 2, b).font = F_BODY
            ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(r, 1).alignment = Alignment(vertical='top')
        r += 1
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 110
    return ws


def _results_sheet(wb, wide):
    """One row per building: absolute module values (data) + formulas for totals and per-m² values."""
    ws = wb.create_sheet('results')
    base = ['building_id', 'country', 'harmonised_type', 'intervention_class', 'gfa_m2', 'ufa_equivalent_m2', 'rsp_years']
    mods = [f'{m} kgCO2e' for m in MODULES]
    calc = ['Embodied kgCO2e', 'Whole life kgCO2e', 'Embodied kgCO2e/m² GFA', 'Embodied kgCO2e/m² UFA-eq',
            'Whole life kgCO2e/m² GFA', 'Whole life kgCO2e/m² UFA-eq', 'Embodied kgCO2e/(m² GFA·a)', 'Module D kgCO2e/m² GFA']
    head = base + mods + calc
    ws.cell(1, 1, 'Embodied = A1-A3 + A4 + A5 + B4-B5 + C1-C4 (modules all datasets report). Whole life = embodied + B2-B3 + B6 + B7, '
                  'blank when B6 is not reported. D is never included. UFA-eq = UFA (AT, HR, LU) or HFA (IT). Blue = data, black = formula.').font = F_NOTE
    for j, h in enumerate(head, 1):
        c = ws.cell(3, j, h)
        c.font, c.fill, c.alignment = F_HEAD, FILL_HEAD, Alignment(wrap_text=True, vertical='top')
    col = {h: get_column_letter(j) for j, h in enumerate(head, 1)}
    M = lambda m: col[f'{m} kgCO2e']
    for i, row in enumerate(wide.itertuples(index=False), 4):
        rec = dict(zip(wide.columns, row))
        for h in base:
            v = rec[h]
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                ws[f'{col[h]}{i}'] = v
                ws[f'{col[h]}{i}'].font = F_DATA
        for m in MODULES:
            v = rec[m]
            if v is not None and not pd.isna(v):
                ws[f'{M(m)}{i}'] = float(v)
                ws[f'{M(m)}{i}'].font, ws[f'{M(m)}{i}'].number_format = F_DATA, NUM0
        E, W = col['Embodied kgCO2e'], col['Whole life kgCO2e']
        G, U, Y = col['gfa_m2'], col['ufa_equivalent_m2'], col['rsp_years']
        f = {
            'Embodied kgCO2e': f'={M("A1-A3")}{i}+{M("A4")}{i}+{M("A5")}{i}+{M("B4-B5")}{i}+{M("C1-C4")}{i}',
            'Whole life kgCO2e': f'=IF({M("B6")}{i}="","",{E}{i}+N({M("B2-B3")}{i})+{M("B6")}{i}+N({M("B7")}{i}))',
            'Embodied kgCO2e/m² GFA': f'=IF({G}{i}="","",{E}{i}/{G}{i})',
            'Embodied kgCO2e/m² UFA-eq': f'=IF({U}{i}="","",{E}{i}/{U}{i})',
            'Whole life kgCO2e/m² GFA': f'=IF(OR({W}{i}="",{G}{i}=""),"",{W}{i}/{G}{i})',
            'Whole life kgCO2e/m² UFA-eq': f'=IF(OR({W}{i}="",{U}{i}=""),"",{W}{i}/{U}{i})',
            'Embodied kgCO2e/(m² GFA·a)': f'=IF({G}{i}="","",{E}{i}/{G}{i}/{Y}{i})',
            'Module D kgCO2e/m² GFA': f'=IF(OR({M("D")}{i}="",{G}{i}=""),"",{M("D")}{i}/{G}{i})',
        }
        for h, formula in f.items():
            c = ws[f'{col[h]}{i}']
            c.value, c.font = formula, F_BODY
            c.number_format = NUM0 if 'kgCO2e' == h.split()[-1] else NUM1
    last = 3 + len(wide)
    ws.freeze_panes = 'B4'
    ws.auto_filter.ref = f'A3:{get_column_letter(len(head))}{last}'
    for j, h in enumerate(head, 1):
        ws.column_dimensions[get_column_letter(j)].width = 18 if j > 1 else 26
    ws.column_dimensions[col['harmonised_type']].width = 28
    ws.row_dimensions[3].height = 42
    return ws, col, last


def _stats_sheet(wb, wide, col, last):
    """Benchmark statistics per country × type × intervention, as formulas over the results sheet."""
    ws = wb.create_sheet('stats_by_group')
    ws.cell(1, 1, 'Statistics per group, calculated with formulas from the results sheet. P25 / median / P75 are the '
                  'statistical benchmark-class limits used in the dashboard (Target ≤ P25, Reference ≤ median, Limit ≤ P75). '
                  'PERCENTILE is inclusive (= PERCENTILE.INC). Groups with fewer than 5 buildings are not classified in the dashboard.').font = F_NOTE
    metrics = [('Embodied kgCO2e/m² GFA', 'Emb/GFA'), ('Embodied kgCO2e/m² UFA-eq', 'Emb/UFA-eq'),
               ('Whole life kgCO2e/m² GFA', 'WL/GFA'), ('Whole life kgCO2e/m² UFA-eq', 'WL/UFA-eq')]
    stats = [('n', None), ('min', 0), ('P25', 0.25), ('median', 0.5), ('P75', 0.75), ('max', 1)]
    head = ['country', 'harmonised_type', 'intervention_class'] + [f'{short} {s}' for _, short in metrics for s, _ in stats]
    for j, h in enumerate(head, 1):
        c = ws.cell(3, j, h)
        c.font, c.fill, c.alignment = F_HEAD, FILL_HEAD, Alignment(wrap_text=True, vertical='top')
    groups = wide.groupby(['country', 'harmonised_type', 'intervention_class']).size().reset_index()[
        ['country', 'harmonised_type', 'intervention_class']]
    rng = lambda h: f"results!${col[h]}$4:${col[h]}${last}"
    for i, g in enumerate(groups.itertuples(index=False), 4):
        for j, v in enumerate(g, 1):
            ws.cell(i, j, v).font = F_DATA
        cond = (f"({rng('country')}=$A{i})*({rng('harmonised_type')}=$B{i})*({rng('intervention_class')}=$C{i})")
        j = 4
        for h, _ in metrics:
            n_ref = f'{get_column_letter(j)}{i}'
            for s, p in stats:
                ref = f'{get_column_letter(j)}{i}'
                if p is None:
                    ws[ref] = f'=SUMPRODUCT({cond}*ISNUMBER({rng(h)}))'
                    ws[ref].number_format = '0'
                else:
                    ws[ref] = ArrayFormula(ref, f'=IF({n_ref}=0,"",PERCENTILE(IF({cond}*ISNUMBER({rng(h)}),{rng(h)}),{p}))')
                    ws[ref].number_format = NUM0
                ws[ref].font = F_BODY
                j += 1
    ws.freeze_panes = 'D4'
    ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['C'].width = 9, 30, 18
    for j in range(4, len(head) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 11
    ws.row_dimensions[3].height = 42
    return ws


def export(d, xlsx_path, csv_dir=None):
    m = d['meta']
    wb = Workbook()
    wb.remove(wb.active)
    b = d['buildings']
    counts = b.groupby('country').size()
    readme = [
        ('wip', 'WORK IN PROGRESS: preliminary harmonisation, verification with the national teams pending. Do not cite or reuse.', ''),
        ('', '', ''),
        ('row', 'Database version', f"{m['version']} (built {m['built']})"),
        ('row', 'Buildings', f"{len(b)} in total: " + ', '.join(f"{COUNTRY_NAMES[c]} {n}" for c, n in counts.items())
         + f"; {int((b.intervention_class == 'Renovation').sum())} of them renovations (Italy)."),
        ('row', 'Project', 'INDICATE LIFE – National Building LCA Data Accelerator EU LIFE (grant agreement 101167585)'),
        ('row', 'Indicator', 'GWP-total (fossil + biogenic + LULUC) in kgCO2e, 50-year reference study period, per building and life-cycle module.'),
        ('row', 'Embodied scope', ' + '.join(m['embodied']) + ' (the modules reported by all four datasets).'),
        ('row', 'Whole-life scope', 'Embodied + B2-B3 + B6 + B7 where reported. Module D is always kept separately.'),
        ('row', 'Areas', 'GFA = gross floor area. UFA-eq = usable floor area (AT, HR, LU) or heated floor area HFA (IT, no UFA reported).'),
        ('row', 'Conventions', 'Blue text = data written by the pipeline. Black text = Excel formulas (results, stats_by_group). Blank = not reported.'),
        ('', '', ''),
        ('h', 'Sheets', ''),
        ('row', 'processing_steps', 'The harmonisation steps, followed by the per-country processing log with row counts.'),
        ('row', 'sources', 'The four national deliverables and source files.'),
        ('row', 'buildings', 'One row per building: IDs, national and harmonised typology, intervention, floor areas, metadata.'),
        ('row', 'results', 'One row per building: absolute module results (data) and embodied / whole-life totals and per-m² values (formulas).'),
        ('row', 'stats_by_group', 'n, min, P25, median, P75, max per country × type × intervention (formulas). These are the benchmark-class limits.'),
        ('row', 'lcm_long', 'Tidy table: one row per building × module with the source value, source unit, source columns and harmonised kgCO2e.'),
        ('row', 'elements_LU', 'Luxembourg: per building × LEVELS Tier 3 code, kgCO2e per m² (assumed GFA) by module, summed from product rows.'),
        ('row', 'elements_IT', 'Italy: per building × tier, kgCO2e/m² GFA as in the source, with a quality flag read from the cell colours.'),
        ('row', 'materials_LU', 'Luxembourg: per building × INDICATE Materials Tier 2 group, kgCO2e per m² (assumed GFA) by module.'),
        ('row', 'element_codes', 'Element codes per classification. LU codes described by their three largest product categories.'),
        ('row', 'map_modules', 'Source column → harmonised module, per country.'),
        ('row', 'map_typologies', 'National typology → harmonised building type, per country.'),
        ('row', 'issues', 'Assumptions and data-quality observations. Status "open – partner to confirm" needs a decision.'),
        ('row', 'qa_checks', 'Reconciliation checks run by the pipeline, with tolerance and result.'),
        ('', '', ''),
        ('row', 'Reproduce', 'python pipeline/run_all.py <folder with the four source workbooks>'),
    ]
    _text_sheet(wb, 'README', readme, 'INDICATE WLC database')

    steps = d['steps']
    ws = _sheet(wb, 'processing_steps', steps, widths={'step': 6, 'name': 16, 'description': 120}, data_cols=[])
    for row in ws.iter_rows(min_row=2, max_row=1 + len(steps)):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical='top')
    r = len(steps) + 4
    ws.cell(r, 1, 'Processing log').font = F_BOLD
    for j, h in enumerate(['country', 'action', 'detail', 'result'], 1):
        c = ws.cell(r + 1, j, h)
        c.font, c.fill = F_HEAD, FILL_HEAD
    for i, row in enumerate(d['log'].itertuples(index=False), r + 2):
        for j, v in enumerate(row, 1):
            ws.cell(i, j, v).font = F_BODY
            ws.cell(i, j).alignment = Alignment(wrap_text=True, vertical='top')
    ws.column_dimensions['D'].width = 50
    ws.auto_filter.ref = None

    _sheet(wb, 'sources', d['sources'], data_cols=[])
    _sheet(wb, 'buildings', d['buildings'], num_formats={'gfa_m2': NUM1, 'ufa_m2': NUM1, 'hfa_m2': NUM1,
           'ufa_equivalent_m2': NUM1, 'gwp_fossil_kgco2e': NUM0, 'gwp_biogenic_kgco2e': NUM0, 'gwp_luluc_kgco2e': NUM0},
           widths={'building_id': 26, 'source_id': 24, 'harmonised_type': 30, 'national_typology': 24})
    ws, col, last = _results_sheet(wb, d['wide'])
    _stats_sheet(wb, d['wide'], col, last)
    _sheet(wb, 'lcm_long', d['lcm'], num_formats={'source_value': NUM2, 'value_kgco2e': NUM0},
           data_cols=list(d['lcm'].columns), widths={'source_columns': 44, 'source_unit': 36, 'building_id': 26})
    nm = {c: NUM2 for c in d['lu_elements'].columns if c in MODULES}
    _sheet(wb, 'elements_LU', d['lu_elements'], num_formats=nm,
           note='kgCO2e per m² as in the source (reference area not stated; assumed GFA). Columns = harmonised modules.')
    _sheet(wb, 'elements_IT', d['it_elements'], num_formats={'value_kgco2e_m2': NUM2},
           widths={'note': 50, 'quality': 30, 'building_id': 26},
           note='kgCO2e per m² GFA from sheet "SumUP Tiers" (GFA half). quality = meaning of the cell colour per the source legend; source_cell = original cell.')
    nm = {c: NUM2 for c in d['lu_materials'].columns if c in MODULES}
    _sheet(wb, 'materials_LU', d['lu_materials'], num_formats=nm, widths={'material_group': 44},
           note='kgCO2e per m² as in the source (assumed GFA), summed per INDICATE Materials Tier 2 group.')
    _sheet(wb, 'element_codes', d['codes'], widths={'description': 90, 'classification': 24}, data_cols=[])
    _sheet(wb, 'map_modules', d['maps']['modules'], widths={'note': 80, 'source_columns': 44}, data_cols=[])
    _sheet(wb, 'map_typologies', d['maps']['typologies'], widths={'note': 90, 'national_typology': 40, 'harmonised_type': 32}, data_cols=[])
    ws = _sheet(wb, 'issues', d['issues'], widths={'observation': 70, 'handling': 60, 'building_ids': 30, 'status': 26}, data_cols=[])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical='top')
    _sheet(wb, 'qa_checks', d['checks'], widths={'check': 90}, data_cols=[], num_formats={'value': '0.000E+00', 'tolerance': '0.0E+00'})

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = ws.title not in ('README',)
    wb.calculation = CalcProperties(fullCalcOnLoad=True)   # Excel computes the formulas when the file is opened
    os.makedirs(os.path.dirname(xlsx_path), exist_ok=True)
    wb.save(xlsx_path)

    if csv_dir:
        os.makedirs(csv_dir, exist_ok=True)
        for name, df in [('buildings', d['buildings']), ('lcm_long', d['lcm']), ('elements_lu', d['lu_elements']),
                         ('elements_it', d['it_elements']), ('materials_lu', d['lu_materials']),
                         ('issues', d['issues']), ('qa_checks', d['checks'])]:
            df.to_csv(os.path.join(csv_dir, f'{name}.csv'), index=False, encoding='utf-8')
