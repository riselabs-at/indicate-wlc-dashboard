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
from config import MODULES, DISPLAY_ORDER, SOURCES
COUNTRY_NAMES = {c: s['country'] for c, s in SOURCES.items()}

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
        q = df[c].astype(str).str.len().quantile(0.9) if len(df) else 10
        w = (widths or {}).get(c) or min(60, max(10, len(str(c)) + 2, int(q) + 2 if q == q else 10))
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
    base = ['building_id', 'country', 'typology_code', 'project_intervention_type', 'project_renovation_depth', 'bldg_area_gfa', 'bldg_area_nfa', 'lca_rsp']
    mods = [f'{m} kgCO2e' for m in DISPLAY_ORDER]
    calc = ['Embodied kgCO2e', 'Whole life kgCO2e', 'Embodied kgCO2e/m² GFA', 'Embodied kgCO2e/m² NFA',
            'Whole life kgCO2e/m² GFA', 'Whole life kgCO2e/m² NFA', 'Embodied kgCO2e/(m² GFA·a)', 'Module D kgCO2e/m² GFA']
    head = base + mods + calc
    ws.cell(1, 1, 'Display modules summed from the template modules in wlc_results (GWPtotal). Embodied = A1-A3 + A4 + A5 + B4-B5 + C1-C4 (reported by all datasets). '
                  'Whole life = embodied + B1-B3 + B6 + B7, blank when B6 is not reported. D is never included. NFA = UFA (AT, HR, LU) or HFA (IT). '
                  'lca_rsp: 50 assumed where not reported (LU). Blue = data, black = formula.').font = F_NOTE
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
        for m in DISPLAY_ORDER:
            v = rec[m]
            if v is not None and not pd.isna(v):
                ws[f'{M(m)}{i}'] = float(v)
                ws[f'{M(m)}{i}'].font, ws[f'{M(m)}{i}'].number_format = F_DATA, NUM0
        E, W = col['Embodied kgCO2e'], col['Whole life kgCO2e']
        G, U, Y = col['bldg_area_gfa'], col['bldg_area_nfa'], col['lca_rsp']
        f = {
            'Embodied kgCO2e': f'={M("A1-A3")}{i}+{M("A4")}{i}+{M("A5")}{i}+{M("B4-B5")}{i}+{M("C1-C4")}{i}',
            'Whole life kgCO2e': f'=IF({M("B6")}{i}="","",{E}{i}+N({M("B1-B3")}{i})+{M("B6")}{i}+N({M("B7")}{i}))',
            'Embodied kgCO2e/m² GFA': f'=IF({G}{i}="","",{E}{i}/{G}{i})',
            'Embodied kgCO2e/m² NFA': f'=IF({U}{i}="","",{E}{i}/{U}{i})',
            'Whole life kgCO2e/m² GFA': f'=IF(OR({W}{i}="",{G}{i}=""),"",{W}{i}/{G}{i})',
            'Whole life kgCO2e/m² NFA': f'=IF(OR({W}{i}="",{U}{i}=""),"",{W}{i}/{U}{i})',
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
    ws.column_dimensions[col['project_renovation_depth']].width = 30
    ws.row_dimensions[3].height = 42
    return ws, col, last


def _stats_sheet(wb, wide, col, last):
    """Benchmark statistics per country × type × intervention, as formulas over the results sheet."""
    ws = wb.create_sheet('stats_by_group')
    ws.cell(1, 1, 'Statistics per group, calculated with formulas from the results sheet. P25 / median / P75 are the '
                  'statistical benchmark-class limits used in the dashboard (Best practice ≤ P25, Reference ≤ median, Limit ≤ P75; terms of D3.1/D3.2). '
                  'PERCENTILE is inclusive (= PERCENTILE.INC). Groups with fewer than 5 buildings are not classified in the dashboard.').font = F_NOTE
    metrics = [('Embodied kgCO2e/m² GFA', 'Emb/GFA'), ('Embodied kgCO2e/m² NFA', 'Emb/NFA'),
               ('Whole life kgCO2e/m² GFA', 'WL/GFA'), ('Whole life kgCO2e/m² NFA', 'WL/NFA')]
    stats = [('n', None), ('min', 0), ('P25', 0.25), ('median', 0.5), ('P75', 0.75), ('max', 1)]
    head = ['country', 'typology_code', 'project_intervention_type'] + [f'{short} {s}' for _, short in metrics for s, _ in stats]
    for j, h in enumerate(head, 1):
        c = ws.cell(3, j, h)
        c.font, c.fill, c.alignment = F_HEAD, FILL_HEAD, Alignment(wrap_text=True, vertical='top')
    groups = wide.groupby(['country', 'typology_code', 'project_intervention_type']).size().reset_index()[
        ['country', 'typology_code', 'project_intervention_type']]
    rng = lambda h: f"results!${col[h]}$4:${col[h]}${last}"
    for i, g in enumerate(groups.itertuples(index=False), 4):
        for j, v in enumerate(g, 1):
            ws.cell(i, j, v).font = F_DATA
        cond = (f"({rng('country')}=$A{i})*({rng('typology_code')}=$B{i})*({rng('project_intervention_type')}=$C{i})")
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
    p = d['projects']
    counts = p.groupby('country').size()
    readme = [
        ('wip', 'WORK IN PROGRESS: preliminary harmonisation, verification with the national teams pending. Do not cite or reuse.', ''),
        ('', '', ''),
        ('row', 'Database version', f"{m['version']} (built {m['built']})"),
        ('row', 'Schema', m['template'] + '. Attribute codes, life-cycle modules, typology codes and element tiers follow the template; '
                          'fields the template does not have start with x_.'),
        ('row', 'Buildings', f"{len(p)}: " + ', '.join(f"{COUNTRY_NAMES[c]} {n}" for c, n in counts.items())
         + f"; {int((p.project_intervention_type == 'Renovation').sum())} renovations (Italy)."),
        ('row', 'Project', 'INDICATE LIFE – National Building LCA Data Accelerator EU LIFE (grant agreement 101167585)'),
        ('row', 'Indicator', 'GWP in kgCO2e over a 50-year reference study period: GWPtotal per template module; GWPfossil / GWPbio / GWPluluc where reported.'),
        ('row', 'Embodied scope', ' + '.join(m['embodied']) + ' (display modules reported by all four datasets). Whole life adds B1-B3, B6, B7. D kept separately.'),
        ('row', 'Areas', 'GFA as reported. NFA = useful floor area (UFA) for AT, HR, LU and heated floor area (HFA) for IT, which reports no UFA.'),
        ('row', 'Combined values', 'A value reported for several modules (e.g. B4-B5) is stored in the first module of the range; the range is in "covers".'),
        ('row', 'Conventions', 'Blue text = data written by the pipeline. Black text = Excel formulas (results, stats_by_group). Blank = not reported.'),
        ('', '', ''),
        ('h', 'Sheets', ''),
        ('row', 'processing_steps', 'The harmonisation steps, then the per-country processing log with row counts.'),
        ('row', 'sources', 'The four national deliverables and source files.'),
        ('row', 'codebook', 'Every attribute (template code or x_ extension) with priority and the share of buildings per country that have a value.'),
        ('row', 'projects', 'One row per building in template attribute codes (Project-information / Multi-Project_Overview).'),
        ('row', 'results', 'One row per building: display-module totals (data) and embodied / whole-life totals and per-m² values (formulas).'),
        ('row', 'stats_by_group', 'n, min, P25, median, P75, max per country × typology × intervention (formulas) – the benchmark-class limits.'),
        ('row', 'wlc_results', 'Tidy table: building × indicator × template module, with source value, unit, columns, "covers" and flags.'),
        ('row', 'elements', 'Element results (LU, IT) per national code, mapped to EU Table 4 tiers (template v1.1), with mapping status and quality flags.'),
        ('row', 'elements_eu_tier2', 'Embodied element results per building aggregated to EU Tier 2 (kgCO2e/m² as in the source).'),
        ('row', 'materials_LU', 'Luxembourg: per building × material group (template material groups) × module, kgCO2e/m².'),
        ('row', 'eu_element_tiers', 'EU Delegated Regulation Table 4 tiers with INDICATE codes, as in template v1.1.'),
        ('row', 'map_typologies / map_modules / map_elements', 'National typology → template code; source column → template module; national element code → EU tier.'),
        ('row', 'issues', 'Assumptions and data-quality observations with the question to the national team. "open – partner to confirm" needs a decision.'),
        ('row', 'qa_checks', 'Reconciliation checks run by the pipeline, with tolerance and result.'),
        ('', '', ''),
        ('row', 'Reproduce', 'python pipeline/run_all.py <folder with the four source workbooks>'),
    ]
    _text_sheet(wb, 'README', readme, 'INDICATE WLC database')

    steps = d['steps']
    ws = _sheet(wb, 'processing_steps', steps, widths={'step': 6, 'name': 26, 'description': 120}, data_cols=[])
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

    _sheet(wb, 'sources', d['sources'], data_cols=[], widths={'title': 60, 'contact_org': 50})
    cb = d['codebook']
    _sheet(wb, 'codebook', cb, data_cols=[], widths={'label': 52, 'attribute_code': 34},
           num_formats={f'coverage_{c}': '0%' for c in ['AT', 'HR', 'IT', 'LU']},
           note='Coverage = share of buildings with a value. HP = high priority in the template. For GWPtotal modules, "covers" notes show combined values.')
    pc = d['projects'].drop(columns=['country', 'typology_code'])
    _sheet(wb, 'projects', pc, num_formats={'bldg_area_gfa': NUM1, 'bldg_area_nfa': NUM1, 'x_gfa_below_m2': NUM1, 'x_gfa_above_m2': NUM1},
           widths={'building_id': 24, 'project_code': 24, 'bldg_use_subtype': 34, 'contact_org': 40, 'bldg_area_references': 50})
    ws, col, last = _results_sheet(wb, d['wide'])
    _stats_sheet(wb, d['wide'], col, last)
    _sheet(wb, 'wlc_results', d['results'], num_formats={'source_value': NUM2, 'value_kgco2e': NUM0},
           data_cols=list(d['results'].columns), widths={'source_columns': 40, 'source_unit': 34, 'building_id': 24, 'flag': 40})
    _sheet(wb, 'elements', d['elements'], num_formats={'value_kgco2e_m2': NUM2, 'value_kgco2e': NUM0},
           widths={'eu_name': 40, 'mapping_basis': 44, 'note': 40, 'quality': 26, 'building_id': 24, 'classification': 22},
           note='value_kgco2e_m2 as in the source (LU: per m² UFA, as in D3.2; IT: per m² GFA, tiers incl. proxies). value_kgco2e = × UFA (LU) or × GFA (IT). '
                'LU rows are per template module; IT rows are A–C as reported.')
    _sheet(wb, 'elements_eu_tier2', d['tier2'], num_formats={c: NUM1 for c in d['tier2'].columns if c[:1].isdigit()},
           note='Embodied element results (LU: A1-A3, A4, A5, B4, C2–C4; IT: tiers as reported) summed per EU Tier 2 code, kgCO2e/m² as in the source.')
    _sheet(wb, 'materials_LU', d['materials'], num_formats={c: NUM2 for c in d['materials'].columns if c in MODULES},
           widths={'material_group': 44}, note='kgCO2e per m² UFA as in the source (area basis per D3.2), per template material group and module.')
    _sheet(wb, 'eu_element_tiers', d['eu_elements'], data_cols=[], widths={'name': 80})
    _sheet(wb, 'map_typologies', d['map_typologies'], data_cols=[], widths={'note': 80, 'national_typology': 40, 'template_label': 36})
    _sheet(wb, 'map_modules', d['map_modules'], data_cols=[], widths={'source_columns': 50, 'flag': 50, 'source_unit': 34})
    _sheet(wb, 'map_elements', d['map_elements'], data_cols=[], widths={'basis': 60, 'eu_name': 50, 'national_content (LU: top-3 product categories)': 80})
    ws = _sheet(wb, 'issues', d['issues'], data_cols=[],
                widths={'observation': 60, 'handling': 50, 'building_ids': 28, 'status': 24, 'question_to_national_team': 60})
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical='top')
    _sheet(wb, 'qa_checks', d['checks'], widths={'check': 90}, data_cols=[], num_formats={'value': '0.000E+00', 'tolerance': '0.0E+00'})
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = ws.title != 'README'
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    os.makedirs(os.path.dirname(xlsx_path), exist_ok=True)
    wb.save(xlsx_path)
    if csv_dir:
        os.makedirs(csv_dir, exist_ok=True)
        for name in ['projects', 'results', 'elements', 'elements_eu_tier2', 'materials', 'codebook', 'issues', 'checks',
                     'map_typologies', 'map_modules', 'map_elements']:
            key = {'results': 'results', 'elements_eu_tier2': 'tier2', 'checks': 'checks'}.get(name, name)
            out = {'results': 'wlc_results', 'materials': 'materials_lu', 'checks': 'qa_checks'}.get(name, name)
            d[key].to_csv(os.path.join(csv_dir, f'{out}.csv'), index=False, encoding='utf-8')
