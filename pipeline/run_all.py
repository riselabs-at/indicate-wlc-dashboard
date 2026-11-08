"""Run the full INDICATE WLC harmonisation pipeline.

Usage (from the repository root):
    pip install -r pipeline/requirements.txt
    python pipeline/run_all.py path/to/source-folder

The source folder must hold the four national deliverable workbooks (not stored in the repository):
    *IND-AT*WLC*.xlsx   D7.3  Austria
    *HR_WLC*.xlsx       D7.12 Croatia
    *D7.21*.xlsx        D7.21 Italy
    *LU_Export*.xlsx    D7.30 Luxembourg

Outputs:
    data/INDICATE_WLC_database.xlsx      harmonised database in template v1.1 vocabulary (formulas recalculate on open)
    data/csv/*.csv                       the main tables as CSV
    data/data-requests/*.xlsx            pre-filled template v1.1 per country, to send to the national teams
    assets/js/data.js                    data file read by the dashboard (index.html)
"""
import os, sys, warnings
warnings.filterwarnings('ignore', category=UserWarning)   # openpyxl drawing warnings from the source files
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import harmonise, export_xlsx, export_web, export_template  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, 'template', 'INDICATE_Building-WLC-Emissions-Reporting_Template_v1.1.xlsx')


def main(folder):
    d = harmonise.build(folder)
    print('Processing log')
    for r in d['log'].itertuples(index=False):
        print(f'  [{r.country}] {r.action:9s} {r.detail}  →  {r.result}')
    print('\nChecks')
    for r in d['checks'].itertuples(index=False):
        print(f"  [{r.country}] {'PASS' if r.passed else 'FAIL'}  {r.check}: {r.value:.4g}")
    if not d['checks'].passed.all():
        sys.exit('A reconciliation check failed; outputs not written.')
    export_xlsx.export(d, os.path.join(ROOT, 'data', 'INDICATE_WLC_database.xlsx'), os.path.join(ROOT, 'data', 'csv'))
    export_web.export(d, os.path.join(ROOT, 'assets', 'js', 'data.js'))
    req = export_template.export(d, TEMPLATE, os.path.join(ROOT, 'data', 'data-requests'))
    p = d['projects']
    print(f"\nWrote {len(p)} buildings ({', '.join(f'{c} {n}' for c, n in p.groupby('country').size().items())}), "
          f"{len(d['issues'])} issues ({(d['issues'].type == 'confirm').sum()} to confirm).")
    for path, n, ask, chk in req:
        print(f'  data request {os.path.basename(path)}: {n} buildings, {ask} empty high-priority cells, {chk} assumptions to check')


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'source'))
