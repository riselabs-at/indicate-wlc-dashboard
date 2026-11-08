# Harmonisation pipeline

This folder turns the four national WLC deliverable workbooks into one database in the vocabulary of the **INDICATE reporting template v1.1**. It also writes the dashboard data and one pre-filled template per country for the data request.

```bash
pip install -r pipeline/requirements.txt
python pipeline/run_all.py path/to/source-folder
```

| Source file pattern | Deliverable | Country |
|---|---|---|
| `*IND-AT*WLC*.xlsx` | D7.3 | Austria |
| `*HR_WLC*.xlsx` | D7.12 | Croatia |
| `*D7.21*.xlsx` | D7.21 | Italy |
| `*LU_Export*.xlsx` | D7.30 | Luxembourg |

## Files

| File | Role |
|---|---|
| `config.py` | All decisions in one place: template modules and display groups, the typology map to the 9 template codes, the EU Table 4 element tiers and the national-code → EU-tier map, the Italian cell-colour legend, and the codebook of template attributes and `x_` extensions. |
| `common.py` | Number parsing, file lookup, and the `Extract` container every country fills. |
| `extract_at.py` | Austria, sheet *Data overview*: GWP total / fossil / biogenic / LULUC per module; UFA; RSP; LCIA method. |
| `extract_hr.py` | Croatia, sheet *Data_overview*: absolute block; GFA, UFA; combined modules flagged. |
| `extract_it.py` | Italy, sheets *SumUP LCM* (absolute block, GFA, HFA, intervention, renovation level) and *SumUP Tiers* (tier values; each value gets a quality flag from its cell colour). |
| `extract_lu.py` | Luxembourg, sheets *Overview of results* (product rows per m²) and *A. Metadata* (areas, floors, users, year, structure, energy). |
| `harmonise.py` | Combines the extracts, maps elements to EU tiers, builds display modules, computes the codebook completeness and the mapping tables, and runs the cross-checks. |
| `export_xlsx.py` | Writes the database workbook (results and statistics as live formulas) and the CSVs. |
| `export_web.py` | Writes `assets/js/data.js`, including completeness per attribute and module. |
| `export_template.py` | Writes `data/data-requests/`: template v1.1 pre-filled per country, with empty high-priority fields in pink, assumptions in yellow with comments, a *Questions* tab and an *Element mapping* tab. |
| `run_all.py` | Runs everything. |

## Processing steps

1. **Read.** Each workbook is read by its own extractor.
2. **Clean.** Blank and repeated header rows are dropped. Numbers are parsed defensively.
3. **Describe buildings in template vocabulary.** Each building becomes one row in `projects`, using the template attribute codes. Information the template has no field for goes into `x_` extension fields.
4. **Map typologies** to SFH, MFH, ABL, OFF, TRA, EDU, HEA, HOR and OTH (multi-family houses → MFH, apartment buildings → ABL; Italian HOS → HOR (hospitality, per D3.2) and XTR → OTH, both to be confirmed).
5. **Map results to the 16 template modules.** Each building × indicator (GWPtotal / fossil / bio / LULUC) × module is stored once. A value that spans several modules is stored in the first module of its range, with the range in `covers` (template rule).
6. **Convert units.** All results are absolute kgCO2e over 50 years. Luxembourg reports per m² with no stated area; D3.2 (Table E.1) reproduces the values exactly as per m² UFA, so absolute = per-m² value × UFA.
7. **Areas.** NFA = UFA for AT, HR and LU, and HFA for IT. `x_nfa_basis` records which.
8. **Elements.** National element codes are mapped to the EU Table 4 tiers. Each mapping has a status (`inferred` until confirmed) and a basis; Italian values keep their source-quality flag.
9. **Display modules** for the charts: A1-A3, A4, A5, B1-B3, B4-B5, B6, B7, C1-C4, D. Embodied = A1-A3 + A4 + A5 + B1-B3 + B4-B5 + C1-C4 as reported (the D3.2 definition).
10. **Check.** Reconciliation against each source:
    - module sums vs source totals;
    - absolute vs per-m² values;
    - sub-indicators vs total;
    - element sums vs building totals.

    The run stops if any check fails.
11. **Flag and ask.** Every assumption and data-quality observation goes into `issues`, with the question for the national team.
12. **Export** the database, CSVs, dashboard data and data-request files.

## Adding a country or returned data

Write or extend `extract_<cc>.py` so it fills the template attribute codes and `x.result(...)` rows. Add its patterns and maps to `config.py` and register the extractor in `harmonise.EXTRACTORS`. Completed data-request files can be read the same way: they use the template layout.
