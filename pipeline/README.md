# Harmonisation pipeline

This folder turns the four national WLC deliverable workbooks into one harmonised database (`data/INDICATE_WLC_database.xlsx`, `data/csv/`) and the dashboard data file (`assets/js/data.js`).

```bash
pip install -r pipeline/requirements.txt
python pipeline/run_all.py path/to/source-folder
```

The source folder must hold the four workbooks. They are not stored in the repository; put them in a local `source/` folder, which git ignores.

| Pattern | Deliverable | Country |
|---|---|---|
| `*IND-AT*WLC*.xlsx` | D7.3 | Austria |
| `*HR_WLC*.xlsx` | D7.12 | Croatia |
| `*D7.21*.xlsx` | D7.21 | Italy |
| `*LU_Export*.xlsx` | D7.30 | Luxembourg |

The run prints a processing log and the reconciliation checks. If any check fails, it stops without writing outputs.

## Files

| File | Role |
|---|---|
| `config.py` | Every harmonisation decision in one place: module list and scopes, source-column → module mapping, national typology → harmonised type mapping, the Italian cell-colour legend. |
| `common.py` | Number parsing, source-file lookup, and the `Extract` container each country fills (buildings, module results, elements, materials, log, issues, checks). |
| `extract_at.py` | Austria: sheet *Data overview*. Absolute kgCO2e per module; UFA only. |
| `extract_hr.py` | Croatia: sheet *Data_overview*, absolute block. GFA and UFA. |
| `extract_it.py` | Italy: sheet *SumUP LCM* (absolute block) and sheet *SumUP Tiers* (tier values plus a quality flag from each cell's fill and font colour). GFA and HFA. |
| `extract_lu.py` | Luxembourg: sheet *Overview of results* (11,133 product rows, per m²) summed per project, per LEVELS Tier 3 code and per material group; sheet *A. Metadata* for areas. |
| `harmonise.py` | Combines the extracts, derives intervention class and UFA-equivalent area, builds the long and wide tables, and runs the cross-checks. |
| `export_xlsx.py` | Writes the Excel database (results and statistics as live formulas) and CSV copies. |
| `export_web.py` | Writes `assets/js/data.js` for the dashboard. |
| `run_all.py` | Runs everything. |

## Processing steps

1. **Read.** Each workbook is read by its own extractor. Only result sheets are used.
2. **Clean.** Blank rows, repeated header rows (e.g. the French "Projet" rows in LU) and non-building rows are dropped. Numbers are parsed defensively: decimal commas are accepted, and text becomes blank.
3. **Identify.** Each building gets the ID `<country>-<source ID>`. The national typology, intervention type and floor areas are kept as reported.
4. **Map modules** to A1-A3, A4, A5, B2-B3, B4-B5, B6, B7, C1-C4 and D (see `MODULE_MAP`). The results are stored as absolute kgCO2e per building over a 50-year reference study period.
   - **AT:** "B5" is treated as B6, following the source's description row. B4-B5 contains B4 only.
   - **HR:** the modules map 1:1, with B3 going to B2-B3.
   - **IT:** A5 = A5 + A5w_DAB + A5w_BOQ. B2 + B3 go to B2-B3. B6 = 0 is treated as not reported.
   - **LU:** the fossil and biogenic columns are summed, and C1 is not reported.
5. **Convert units.** AT, HR and IT report absolute values. LU reports per m² without naming the area; absolute = value × GFA. This is an assumption, and the dashboard can switch it to UFA.
6. **Map typologies** (`TYPOLOGY_MAP`). The Italian codes HOS and XTR have no definition in the file, so they stay as their own types.
7. **Normalise.** Per m² = absolute ÷ area. The area is either GFA or the UFA-equivalent: UFA for AT, HR and LU, and heated floor area (HFA) for IT, which reports no UFA. Per year = ÷ 50.
8. **Elements.** These stay in each dataset's own classification and are not merged across countries. LU uses its LEVELS Tier 3 codes (products summed). IT uses the tiers from *SumUP Tiers*, each with a quality flag: original, proxy for a missing value, proxy replacing an outlier, proxy taken from the MFH-HOS or EDU typology, or +16.67 kgCO2e/m² paint.
9. **Check.** Each dataset is reconciled against its own source:
   - AT: GWP total = fossil + biogenic + LULUC.
   - HR: the module sum matches the source total (within 2 kgCO2e).
   - IT: absolute = per-m² × GFA, the module sum = GWPtot, and RSP = 50.
   - LU: the element rows sum to the building totals.
10. **Flag.** Assumptions and data-quality observations are logged in the *issues* sheet. Anything marked "open – partner to confirm" needs a decision.
11. **Export.** The pipeline writes the Excel database, the CSVs and `data.js`.

## Adding a country

Write `extract_<cc>.py` returning an `Extract` (copy the closest existing extractor), add the source pattern to `SOURCES` and the mappings to `MODULE_MAP` / `TYPOLOGY_MAP` in `config.py`, add it to `EXTRACTORS` in `harmonise.py`, and add the country button and coverage column in `index.html`.
