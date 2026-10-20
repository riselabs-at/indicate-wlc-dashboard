"""Shared configuration for the INDICATE WLC harmonisation pipeline.

Everything that is a *decision* (module mapping, typology mapping, assumptions) lives here,
so it can be reviewed in one place and is written into the database's mapping sheets.
"""

DATABASE_VERSION = '0.2'

# Source workbooks (glob patterns inside the source folder)
SOURCES = {
    'AT': dict(pattern='*IND-AT*WLC*.xlsx', deliverable='D7.3', lead='TU Graz',
               title='D7.3 AT national WLC datasets following documentation requirements'),
    'HR': dict(pattern='*HR_WLC*.xlsx', deliverable='D7.12', lead='UZ-FCE',
               title='D7.12 HR national WLC datasets following documentation requirements'),
    'LU': dict(pattern='*LU_Export*.xlsx', deliverable='D7.30',
               lead='Energie et Environnement S.A. (on behalf of the OAI)',
               title='D7.30 LU national WLC datasets following documentation requirements'),
    'IT': dict(pattern='*D7.21*.xlsx', deliverable='D7.21', lead='POLIMI (with POLITO, UNIBG, Harpaceas)',
               title='D7.21 IT national WLC datasets following documentation requirements'),
}

COUNTRY_NAMES = {'AT': 'Austria', 'HR': 'Croatia', 'LU': 'Luxembourg', 'IT': 'Italy'}

# Harmonised life-cycle modules, in stacking order.
# B2 (maintenance) and B3 (repair) are combined so the set stays within 8 chart colours.
MODULES = ['A1-A3', 'A4', 'A5', 'B2-B3', 'B4-B5', 'B6', 'B7', 'C1-C4', 'D']
EMBODIED = ['A1-A3', 'A4', 'A5', 'B4-B5', 'C1-C4']          # reported by all four datasets
WHOLE_LIFE = EMBODIED + ['B2-B3', 'B6', 'B7']                # D always kept separate
MODULE_LABELS = {
    'A1-A3': 'A1–A3 Product', 'A4': 'A4 Transport', 'A5': 'A5 Construction', 'B2-B3': 'B2–B3 Maintenance & repair',
    'B4-B5': 'B4–B5 Replacement', 'B6': 'B6 Operational energy', 'B7': 'B7 Operational water',
    'C1-C4': 'C1–C4 End of life', 'D': 'D Beyond boundary'}

# Source column -> harmonised module, per country. Written to the 'map_modules' sheet.
MODULE_MAP = [
    # country, source sheet, source column(s), harmonised module, note
    ('AT', 'Data overview', 'GWP total - A1A3', 'A1-A3', ''),
    ('AT', 'Data overview', 'GWP total - A4', 'A4', ''),
    ('AT', 'Data overview', 'GWP total - A5', 'A5', ''),
    ('AT', 'Data overview', 'GWP total - B4', 'B4-B5', 'B4 only; B5 not reported'),
    ('AT', 'Data overview', 'GWP total - B5', 'B6', 'Header says B5, description row says B6 (operational energy). Treated as B6. TO CONFIRM'),
    ('AT', 'Data overview', 'GWP total - C1 + C2 + C3C4', 'C1-C4', 'Summed'),
    ('HR', 'Data_overview', 'A1_A3_GWP_TOTAL (absolute)', 'A1-A3', ''),
    ('HR', 'Data_overview', 'A4_GWP_TOTAL', 'A4', ''),
    ('HR', 'Data_overview', 'A5_GWP_total', 'A5', ''),
    ('HR', 'Data_overview', 'B3_GWP total', 'B2-B3', 'B3 only'),
    ('HR', 'Data_overview', 'B4-B5_GWP_TOTAL', 'B4-B5', ''),
    ('HR', 'Data_overview', 'B6_GWP_TOTAL', 'B6', ''),
    ('HR', 'Data_overview', 'B7_GWP_total', 'B7', ''),
    ('HR', 'Data_overview', 'C1-C4_GWP total', 'C1-C4', ''),
    ('HR', 'Data_overview', 'D_GWP_total', 'D', ''),
    ('LU', 'Overview of results', 'A1-A3, fossil + A1-A3, biog', 'A1-A3', 'Per m² (area not stated, assumed GFA); summed over all products'),
    ('LU', 'Overview of results', 'A4, fossil', 'A4', 'As above'),
    ('LU', 'Overview of results', 'A5, fossil', 'A5', 'As above'),
    ('LU', 'Overview of results', 'B4-B5, fossil', 'B4-B5', 'As above'),
    ('LU', 'Overview of results', 'C2 + C3 + C4, fossil + C3-C4, biog', 'C1-C4', 'C1 not reported'),
    ('LU', 'Overview of results', 'D, tot', 'D', ''),
    ('IT', 'SumUP LCM', 'A1_A3_GWP_TOTAL (kgCO2e)', 'A1-A3', 'Absolute block (cols AT–BF)'),
    ('IT', 'SumUP LCM', 'A4_GWP_TOTAL', 'A4', ''),
    ('IT', 'SumUP LCM', 'A5_GWP_Fossil + A5w_DAB + A5w_BOQ', 'A5', 'Equals A5_GWP_Tot per m² × GFA (checked)'),
    ('IT', 'SumUP LCM', 'B2_GWP_Fossil + B3_GWP_Fossil', 'B2-B3', 'GWP-fossil only (not GWP-total) in source'),
    ('IT', 'SumUP LCM', 'B4_GWP_TOTAL', 'B4-B5', 'B4 only; B5 not reported'),
    ('IT', 'SumUP LCM', 'B6_GWP_TOTAL', 'B6', 'B6 = 0 treated as not reported'),
    ('IT', 'SumUP LCM', 'C1_GWP_Fossil + C2 + C3 + C4 _GWP_TOTAL', 'C1-C4', 'C1 is GWP-fossil only'),
]

HARMONISED_TYPES = ['Single-family house', 'Multi-family / apartments', 'Office', 'Education', 'Retail', 'Hotel',
                    'Hospital', 'Hall / sports', 'Other residential (IT: HOS)', 'Other non-residential (IT: XTR)']

# country, national typology (as in source), harmonised type, note
TYPOLOGY_MAP = [
    ('AT', 'Single-family house', 'Single-family house', ''),
    ('AT', 'Apartment block', 'Multi-family / apartments', ''),
    ('AT', 'Multi-family house', 'Multi-family / apartments', ''),
    ('AT', 'Office building', 'Office', ''),
    ('AT', 'Education building', 'Education', ''),
    ('AT', 'Retail', 'Retail', ''),
    ('AT', 'Hotel and Restaurant', 'Hotel', ''),
    ('HR', 'SINGLE FAMILY HOUSE', 'Single-family house', ''),
    ('HR', 'MULTY FAMILY HOUSE', 'Multi-family / apartments', ''),
    ('HR', 'OFFICE BUILDING', 'Office', ''),
    ('HR', 'EDUCATION', 'Education', ''),
    ('HR', 'HOSPITAL', 'Hospital', ''),
    ('HR', 'HOTEL', 'Hotel', ''),
    ('HR', 'SPORT HALL BUILDING', 'Hall / sports', ''),
    ('HR', 'NEW BUILDING FOR RETAIL AND WHOLESALE', 'Retail', ''),
    ('LU', 'House', 'Single-family house', 'Includes some larger multi-unit residences (e.g. MI_Con_1). TO CONFIRM'),
    ('LU', 'Apartments', 'Multi-family / apartments', ''),
    ('LU', 'Office', 'Office', ''),
    ('LU', 'School', 'Education', ''),
    ('LU', 'Hall', 'Hall / sports', ''),
    ('IT', 'SFH', 'Single-family house', ''),
    ('IT', 'MFH', 'Multi-family / apartments', ''),
    ('IT', 'OFF', 'Office', ''),
    ('IT', 'EDU', 'Education', ''),
    ('IT', 'HOS', 'Other residential (IT: HOS)', 'Code not defined in the file; residential (Typ_code R), grouped with MFH as "MFH-HOS". Meaning TO CONFIRM'),
    ('IT', 'XTR', 'Other non-residential (IT: XTR)', 'Code not defined in the file; non-residential (Typ_code NR). Meaning TO CONFIRM'),
]
TYPE_OF = {(c, n): h for c, n, h, _ in TYPOLOGY_MAP}

# Italian element-tier cell colours (sheet 'SumUP Tiers') -> data-quality flag, from the sheet's legend.
# GFA and HFA halves use different shades of the same theme colours; all shades are listed.
IT_FILL_QUALITY = {
    'FFC0C5E1': 'original',
    ('theme', 6, 0.8): 'proxy: missing value', 'FFC1F0C8': 'proxy: missing value', 'FFDAF2D0': 'proxy: missing value',
    ('theme', 9, 0.0): 'proxy: replaces outlier', 'FF4EA72E': 'proxy: replaces outlier',
    'FFFFC000': 'proxy: from MFH-HOS typology',
    ('theme', 8, 0.6): 'proxy: from EDU typology',
}
IT_FONT_BLUE = 'FF0070C0'          # legend: proxy 16.67 kgCO2e/m² GFA added for internal paint (tier 2.1)
IT_FONT_GREY = ('theme', 0, -0.5)  # note: grey values had insufficient data for benchmark generation

RSP_YEARS = 50
