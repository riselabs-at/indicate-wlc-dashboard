"""INDICATE WLC database – configuration and decisions.

Schema follows the INDICATE Building WLC Emissions Reporting Template v1.1 (attribute codes, module columns,
typology codes, element tiers). Fields the template does not have carry the prefix x_ (extension).
Every decision that changes data lives in this file and is written to the database's mapping sheets.
"""

DATABASE_VERSION = '0.3.1'
TEMPLATE_VERSION = 'INDICATE Building WLC Emissions Reporting Template v1.1 (draft 261020), based on v1.0 (Röck 2026, doi:10.5281/zenodo.21169904)'

SOURCES = {
    'AT': dict(pattern='*IND-AT*WLC*.xlsx', deliverable='D7.3', org='TU Graz', country='Austria', iso3='AUT',
               title='D7.3 AT national WLC datasets following documentation requirements'),
    'HR': dict(pattern='*HR_WLC*.xlsx', deliverable='D7.12', org='University of Zagreb, Faculty of Civil Engineering (UZ-FCE)', country='Croatia', iso3='HRV',
               title='D7.12 HR national WLC datasets following documentation requirements'),
    'IT': dict(pattern='*D7.21*.xlsx', deliverable='D7.21', org='Politecnico di Milano (POLIMI), with POLITO, UNIBG and Harpaceas', country='Italy', iso3='ITA',
               title='D7.21 IT national WLC datasets following documentation requirements'),
    'LU': dict(pattern='*LU_Export*.xlsx', deliverable='D7.30', org='Energie et Environnement S.A. (on behalf of the OAI)', country='Luxembourg', iso3='LUX',
               title='D7.30 LU national WLC datasets following documentation requirements'),
}

# ------------------------------------------------------------------ template vocabulary
# Native life-cycle modules of the template v1.1 (B5 added in v1.1)
MODULES = ['A1-A3', 'A4', 'A5', 'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'C1', 'C2', 'C3', 'C4', 'D1', 'D2']
INDICATORS = ['GWPtotal', 'GWPfossil', 'GWPbio', 'GWPluluc']

# Display modules used by the dashboard and the 'results' sheet (≤ 8 + D, so every chart colour stays distinct)
DISPLAY = {'A1-A3': ['A1-A3'], 'A4': ['A4'], 'A5': ['A5'], 'B1-B3': ['B1', 'B2', 'B3'], 'B4-B5': ['B4', 'B5'],
           'B6': ['B6'], 'B7': ['B7'], 'C1-C4': ['C1', 'C2', 'C3', 'C4'], 'D': ['D1', 'D2']}
DISPLAY_ORDER = list(DISPLAY)
EMBODIED = ['A1-A3', 'A4', 'A5', 'B1-B3', 'B4-B5', 'C1-C4']   # all non-operational modules, as in D3.2 (B1-B3 only where reported)
WHOLE_LIFE = EMBODIED + ['B6', 'B7']                            # D always separate
RSP_YEARS = 50

# Template typology list (Options tab) – code -> label
TYPOLOGIES = {'SFH': 'Single family houses (SFH)', 'MFH': 'Multifamily houses (MFH)', 'ABL': 'Apartment blocks (ABL)',
              'OFF': 'Offices (OFF)', 'TRA': 'Trade (TRA)', 'EDU': 'Education (EDU)', 'HEA': 'Health (HEA)',
              'HOR': 'Hotels and Restaurants (HOR)', 'OTH': 'Other non-residential buildings (OTH)'}

# country, national typology (as in source), template code, status, note
TYPOLOGY_MAP = [
    ('AT', 'Single-family house', 'SFH', 'direct', ''),
    ('AT', 'Multi-family house', 'MFH', 'direct', ''),
    ('AT', 'Apartment block', 'ABL', 'direct', ''),
    ('AT', 'Office building', 'OFF', 'direct', ''),
    ('AT', 'Education building', 'EDU', 'direct', ''),
    ('AT', 'Retail', 'TRA', 'direct', ''),
    ('AT', 'Hotel and Restaurant', 'HOR', 'direct', ''),
    ('HR', 'SINGLE FAMILY HOUSE', 'SFH', 'direct', ''),
    ('HR', 'MULTY FAMILY HOUSE', 'MFH', 'direct', ''),
    ('HR', 'OFFICE BUILDING', 'OFF', 'direct', ''),
    ('HR', 'EDUCATION', 'EDU', 'direct', ''),
    ('HR', 'HOSPITAL', 'HEA', 'direct', ''),
    ('HR', 'HOTEL', 'HOR', 'direct', ''),
    ('HR', 'SPORT HALL BUILDING', 'OTH', 'direct', 'No sport type in the template list'),
    ('HR', 'NEW BUILDING FOR RETAIL AND WHOLESALE', 'TRA', 'direct', ''),
    ('IT', 'SFH', 'SFH', 'direct', ''),
    ('IT', 'MFH', 'MFH', 'direct', ''),
    ('IT', 'OFF', 'OFF', 'direct', ''),
    ('IT', 'EDU', 'EDU', 'direct', ''),
    ('IT', 'HOS', 'HOR', 'to confirm', 'Code not defined in the file; hospitality (mainly long-term stays and student rooms) per D3.2, which groups it with MFH; HOR agreed 2026-11-08, to be confirmed by the Italian team'),
    ('IT', 'XTR', 'OTH', 'to confirm', 'Code not defined in the file (non-residential); OTH until confirmed'),
    ('LU', 'House', 'SFH', 'to confirm', 'Includes some larger multi-unit residences (e.g. MI_Con_1, 35 occupants)'),
    ('LU', 'Apartments', 'ABL', 'direct', 'Apartment buildings → ABL (agreed 2026-10-20)'),
    ('LU', 'Office', 'OFF', 'direct', ''),
    ('LU', 'School', 'EDU', 'direct', ''),
    ('LU', 'Hall', 'OTH', 'direct', 'No hall type in the template list'),
]
TYPE_OF = {(c, n): code for c, n, code, _, _ in TYPOLOGY_MAP}

DEPTH = {'-': 'Not applicable (new construction)', 'Minor': 'Minor renovation', 'Major': 'Major renovation'}

# ------------------------------------------------------------------ EU Table 4 element tiers (INDICATE numbering, as in template v1.1)
EU_ELEMENTS = [  # code, tier, name
    ('1', 1, 'Shell'), ('1.1', 2, 'Substructure'), ('1.2', 2, 'Structure'), ('1.3', 2, 'External architectural works (non-structural)'),
    ('1.3.1', 3, 'Façade'), ('1.3.2', 3, 'Roof'),
    ('2', 1, 'Core'), ('2.1', 2, 'Internal or under cover architectural works (non-structural)'), ('2.2', 2, 'Building services and equipment'),
    ('2.2.1', 3, 'Water and wastewater-related systems'), ('2.2.2', 3, 'Heating systems'), ('2.2.3', 3, 'Dedicated cooling systems'),
    ('2.2.4', 3, 'Ventilation systems'), ('2.2.5', 3, 'Lighting systems'),
    ('2.2.6', 3, 'Electrical services for power, communications, security, IT and fire detection'),
    ('2.2.7', 3, 'On-site renewable energy generation'), ('2.2.8', 3, 'Life safety, fuel and movement system installations'),
    ('2.2.9', 3, 'Waste disposal systems'),
    ('3', 1, 'External works'), ('3.1', 2, 'External roads, paths, paving and other surfaces'), ('3.2', 2, 'External fittings and fixtures'),
    ('3.3', 2, 'External building services'), ('3.4', 2, 'External buildings'),
    ('4', 1, 'Other elements (not in EU Table 4)'), ('4.1', 2, 'Other or not allocated'),
]
EU_NAME = {c: n for c, _, n in EU_ELEMENTS}
EU_TIER = {c: t for c, t, _ in EU_ELEMENTS}

# national element code -> EU code (deepest level the evidence supports), status, basis
ELEMENT_MAP = [
    # Luxembourg "LEVELS Tier 3" codes; meaning inferred from the product categories each code contains
    ('LU', '1.1.1', '1.1', 'inferred', 'concrete, reinforcement, drainage → substructure'),
    ('LU', '1.1.2', '1.1', 'inferred', 'concrete, reinforcement, insulation, aggregates → substructure'),
    ('LU', '1.1.3', '1.1', 'inferred', 'concrete, insulation, waterproofing → substructure / basement'),
    ('LU', '1.1.4', '1.1', 'inferred', 'concrete, reinforcement, masonry → substructure'),
    ('LU', '1.2.1', '1.2', 'inferred', 'structural concrete, steel, wood → structure'),
    ('LU', '1.2.2', '1.2', 'inferred', 'concrete, stairs, tanks → structure'),
    ('LU', '1.3.1', '1.3.1', 'inferred', 'shutters, glazing, joinery, insulation → façade'),
    ('LU', '1.3.2', '1.3.2', 'inferred', 'insulation, metal, EPDM / bituminous membranes → roof'),
    ('LU', '2.1.1', '2.1', 'inferred', 'internal doors, partitions → internal architectural works'),
    ('LU', '2.1.2', '2.1', 'inferred', 'stairs, guardrails, furniture → internal architectural works'),
    ('LU', '2.1.3', '2.1', 'inferred', 'paint, screeds, floor coverings → internal finishes'),
    ('LU', '2.1.4', '2.1', 'inferred', 'paint, floor coverings, gypsum → internal finishes'),
    ('LU', '2.2.1', '2.2.2', 'inferred', 'valves, hydronic emitters, pipes → heating distribution'),
    ('LU', '2.2.2', '2.2.2', 'inferred', 'hydronic heat / cooling production → heating (may include cooling)'),
    ('LU', '2.2.3', '2.2.4', 'inferred', 'ducts, air handling units, grilles → ventilation'),
    ('LU', '2.2.4', '2.2.7', 'inferred', 'solar panels, conversion, generators → on-site renewables'),
    ('LU', '2.2.5', '2.2.6', 'inferred', 'LV wiring, cable trays, cabinets → electrical services'),
    ('LU', '2.2.6', '2.2.5', 'inferred', 'interior / emergency lighting → lighting'),
    ('LU', '2.2.7', '2.2.5', 'inferred', 'interior / exterior lighting → lighting'),
    ('LU', '2.2.8', '2.2.6', 'inferred', 'cameras, access control, intercoms → electrical services (security, comms)'),
    ('LU', '2.3.1', '2.2.1', 'inferred', 'pipes, network accessories → water'),
    ('LU', '2.3.2', '2.2.1', 'inferred', 'pipes, water treatment, chambers → water and wastewater'),
    ('LU', '2.4.1', '2.2.1', 'inferred', 'WCs, sinks, washbasins → sanitaryware (water systems)'),
    ('LU', '2.4.2', '2.2.1', 'inferred', 'washbasins, baths, taps, showers → sanitaryware (water systems)'),
    ('LU', '2.5.1', '2.2.8', 'inferred', 'fire safety, alarm systems → life safety'),
    ('LU', '2.5.2', '2.2.8', 'inferred', 'pipes (single product) → life safety'),
    ('LU', '2.5.3', '2.2.8', 'inferred', 'lifts → movement installations'),
    ('LU', '3.1.1', '3.3', 'inferred', 'drainage, pipes → external building services'),
    ('LU', '3.1.2', '3.3', 'inferred', 'pipes, chambers, drainage → external building services'),
    ('LU', '3.1.3', '3.3', 'inferred', 'chambers, pipes, drainage → external building services'),
    ('LU', '3.2.1', '3.2', 'inferred', 'concrete, aggregates, fencing → external fittings'),
    ('LU', '3.2.2', '3.1', 'inferred', 'concrete, road surfacing → external roads and surfaces'),
    ('LU', '3.2.3', '3.1', 'inferred', 'roadway elements, green-roof substrate → external surfaces (mixed)'),
    ('LU', '3.2.4', '3.3', 'inferred', 'pipes, concrete, solar panels → external services (mixed)'),
    # Italian tiers; reading agreed 2026-10-20 as the EU Tier 3 services list, to be confirmed by the Italian team
    ('IT', '1.1', '1.1', 'inferred', 'same code as EU Tier 2'),
    ('IT', '1.2', '1.2', 'inferred', 'same code as EU Tier 2'),
    ('IT', '1.3.1', '1.3.1', 'inferred', 'same code as EU Tier 3 façade'),
    ('IT', '1.3.2', '1.3.2', 'inferred', 'same code as EU Tier 3 roof'),
    ('IT', '2.1', '2.1', 'inferred', 'same code as EU Tier 2'),
    ('IT', '2.2', '2.2.1', 'inferred', '1st EU services item: water and wastewater'),
    ('IT', '2.3', '2.2.2', 'inferred', '2nd and 3rd EU services items: heating, incl. cooling (source: "2.3 includes 2.4")'),
    ('IT', '2.5', '2.2.4', 'inferred', '4th EU services item: ventilation'),
    ('IT', '2.6', '2.2.5', 'inferred', '5th EU services item: lighting'),
    ('IT', '2.7', '2.2.6', 'inferred', '6th EU services item: electrical services'),
    ('IT', '2.8', '2.2.7', 'inferred', '7th EU services item: on-site renewables'),
    ('IT', '2.9', '2.2.8', 'inferred', '8th EU services item: life safety, fuel, movement'),
    ('IT', '4.1', '4.1', 'unknown', 'renovation cases only; meaning not stated'),
]
EL_MAP = {(c, n): (e, s, b) for c, n, e, s, b in ELEMENT_MAP}

# Italian element-tier cell colours (sheet 'SumUP Tiers') -> data-quality flag, from the sheet's legend
IT_FILL_QUALITY = {
    'FFC0C5E1': 'original',
    ('theme', 6, 0.8): 'proxy: missing value', 'FFC1F0C8': 'proxy: missing value', 'FFDAF2D0': 'proxy: missing value',
    ('theme', 9, 0.0): 'proxy: replaces outlier', 'FF4EA72E': 'proxy: replaces outlier',
    'FFFFC000': 'proxy: from MFH-HOS typology',
    ('theme', 8, 0.6): 'proxy: from EDU typology',
}
IT_FONT_BLUE = 'FF0070C0'          # legend: +16.67 kgCO2e/m² GFA proxy for internal paint (tier 2.1)
IT_FONT_GREY = ('theme', 0, -0.5)  # grey: insufficient data for D7.22 benchmarks

# ------------------------------------------------------------------ template attribute codebook (Project-information + overview)
# code, label, unit, priority, kind ('template' / 'extension')
CODEBOOK = [
    ('building_id', 'Database building ID (<country>-<source ID>)', '-', '', 'extension'),
    ('contact_name', 'Contact name', '-', '', 'template'), ('contact_email', 'Contact email', '-', 'HP', 'template'),
    ('contact_org', 'Contact organization', '-', '', 'template'),
    ('project_name', 'Project name', '-', '', 'template'), ('project_code', 'Project code', '-', 'HP', 'template'),
    ('project_intervention_type', 'Intervention type', '-', 'HP', 'template'), ('project_renovation_depth', 'Renovation depth', '-', 'HP', 'template'),
    ('site_country', 'Country name', '-', 'HP', 'template'), ('site_country_iso3', 'Country code (ISO 3166-1 alpha-3)', '-', 'HP', 'template'),
    ('site_region_local', 'City / region', '-', '', 'template'),
    ('bldg_use_subtype', 'Building typology', '-', 'HP', 'template'), ('bldg_year_complete', 'Year of completion', '-', '', 'template'),
    ('bldg_users_total', 'Number of users', 'users', 'HP', 'template'),
    ('bldg_area_gfa', 'Gross Floor Area', 'm²', 'HP', 'template'), ('bldg_area_nfa', 'Net Floor Area', 'm²', 'HP', 'template'),
    ('bldg_floors_ag', 'Floors above ground', 'floors', 'HP', 'template'), ('bldg_floors_bg', 'Floors below ground', 'floors', 'HP', 'template'),
    ('bldg_area_references', 'References for floor area definition', '-', '', 'template'),
    ('bldg_height', 'Building height', 'm', 'HP', 'template'), ('bldg_footprint', 'Building footprint', 'm²', '', 'template'),
    ('bldg_energy_class_country', 'Energy performance class', '-', '', 'template'),
    ('inv_energy_consumption', 'Energy consumption per year', 'kWh/a', 'HP', 'template'),
    ('inv_energy_consumption_description', 'Scope of energy consumption values', '-', '', 'template'),
    ('inv_energy_consumption_type', 'Type of energy source(s)', '-', '', 'template'),
    ('bldg_struct_type', 'Structure type', '-', 'HP', 'template'), ('inv_mat_mass_total', 'Total mass', 'kg', 'HP', 'template'),
    ('lca_rsp', 'Reference study period (RSP)', 'years', 'HP', 'template'), ('lca_software', 'LCA software', '-', '', 'template'),
    ('lca_database', 'LCA database name', '-', '', 'template'), ('lca_database_version', 'LCA database version', '-', '', 'template'),
    ('lca_lcia_method', 'LCIA method (CO2/GHG emissions)', '-', '', 'template'),
    ('x_typology_national', 'National typology as reported', '-', '', 'extension'),
    ('x_typology_status', 'Typology mapping status', '-', '', 'extension'),
    ('x_nfa_basis', 'Area used as NFA (UFA or HFA)', '-', '', 'extension'),
    ('x_rsp_basis', 'How the RSP was obtained', '-', '', 'extension'),
    ('x_intervention_basis', 'How the intervention type was obtained', '-', '', 'extension'),
    ('x_results_basis', 'Basis of the reported results (absolute or per m²)', '-', '', 'extension'),
    ('x_storeys_total', 'Storeys, total (AT; not split above / below ground)', 'floors', '', 'extension'),
    ('x_struct_material_national', 'Main structural material as reported (LU)', '-', '', 'extension'),
    ('x_energy_source_national', 'Heating energy source as reported (LU)', '-', '', 'extension'),
    ('x_primary_energy_kwh_m2a', 'Primary energy demand as reported (LU, kWh/m²·a, per heated area An)', 'kWh/m²a', '', 'extension'),
    ('x_heated_area_an_m2', 'Heated area An (LU)', 'm²', '', 'extension'),
    ('x_gfa_below_m2', 'GFA below ground (LU)', 'm²', '', 'extension'), ('x_gfa_above_m2', 'GFA above ground (LU)', 'm²', '', 'extension'),
    ('x_compactness', 'Compactness (LU)', '-', '', 'extension'), ('x_window_wall_ratio', 'Window-to-wall ratio (LU)', '-', '', 'extension'),
    ('x_users_basis', 'How the number of users was obtained', '-', '', 'extension'),
    ('x_source_deliverable', 'Source deliverable', '-', '', 'extension'), ('x_source_file', 'Source file', '-', '', 'extension'),
    ('x_source_id', 'ID in the source file', '-', '', 'extension'), ('x_has_elements', 'Element-level results available', '-', '', 'extension'),
]
PROJECT_COLUMNS = [c for c, *_ in CODEBOOK]
