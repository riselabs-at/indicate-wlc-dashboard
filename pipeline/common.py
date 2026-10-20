"""Helpers shared by the country extractors."""
import glob, math, os
from dataclasses import dataclass, field


def num(x):
    """Parse a number from a cell; returns None for blanks/text. Accepts decimal commas."""
    if x is None:
        return None
    try:
        v = float(str(x).replace(',', '.'))
    except (TypeError, ValueError):
        return None
    return None if math.isnan(v) else v


def find_source(folder, pattern):
    hits = sorted(glob.glob(os.path.join(folder, pattern)))
    if not hits:
        raise FileNotFoundError(f'No file matching {pattern!r} in {folder}')
    if len(hits) > 1:
        raise RuntimeError(f'Several files match {pattern!r}: {hits}')
    return hits[0]


@dataclass
class Extract:
    """Everything one country contributes to the database."""
    country: str
    source_file: str
    buildings: list = field(default_factory=list)   # dicts, see harmonise.BUILDING_COLUMNS
    lcm: list = field(default_factory=list)         # dicts: building_id, module, source_value, source_unit, value_kgco2e, source_columns
    elements: list = field(default_factory=list)    # dicts: building_id, classification, code, group, module, value_kgco2e_m2, area_basis, quality
    materials: list = field(default_factory=list)   # dicts: building_id, material_group, module, value_kgco2e_m2
    element_desc: dict = field(default_factory=dict)  # code -> description derived from the data
    log: list = field(default_factory=list)         # (step, detail, count)
    issues: list = field(default_factory=list)      # (severity, building_id or '', description, handling)
    checks: list = field(default_factory=list)      # (check, value, tolerance, passed)

    def step(self, step, detail, count=''):
        self.log.append((self.country, step, detail, count))

    def issue(self, severity, bid, description, handling):
        self.issues.append((self.country, severity, bid, description, handling))

    def check(self, name, value, tolerance):
        self.checks.append((self.country, name, value, tolerance, bool(abs(value) <= tolerance)))
