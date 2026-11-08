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
    """Everything one country contributes, already in template vocabulary."""
    country: str
    source_file: str
    projects: list = field(default_factory=list)    # dicts keyed by template attribute codes (config.PROJECT_COLUMNS)
    results: list = field(default_factory=list)     # building × indicator × template module
    elements: list = field(default_factory=list)    # building × national element code (× module for LU)
    materials: list = field(default_factory=list)   # LU: building × material group × template module
    element_desc: dict = field(default_factory=dict)
    log: list = field(default_factory=list)
    issues: list = field(default_factory=list)
    checks: list = field(default_factory=list)

    def result(self, bid, indicator, module, value, source_value, source_unit, source_columns, covers='', flag=''):
        """One reported value. `covers` names the module range when one value spans several modules
        (template rule: entered in the first module of the range)."""
        self.results.append(dict(building_id=bid, indicator=indicator, module=module, value_kgco2e=value,
                                 source_value=source_value, source_unit=source_unit, source_columns=source_columns,
                                 covers=covers or module, flag=flag))

    def step(self, step, detail, count=''):
        self.log.append((self.country, step, detail, count))

    def issue(self, kind, bid, observation, handling, question=''):
        self.issues.append((self.country, kind, bid, observation, handling, question))

    def check(self, name, value, tolerance):
        self.checks.append((self.country, name, value, tolerance, bool(abs(value) <= tolerance)))
