"""Convert a COBie handover workbook into a facility asset register CSV."""

from __future__ import annotations

import csv
import io
from dataclasses import dataclass, field
from datetime import date, datetime

from openpyxl import load_workbook

FIELDS = ["asset_id", "name", "category", "building", "location", "manufacturer",
          "install_year", "expected_life_years", "condition", "replacement_cost", "status"]

# category -> (ID prefix, keywords found in the COBie type category or name)
CATEGORIES = {
    "HVAC": ("HVAC", ("hvac", "heat", "cool", "air", "ventilat", "chiller", "boiler", "pump", "fan")),
    "Electrical": ("ELEC", ("electric", "light", "power", "generator", "switch", "panel")),
    "Plumbing": ("PLMB", ("plumb", "water", "drain", "sanit", "valve")),
    "Fire Safety": ("FIRE", ("fire", "alarm", "sprinkler", "smoke", "extinguish")),
    "Elevator": ("ELEV", ("elevator", "lift", "escalator")),
    "Envelope": ("ENVL", ("roof", "window", "door", "facade", "envelope", "curtain")),
}
OTHER = ("Other", "OTH")


@dataclass
class Result:
    rows: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def map_category(*texts: str) -> tuple[str, str]:
    """Return (register category, ID prefix) for a COBie category or type name."""
    hay = " ".join(t.lower() for t in texts if t)
    for cat, (prefix, words) in CATEGORIES.items():
        if any(w in hay for w in words):
            return cat, prefix
    return OTHER


def _text(v) -> str:
    return "" if v is None else str(v).strip()


def _sheet(wb, name: str) -> list[dict]:
    if name not in wb.sheetnames:
        return []
    rows = list(wb[name].iter_rows(values_only=True))
    if not rows:
        return []
    head = [_text(h) for h in rows[0]]
    return [{h: r[i] if i < len(r) else None for i, h in enumerate(head) if h} for r in rows[1:]
            if any(_text(v) for v in r)]


def _year(v) -> str:
    if isinstance(v, (datetime, date)):
        return str(v.year)
    s = _text(v)
    return s[:4] if len(s) >= 4 and s[:4].isdigit() else ""


def _whole(v) -> str:
    try:
        n = float(v)
    except (TypeError, ValueError):
        return ""
    return str(int(n)) if n == int(n) and n > 0 else ""


def convert(path: str) -> Result:
    """Build register rows from the Component, Type, Space, Floor and Facility sheets."""
    wb = load_workbook(path, read_only=True, data_only=True)
    types = {_text(r.get("Name")): r for r in _sheet(wb, "Type")}
    spaces = {_text(r.get("Name")): r for r in _sheet(wb, "Space")}
    facility = _sheet(wb, "Facility")
    building = _text(facility[0].get("Name")) if facility else ""
    res = Result()
    counters: dict[str, int] = {}
    if "Component" not in wb.sheetnames:
        res.warnings.append("workbook has no Component sheet")
        return res
    for r in _sheet(wb, "Component"):
        name = _text(r.get("Name"))
        t = types.get(_text(r.get("TypeName")))
        if t is None:
            res.warnings.append(f"{name}: type '{_text(r.get('TypeName'))}' not found, category set to Other")
        cat, prefix = map_category(_text(t.get("Category")) if t else "", _text(r.get("TypeName")))
        counters[prefix] = counters.get(prefix, 0) + 1
        sp = spaces.get(_text(r.get("Space")))
        if sp is not None:
            location = f"{_text(r.get('Space'))} ({_text(sp.get('FloorName'))})"
        else:
            location = _text(r.get("Space"))
            res.warnings.append(f"{name}: space '{location}' not found, location left as given")
        year = _year(r.get("InstallationDate"))
        if not year:
            res.warnings.append(f"{name}: no usable InstallationDate")
        res.rows.append({
            "asset_id": f"{prefix}-{counters[prefix]:04d}", "name": name, "category": cat,
            "building": building, "location": location,
            "manufacturer": _text(t.get("Manufacturer")) if t else "",
            "install_year": year, "expected_life_years": _whole(t.get("ExpectedLife")) if t else "",
            "condition": "", "replacement_cost": "", "status": "Active",
        })
    return res


def to_csv(res: Result) -> str:
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=FIELDS, lineterminator="\n")
    w.writeheader()
    w.writerows(res.rows)
    return buf.getvalue()
