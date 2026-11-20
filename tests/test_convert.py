import csv
import io
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "examples"))

from cobieregister.__main__ import main  # noqa: E402
from cobieregister.convert import FIELDS, convert, map_category, to_csv  # noqa: E402

SAMPLE = str(ROOT / "examples" / "cobie_handover.xlsx")


@pytest.fixture(scope="module")
def res():
    return convert(SAMPLE)


def by_name(res, name):
    return next(r for r in res.rows if r["name"] == name)


def make(tmp_path, sheets):
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    p = tmp_path / "w.xlsx"
    wb.save(p)
    return str(p)


def test_four_components_converted(res):
    assert len(res.rows) == 4


def test_ids_follow_register_pattern(res):
    import re
    assert all(re.fullmatch(r"[A-Z]{2,4}-[0-9]{4}", r["asset_id"]) for r in res.rows)


def test_ids_count_per_prefix(res):
    assert [r["asset_id"] for r in res.rows if r["category"] == "HVAC"] == ["HVAC-0001", "HVAC-0002"]


def test_category_mapping(res):
    assert by_name(res, "LP-201")["category"] == "Electrical"
    assert by_name(res, "DK-101")["category"] == "Other"
    assert by_name(res, "HP-101")["category"] == "HVAC"


def test_location_includes_floor(res):
    assert by_name(res, "HP-101")["location"] == "101 (Level 1)"
    assert by_name(res, "LP-201")["location"] == "201 (Level 2)"


def test_building_and_manufacturer(res):
    r = by_name(res, "HP-101")
    assert r["building"] == "Harbor Office" and r["manufacturer"] == "Acme Climate"


def test_install_year_and_life(res):
    r = by_name(res, "HP-101")
    assert r["install_year"] == "2025" and r["expected_life_years"] == "18"
    assert by_name(res, "DK-101")["expected_life_years"] == ""


def test_csv_has_register_columns(res):
    rows = list(csv.DictReader(io.StringIO(to_csv(res))))
    assert list(rows[0].keys()) == FIELDS and rows[0]["status"] == "Active"


def test_clean_sample_has_no_warnings(res):
    assert res.warnings == []


def test_map_category_keywords():
    assert map_category("Chiller")[0] == "HVAC"
    assert map_category("Fire Alarm")[1] == "FIRE"
    assert map_category("Passenger Lift")[0] == "Elevator"
    assert map_category("")[0] == "Other"


def test_missing_type_and_space_warn(tmp_path):
    p = make(tmp_path, {"Component": [["Name", "TypeName", "Space", "InstallationDate"], ["X-1", "Ghost", "Nowhere", "soon"]]})
    out = convert(p)
    assert len(out.warnings) == 3 and out.rows[0]["category"] == "Other"


def test_date_cell_gives_year(tmp_path):
    from datetime import datetime
    p = make(tmp_path, {"Component": [["Name", "TypeName", "Space", "InstallationDate"], ["X-1", "T", "S", datetime(2019, 5, 1)]]})
    assert convert(p).rows[0]["install_year"] == "2019"


def test_no_component_sheet(tmp_path):
    p = make(tmp_path, {"Type": [["Name"], ["T"]]})
    assert convert(p).warnings == ["workbook has no Component sheet"]


def test_cli_codes(tmp_path, capsys):
    out = tmp_path / "r.csv"
    assert main(["convert", SAMPLE, "-o", str(out)]) == 0
    assert out.read_text().count("\n") == 5
    assert main(["convert", SAMPLE, "--strict"]) == 0
    assert main(["convert", "nope.xlsx"]) == 2
    bad = tmp_path / "b.xlsx"
    bad.write_text("x")
    assert main(["convert", str(bad)]) == 2
