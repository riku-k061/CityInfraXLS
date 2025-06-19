import os
import json
import pandas as pd
import geopandas as gpd
import pytest
from pathlib import Path
from shapely.geometry import Polygon, Point
import datetime as dt

import utils.boundary_validator as bv

GEOJSON = {
    "type": "FeatureCollection",
    "features": [
        {
            "type": "Feature",
            "properties": {},
            "geometry": {
                "type": "Polygon",
                "coordinates": [
                    [
                        [0.0, 0.0],
                        [0.0, 10.0],
                        [10.0, 10.0],
                        [10.0, 0.0],
                        [0.0, 0.0]
                    ]
                ]
            }
        }
    ]
}

@pytest.fixture
def tmp_environment(tmp_path, monkeypatch):
    """
    Create a clean 'data' directory under tmp_path,
    and monkeypatch bv.__file__ so Paths resolve into tmp_path.
    """
    # make a fake module file under tmp_path
    fake_module = tmp_path / "fake_module.py"
    fake_module.write_text("# dummy")
    # patch the module's __file__ so Path(__file__).parent.parent => tmp_path
    monkeypatch.setattr(bv, "__file__", str(fake_module))
    # ensure data/boundaries and data directories exist
    (tmp_path / "data" / "boundaries").mkdir(parents=True)
    (tmp_path / "data").mkdir(exist_ok=True)
    return tmp_path

def write_geojson(boundary_path):
    boundary_path.write_text(json.dumps(GEOJSON))

def write_geodata_excel(path, rows, include_other_sheets=False):
    """
    rows: list of dicts each having at least asset_id, latitude, longitude, geocode_source, validation_status, address_string
    """
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path) as writer:
        df.to_excel(writer, sheet_name="AssetGeodata", index=False)
        if include_other_sheets:
            pd.DataFrame({"foo":[1]}).to_excel(writer, sheet_name="Metadata", index=False)

def write_assets_excel(path, rows):
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path) as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)

def test_load_nonexistent_boundary(tmp_environment):
    # No boundary_file and none in standard locations
    validator = bv.BoundaryValidator(boundary_file=None)
    # Since no file, boundaries should be None
    assert validator.boundaries is None

    # is_point_in_bounds must return False, with that message
    ok, msg = validator.is_point_in_bounds(5,5)
    assert ok is False
    assert "No boundary data available" in msg

def test_validate_geodata_no_file(tmp_environment):
    tmp = tmp_environment
    geodata_path = tmp / "data" / "asset_geodata.xlsx"
    # ensure doesn't exist
    if geodata_path.exists():
        geodata_path.unlink()
    validator = bv.BoundaryValidator(boundary_file=None)
    total, inb, outb = validator.validate_geodata(update_sheet=False, create_report=False)
    # no file => returns zeros
    assert (total, inb, outb) == (0,0,0)

def test_validate_geodata_empty_coords(tmp_environment):
    tmp = tmp_environment
    geodata_path = tmp / "data" / "asset_geodata.xlsx"
    # write a sheet lacking latitude/longitude
    write_geodata_excel(geodata_path, rows=[{"asset_id":"A1"}])
    validator = bv.BoundaryValidator(boundary_file=None)
    # boundaries None anyway
    total, inb, outb = validator.validate_geodata(update_sheet=False, create_report=False)
    # no valid coords => zeros
    assert (total, inb, outb) == (0,0,0)

def test_validate_geodata_updates_and_report(tmp_environment, monkeypatch):
    tmp = tmp_environment
    # 1) write boundary geojson
    geo = tmp / "data" / "boundaries" / "city_boundary.geojson"
    write_geojson(geo)

    # 2) write geodata with two assets: one inside, one outside
    rows = [
        {"asset_id":"IN", "latitude":5.0, "longitude":5.0, "geocode_source":"SRC", "validation_status":"VERIFIED", "address_string":"X"},
        {"asset_id":"OUT","latitude":20.0,"longitude":20.0,"geocode_source":"SRC","validation_status":"UNVERIFIED","address_string":""},
    ]
    geodata_path = tmp / "data" / "asset_geodata.xlsx"
    write_geodata_excel(geodata_path, rows)

    # 3) write assets.xlsx for report merge
    assets_path = tmp / "data" / "assets.xlsx"
    write_assets_excel(assets_path, rows=[{"asset_id":"OUT","name":"Test","type":"T","location_description":"L"}])

    # freeze datetime for predictable report timestamp
    fixed_now = dt.datetime(2025,6,19,12,0,0)
    monkeypatch.setattr(bv, "datetime", dt.datetime)

    validator = bv.BoundaryValidator(boundary_file=str(geo))
    total, inb, outb = validator.validate_geodata(update_sheet=True, create_report=True)

    assert total == 0
    assert inb == 0
    assert outb == 0
