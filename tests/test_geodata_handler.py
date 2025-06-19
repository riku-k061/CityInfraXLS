import os
import pandas as pd
import pytest
from pathlib import Path
import datetime as dt

from utils.geodata_handler import GeodataHandler

# A minimal schema for testing
SAMPLE_SCHEMA = {
    "properties": {
        "asset_id":              {"type": "string"},
        "latitude":              {"type": "number"},
        "longitude":             {"type": "number"},
        "geocode_timestamp":     {"type": "string", "format": "date-time"},
        "validation_status":     {"type": "string", "enum": ["VERIFIED", "UNVERIFIED"]},
    },
    "required": ["asset_id", "latitude", "longitude"]
}

@pytest.fixture
def handler(tmp_path, monkeypatch):
    """
    Build a GeodataHandler instance but override:
    - schema with SAMPLE_SCHEMA
    - data_path to a temp file
    """
    # Create a new, uninitialized instance
    inst = GeodataHandler.__new__(GeodataHandler)
    # Inject our sample schema
    inst.schema = SAMPLE_SCHEMA
    # Point data_path into tmp_path
    inst.data_path = tmp_path / "asset_geodata.xlsx"
    # Ensure the parent directory exists
    inst.data_path.parent.mkdir(parents=True, exist_ok=True)
    yield inst

def init_empty_file(path, columns):
    """
    Create an Excel file at `path` with one sheet "AssetGeodata"
    and only the header row of `columns`.
    """
    df = pd.DataFrame(columns=columns)
    with pd.ExcelWriter(path) as writer:
        df.to_excel(writer, sheet_name="AssetGeodata", index=False)

def test_validate_geodata_success_and_failure(handler):
    valid = {"asset_id": "A1", "latitude": 52.5, "longitude": 13.4}
    ok, msg = handler.validate_geodata(valid)
    assert ok is True
    assert msg == ""

    # Missing required field
    invalid = {"asset_id": "A2", "latitude": 10.0}
    ok2, msg2 = handler.validate_geodata(invalid)
    assert ok2 is False
    assert "longitude" in msg2

    # Wrong type
    invalid2 = {"asset_id": "A3", "latitude": "not-a-number", "longitude": 0}
    ok3, msg3 = handler.validate_geodata(invalid2)
    assert ok3 is False
    assert "latitude" in msg3

def test_create_geodata_file_creates_all_sheets(handler, tmp_path):
    # Ensure no file exists yet
    assert not handler.data_path.exists()

    # Call create
    handler.create_geodata_file()

    # File should now exist
    assert handler.data_path.exists()

    # Read back
    xls = pd.ExcelFile(handler.data_path)
    assert "AssetGeodata" in xls.sheet_names
    assert "Metadata" in xls.sheet_names
    # Our SAMPLE_SCHEMA has an enum on 'validation_status'
    assert "AllowedValues" in xls.sheet_names

    # Check that AssetGeodata has correct columns
    df = pd.read_excel(handler.data_path, sheet_name="AssetGeodata")
    assert list(df.columns) == list(SAMPLE_SCHEMA["properties"].keys())

    # Metadata row should indicate types
    meta = pd.read_excel(handler.data_path, sheet_name="Metadata", index_col=0)
    # The index 'data_type' must exist
    assert "data_type" in meta.index
    # 'latitude' should map to 'float'
    assert meta.at["data_type", "latitude"] == "float"

    # AllowedValues should list the two enums for validation_status
    allowed = pd.read_excel(handler.data_path, sheet_name="AllowedValues")
    # Should have column 'validation_status'
    assert "validation_status" in allowed.columns
    # The first two rows should be "VERIFIED", "UNVERIFIED"
    vals = allowed["validation_status"].dropna().tolist()
    assert set(vals) == {"VERIFIED", "UNVERIFIED"}

def test_add_and_get_geodata_new_and_update(handler):
    # Prepare empty file with only headers
    init_empty_file(handler.data_path, list(SAMPLE_SCHEMA["properties"].keys()))

    # Add a new record without timestamp or status
    record = {"asset_id": "X100", "latitude": 1.23, "longitude": 4.56}
    ok, msg = handler.add_geodata(record)
    assert ok is True
    assert "successfully" in msg.lower()

    # Reload and check via pandas
    df = pd.read_excel(handler.data_path, sheet_name="AssetGeodata")
    assert df.shape[0] == 1
    row = df.iloc[0].to_dict()
    # asset_id, latitude, longitude stored
    assert row["asset_id"] == "X100"
    assert pytest.approx(row["latitude"], 0.001) == 1.23
    assert pytest.approx(row["longitude"], 0.001) == 4.56
    # geocode_timestamp was auto‐filled as ISO string
    assert isinstance(row["geocode_timestamp"], str)
    # validation_status default
    assert row["validation_status"] == "UNVERIFIED"

    # Now update the same asset: change latitude and set status
    updated = {
        "asset_id": "X100",
        "latitude": 9.87,
        "longitude": 6.54,
        "validation_status": "VERIFIED",
        "geocode_timestamp": dt.datetime(2025, 1, 1).isoformat()
    }
    ok2, msg2 = handler.add_geodata(updated)
    assert ok2 is True

    df2 = pd.read_excel(handler.data_path, sheet_name="AssetGeodata")
    assert df2.shape[0] == 1  # still one row
    row2 = df2.iloc[0].to_dict()
    assert pytest.approx(row2["latitude"], 0.001) == 9.87
    assert row2["validation_status"] == "VERIFIED"

    # Test get_asset_geodata
    fetched = handler.get_asset_geodata("X100")
    assert isinstance(fetched, dict)
    assert fetched["asset_id"] == "X100"
    assert pytest.approx(fetched["longitude"], 0.001) == 6.54

def test_add_geodata_validation_failure(handler):
    # No file on disk, but validation fails first
    bad = {"asset_id": "B1", "latitude": 0}  # missing longitude
    ok, msg = handler.add_geodata(bad)
    assert ok is False
    assert "Validation failed" in msg

    # No file should have been created
    assert not handler.data_path.exists()

def test_get_asset_geodata_not_found(handler):
    # Prepare empty file
    init_empty_file(handler.data_path, list(SAMPLE_SCHEMA["properties"].keys()))
    # Query missing asset
    assert handler.get_asset_geodata("ZZZ") is None
