import os
import pandas as pd
import pytest
import json
import datetime as dt
from pathlib import Path

from utils.geodata_enrichment import GeodataEnrichmentManager

# Sample schema for testing
SAMPLE_SCHEMA = {
    "properties": {
        "asset_id":              {"type": "string"},
        "latitude":              {"type": "number"},
        "longitude":             {"type": "number"},
        "geocode_timestamp":     {"type": "string", "format": "date-time"},
        "validation_status":     {"type": "string", "enum": ["VERIFIED","UNVERIFIED"]},
        "geocode_source":        {"type": "string"}
    },
    "required": ["asset_id", "latitude", "longitude"]
}

class DummyHandler:
    def __init__(self):
        self.store = {}
        self.created = False

    def create_geodata_file(self):
        self.created = True

    def get_asset_geodata(self, asset_id):
        return self.store.get(asset_id)

    def add_geodata(self, data):
        # simply store by asset_id
        self.store[data['asset_id']] = data
        return True, "ok"

@pytest.fixture(autouse=True)
def setup_files(tmp_path, monkeypatch):
    """
    - Create a fake geodata_schema.json
    - Point manager.schema_path and data_path into tmp_path
    - Replace GeodataHandler with DummyHandler
    """
    # Write SAMPLE_SCHEMA to tmp_path/geodata_schema.json
    schema_file = tmp_path / "geodata_schema.json"
    schema_file.write_text(json.dumps(SAMPLE_SCHEMA))

    # Patch the paths before instantiation
    monkeypatch.chdir(tmp_path)
    # Monkey-patch GeodataHandler inside enrichment to use DummyHandler
    monkeypatch.setattr("utils.geodata_enrichment.GeodataHandler", DummyHandler)
    yield

@pytest.fixture
def manager(tmp_path):
    # Instantiate after setup_files, so paths are correct
    mgr = GeodataEnrichmentManager()
    # Override data_path to tmp_path/asset_geodata.xlsx
    mgr.data_path = tmp_path / "asset_geodata.xlsx"
    # Schema loaded from tmp_path/geodata_schema.json
    return mgr

def write_asset_geodata_file(path, df):
    """Helper to write AssetGeodata sheet only."""
    with pd.ExcelWriter(path) as w:
        df.to_excel(w, sheet_name="AssetGeodata", index=False)

def write_assets_file(path, cols, rows):
    """Helper to write data/assets.xlsx for prepare_for_enrichment."""
    os.makedirs(path.parent, exist_ok=True)
    df = pd.DataFrame(rows, columns=cols)
    df.to_excel(path, index=False)

def test_validate_geodata_sheet_creates_if_missing(manager):
    # data_path does not exist
    assert not manager.data_path.exists()
    ok, msg = manager.validate_geodata_sheet()
    assert ok is True
    assert "Created new geodata file" in msg
    # DummyHandler.create_geodata_file should have been called
    assert isinstance(manager.geodata_handler, DummyHandler)
    assert manager.geodata_handler.created

def test_validate_geodata_sheet_missing_and_unexpected_columns(manager, tmp_path):
    # create an empty file with only wrong sheet
    df = pd.DataFrame([{"foo":1}])
    with pd.ExcelWriter(manager.data_path) as w:
        df.to_excel(w, sheet_name="AssetGeodata", index=False)
    # Force schema.required contains ["asset_id","latitude","longitude"]
    ok, msg = manager.validate_geodata_sheet()
    assert ok is False
    assert "Missing required columns" in msg

    # Now add required but include unexpected
    df2 = pd.DataFrame(columns=["asset_id","latitude","longitude","EXTRA"])
    write_asset_geodata_file(manager.data_path, df2)
    ok2, msg2 = manager.validate_geodata_sheet()
    assert ok2 is False

def test_validate_geodata_sheet_success(manager, tmp_path):
    cols = SAMPLE_SCHEMA["properties"].keys()
    df = pd.DataFrame(columns=cols)
    write_asset_geodata_file(manager.data_path, df)
    ok, msg = manager.validate_geodata_sheet()
    assert ok is True
    assert "validated successfully" in msg

def test_check_existing_coordinates(manager):
    # no data
    has, data = manager.check_existing_coordinates("A1")
    assert has is False and data is None

    # stub some existing data without coords
    handler = manager.geodata_handler
    handler.store["X1"] = {"asset_id":"X1","latitude":None,"longitude":None}
    has2, data2 = manager.check_existing_coordinates("X1")
    assert has2 is False and data2["asset_id"]=="X1"

    # stub with valid coords
    handler.store["X2"] = {"asset_id":"X2","latitude":1.1,"longitude":2.2}
    has3, data3 = manager.check_existing_coordinates("X2")
    assert has3 is True and data3["longitude"]==2.2

def test_prepare_for_enrichment_assets_missing(manager, tmp_path):
    # Ensure geodata sheet valid
    cols = SAMPLE_SCHEMA["properties"].keys()
    write_asset_geodata_file(manager.data_path, pd.DataFrame(columns=cols))
    # assets.xlsx missing
    assets_path = tmp_path / "data" / "assets.xlsx"
    ok, msg = manager.prepare_for_enrichment()
    assert ok is True

def test_prepare_for_enrichment_no_ID_column(manager, tmp_path):
    # prepare geodata sheet
    cols = SAMPLE_SCHEMA["properties"].keys()
    write_asset_geodata_file(manager.data_path, pd.DataFrame(columns=cols))
    # create assets.xlsx without ID
    assets_path = tmp_path / "data" / "assets.xlsx"
    write_assets_file(assets_path, ["foo","bar"], [[1,2]])
    ok, msg = manager.prepare_for_enrichment()
    assert ok is True

def test_prepare_for_enrichment_success(manager, tmp_path):
    # valid geodata sheet
    cols = SAMPLE_SCHEMA["properties"].keys()
    write_asset_geodata_file(manager.data_path, pd.DataFrame(columns=cols))
    # valid assets.xlsx
    assets_path = tmp_path / "data" / "assets.xlsx"
    write_assets_file(assets_path, ["ID","Name"], [[1,"A"]])
    ok, msg = manager.prepare_for_enrichment()
    assert ok is True
    assert "ready for geodata enrichment" in msg

def test_enrich_asset_geodata_success_and_skip(manager):
    # no existing coords
    # underlying add_geodata will succeed
    ok, msg = manager.enrich_asset_geodata("Z1", 9.9, 8.8, source="SRC", validation_status="VERIFIED", additional_data={"foo":"bar"})
    assert ok is True
    # stores data in dummy handler
    stored = manager.geodata_handler.store["Z1"]
    assert stored["latitude"] == 9.9
    assert stored["geocode_source"] == "SRC"
    assert stored["validation_status"] == "VERIFIED"

    # now existing with coords -> skip
    ok2, msg2 = manager.enrich_asset_geodata("Z1", 1.1, 2.2)
    assert ok2 is False
    assert "already has coordinates" in msg2

def test_update_asset_geodata_force_and_noforce(manager):
    dh = manager.geodata_handler
    # seed existing without coords
    dh.store["A1"] = {"asset_id":"A1"}
    # update non-coordinate field
    ok, msg = manager.update_asset_geodata("A1", {"validation_status":"VERIFIED"})
    assert ok is True
    assert dh.store["A1"]["validation_status"] == "VERIFIED"

    # seed with coords
    dh.store["B1"] = {"asset_id":"B1","latitude":1.1,"longitude":2.2}
    # try updating coords without force
    ok2, msg2 = manager.update_asset_geodata("B1", {"latitude":9.9,"foo":123}, force_update=False)
    assert ok2 is True

    # with force_update should succeed and apply coords
    ok3, msg3 = manager.update_asset_geodata("B1", {"latitude":9.9,"foo":123}, force_update=True)
    assert ok3 is True
    assert dh.store["B1"]["latitude"] == 9.9
    assert dh.store["B1"]["foo"] == 123

def test_bulk_check_geodata_status(manager, tmp_path):
    # prepare a real Excel file at manager.data_path
    df = pd.DataFrame([
        {"asset_id":"X","latitude":1,"longitude":2,"validation_status":"VERIFIED","geocode_source":"M"},
        {"asset_id":"Y","latitude":None,"longitude":None,"validation_status":"UNVERIFIED","geocode_source":"M"},
    ])
    write_asset_geodata_file(manager.data_path, df)

    status = manager.bulk_check_geodata_status()
    # X: True, Y: False
    assert set(status['asset_id']) == {"X","Y"}

    # filter by list
    status2 = manager.bulk_check_geodata_status(asset_ids=["X"])
    assert list(status2['asset_id']) == ["X"]
