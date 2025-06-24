import os
import pandas as pd
import pytest
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook

from utils.geodata_sync import GeodataAssetSynchronizer

@pytest.fixture
def tmp_files(tmp_path):
    # Paths for geodata and asset registry
    geodata_file = tmp_path / "geodata.xlsx"
    registry_file = tmp_path / "assets.xlsx"
    return geodata_file, registry_file

def write_geodata(path, rows, include_optional=False):
    cols = ['asset_id', 'latitude', 'longitude']
    if include_optional:
        cols += ['geocode_source', 'geocode_timestamp', 'validation_status', 'accuracy_meters']
    df = pd.DataFrame(rows, columns=cols)
    with pd.ExcelWriter(path) as w:
        df.to_excel(w, sheet_name="Geodata", index=False)

def write_registry(path, rows, cols=None):
    if cols is None:
        cols = ['asset_id']
    df = pd.DataFrame(rows, columns=cols)
    df.to_excel(path, index=False)

def test_validate_files_missing_geodata(tmp_files):
    geodata_path, registry_path = tmp_files
    # registry exists but geodata missing
    write_registry(registry_path, [[ "A1" ]])
    sync = GeodataAssetSynchronizer(str(geodata_path), str(registry_path))
    valid, msg = sync._validate_files()
    assert not valid
    assert "Geodata file not found" in msg

def test_validate_files_missing_sheet(tmp_files):
    geodata_path, registry_path = tmp_files
    # create geodata with wrong sheet
    df = pd.DataFrame([{"foo":1}])
    with pd.ExcelWriter(geodata_path) as w:
        df.to_excel(w, sheet_name="Wrong", index=False)
    write_registry(registry_path, [["A1"]])
    sync = GeodataAssetSynchronizer(str(geodata_path), str(registry_path))
    valid, msg = sync._validate_files()
    assert not valid
    assert "Error reading geodata file" in msg

def test_validate_files_missing_columns(tmp_files):
    geodata_path, registry_path = tmp_files
    # create geodata with missing cols
    write_geodata(geodata_path, [["A1", None, None]], include_optional=False)
    # drop latitude/longitude
    wb = load_workbook(geodata_path)
    ws = wb["Geodata"]
    # rewrite header to only asset_id
    ws.delete_cols(2,2)
    wb.save(geodata_path)
    write_registry(registry_path, [["A1"]])
    sync = GeodataAssetSynchronizer(str(geodata_path), str(registry_path))
    valid, msg = sync._validate_files()
    assert not valid
    assert "Missing required columns" in msg

def test_validate_files_missing_registry_asset_id(tmp_files):
    geodata_path, registry_path = tmp_files
    # correct geodata
    write_geodata(geodata_path, [["A1", 1.0, 2.0]])
    # registry missing asset_id column
    write_registry(registry_path, [[1]], cols=["foo"])
    sync = GeodataAssetSynchronizer(str(geodata_path), str(registry_path))
    valid, msg = sync._validate_files()
    assert not valid
    assert "must have 'asset_id' column" in msg

def test_validate_files_success(tmp_files):
    geodata_path, registry_path = tmp_files
    write_geodata(geodata_path, [["A1", 1.0, 2.0]])
    write_registry(registry_path, [["A1"]])
    sync = GeodataAssetSynchronizer(str(geodata_path), str(registry_path))
    valid, msg = sync._validate_files()
    assert valid
    assert msg == ""

def test_prepare_and_add_metadata():
    sync = GeodataAssetSynchronizer.__new__(GeodataAssetSynchronizer)
    # prepare registry lacking all sync fields
    asset_df = pd.DataFrame([{'asset_id':'A1'}])
    cols = ['asset_id','latitude','longitude','geocode_source']
    updated = sync._prepare_asset_registry(asset_df, cols)
    assert 'latitude' in updated.columns
    assert 'geocode_source' in updated.columns
    # test metadata addition
    meta_df = sync._add_sync_metadata(updated)
    assert 'geodata_sync_timestamp' in meta_df.columns

def test_synchronize_skip_and_not_found(tmp_files):
    geodata_path, registry_path = tmp_files
    # geodata two assets
    write_geodata(geodata_path, [
        ["A1", 1.0, 2.0],
        ["Z9", 3.0, 4.0]
    ], include_optional=False)
    # registry with A1 having existing coords and B2
    df_reg = pd.DataFrame([
        {'asset_id':'A1','latitude':9.9,'longitude':9.9},
        {'asset_id':'B2'}
    ])
    df_reg.to_excel(registry_path, index=False)

    sync = GeodataAssetSynchronizer(str(geodata_path), str(registry_path))
    stats = sync.synchronize(force_update=False, backup=False)
    # A1 skipped, Z9 not found
    assert stats["assets_updated"] == 0
    assert stats["assets_skipped"] == 1
    assert stats["assets_not_found"] == 1

def test_generate_report():
    sync = GeodataAssetSynchronizer.__new__(GeodataAssetSynchronizer)
    # Manually set stats
    sync.stats = {
        "total_assets": 5,
        "assets_with_geodata": 3,
        "assets_updated": 2,
        "assets_skipped": 1,
        "assets_not_found": 0
    }
    sync.geodata_file = "g.xlsx"
    sync.asset_registry_file = "a.xlsx"
    report = sync.generate_report()
    assert "Geodata Synchronization Report" in report
    assert "Total assets in registry: 5" in report
    assert "Assets with geodata: 3" in report
    assert "Assets updated: 2" in report
