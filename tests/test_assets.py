# tests/test_assets.py

import pytest
import pandas as pd
import json
import os

from register_asset import register_asset
from query_assets import query_assets
from delete_asset import delete_asset, find_asset

# ---------------- Schema Mock ----------------
SCHEMA_CONTENT = {
    "Road": ["ID", "Name", "Location", "Length", "Width", "Surface Type", "Condition", "Installation Date"],
}

# ---------------- Fixture to Setup Paths ----------------
@pytest.fixture
def setup_paths(tmp_path, monkeypatch):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    schema_path = tmp_path / "asset_schema.json"
    with open(schema_path, "w") as f:
        json.dump(SCHEMA_CONTENT, f)

    assets_path = data_dir / "assets.xlsx"
    log_path = data_dir / "asset_log.xlsx"

    monkeypatch.setattr("register_asset.SCHEMA_PATH", str(schema_path))
    monkeypatch.setattr("register_asset.ASSETS_PATH", str(assets_path))
    monkeypatch.setattr("register_asset.LOG_PATH", str(log_path))
    monkeypatch.setattr("query_assets.ASSETS_PATH", str(assets_path))
    monkeypatch.setattr("query_assets.SCHEMA_PATH", str(schema_path))
    monkeypatch.setattr("delete_asset.ASSETS_PATH", str(assets_path))
    monkeypatch.setattr("delete_asset.LOG_PATH", str(log_path))

    return {
        "schema_path": schema_path,
        "assets_path": assets_path,
        "log_path": log_path,
        "data_dir": data_dir
    }

# ---------------- Helper to Register a Test Asset ----------------
def create_test_asset(setup_paths, monkeypatch, capsys):
    inputs = [
        "1", "Main Street", "Downtown", "1000", "20",
        "Asphalt", "Good", "2022-01-15"
    ]
    input_mock = iter(inputs)
    monkeypatch.setattr("builtins.input", lambda _: next(input_mock))
    monkeypatch.setattr("uuid.uuid4", lambda: type("obj", (object,), {"__str__": lambda self: "00000000-0000-0000-0000-000000000001"})())
    register_asset()
    capsys.readouterr()
    return "00000000-0000-0000-0000-000000000001"

# ---------------- Test Cases ----------------

def test_register_asset(setup_paths, monkeypatch, capsys):
    asset_id = create_test_asset(setup_paths, monkeypatch, capsys)
    df = pd.read_excel(setup_paths["assets_path"], sheet_name="Road")
    assert df.iloc[0]["ID"] == asset_id
    log = pd.read_excel(setup_paths["log_path"])
    assert log.iloc[0]["Asset ID"] == asset_id
    assert log.iloc[0]["Action"] == "REGISTER"

def test_query_assets(setup_paths, monkeypatch, capsys):
    asset_id = create_test_asset(setup_paths, monkeypatch, capsys)
    capsys.readouterr()

    # 1. Query by type
    query_assets(asset_type="Road")
    captured = capsys.readouterr()
    assert "Main Street" in captured.out
    assert "Total: 1 assets found" in captured.out

    # 2. Location filter
    query_assets(location="Downtown")
    assert "Main Street" in capsys.readouterr().out

    # 3. Installed after
    query_assets(installed_after="2022-01-01")
    assert "Main Street" in capsys.readouterr().out

    # 4. Non-match
    query_assets(location="Suburb")
    assert "No assets match" in capsys.readouterr().out

    # 5. Export test
    export_path = os.path.join(setup_paths["data_dir"], "export.xlsx")
    query_assets(asset_type="Road", export_path=export_path)
    assert os.path.exists(export_path)
    df = pd.read_excel(export_path)
    assert df.iloc[0]["Name"] == "Main Street"

def test_find_asset(setup_paths, monkeypatch, capsys):
    asset_id = create_test_asset(setup_paths, monkeypatch, capsys)
    sheet, idx, data = find_asset(asset_id)
    assert sheet == "Road"
    assert idx == 2
    assert data["ID"] == asset_id

    # Non-existent
    sheet, idx, data = find_asset("fake-id")
    assert sheet is None and idx is None and data is None

def test_delete_asset(setup_paths, monkeypatch, capsys):
    asset_id = create_test_asset(setup_paths, monkeypatch, capsys)
    monkeypatch.setattr("builtins.input", lambda _: "y")
    result = delete_asset(asset_id)
    output = capsys.readouterr().out
    assert result is True
    assert "Success!" in output

    df = pd.read_excel(setup_paths["assets_path"], sheet_name="Road")
    assert len(df) == 0
    log = pd.read_excel(setup_paths["log_path"])
    assert len(log) == 2
    assert log.iloc[1]["Action"] == "DELETE"

    # Try deleting again
    result = delete_asset("non-existent-id", confirm=False)
    assert result is False
    assert "not found" in capsys.readouterr().out

def test_delete_without_confirmation(setup_paths, monkeypatch, capsys):
    asset_id = create_test_asset(setup_paths, monkeypatch, capsys)
    result = delete_asset(asset_id, confirm=False)
    output = capsys.readouterr().out
    assert result is True
    assert "Success!" in output

def test_delete_cancelled(setup_paths, monkeypatch, capsys):
    asset_id = create_test_asset(setup_paths, monkeypatch, capsys)
    monkeypatch.setattr("builtins.input", lambda _: "n")
    result = delete_asset(asset_id)
    output = capsys.readouterr().out
    assert result is False
    assert "cancelled" in output