# tests/test_incidents.py

import pytest
import pandas as pd
import json
import uuid
from report_incident import main as report_main
from query_incidents import load_incidents_data
from delete_incident import find_incident, delete_incident as delete_row
from utils.incident_handler import create_incident_sheet

# ---------------- Sample Severity Matrix ----------------
SEVERITY_MATRIX = {
    "Critical": {
        "hours": 24,
        "description": "Critical infrastructure failure"
    },
    "High": {
        "hours": 48,
        "description": "Major service disruption"
    }
}

# ---------------- Fixtures ----------------
@pytest.fixture
def setup_incident_paths(tmp_path, monkeypatch):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    matrix_path = tmp_path / "severity_matrix.json"
    incidents_path = data_dir / "incidents.xlsx"

    with open(matrix_path, "w") as f:
        json.dump(SEVERITY_MATRIX, f)

    create_incident_sheet(str(incidents_path))

    # Patch for report_incident
    monkeypatch.setattr("report_incident.ensure_incident_sheet", lambda: str(incidents_path))
    monkeypatch.setattr("report_incident.validate_severity_matrix", lambda _: SEVERITY_MATRIX)
    monkeypatch.setattr("report_incident.load_severity_matrix", lambda: SEVERITY_MATRIX)

    return {
        "matrix_path": matrix_path,
        "incidents_path": incidents_path
    }

# ---------------- Helpers ----------------
def simulate_incident_input(monkeypatch):
    inputs = iter([
        "R001",       # Asset ID
        "Sinkhole",   # Issue Type
        "TestUser",   # Reporter
        "1"           # Severity selection ("Critical")
    ])
    monkeypatch.setattr("builtins.input", lambda _: next(inputs))
    monkeypatch.setattr("uuid.uuid4", lambda: uuid.UUID("00000000-0000-0000-0000-000000000123"))

# ---------------- Tests ----------------
def test_report_incident(setup_incident_paths, monkeypatch, capsys):
    simulate_incident_input(monkeypatch)
    report_main()
    capsys.readouterr()  # Clear buffer

    df = pd.read_excel(setup_incident_paths["incidents_path"])
    assert not df.empty

def test_query_loaded_incident(setup_incident_paths, monkeypatch, capsys):
    simulate_incident_input(monkeypatch)
    report_main()
    capsys.readouterr()

    df = load_incidents_data()
    assert not df.empty
    assert df.iloc[0]["Severity"] == "Critical"
    assert df.iloc[0]["Status"] == "Open"

def test_find_and_delete_incident(setup_incident_paths, monkeypatch, capsys):
    simulate_incident_input(monkeypatch)
    report_main()
    capsys.readouterr()

    incident_id = "00000000-0000-0000-0000-000000000123"
    path = str(setup_incident_paths["incidents_path"])
    found, incident, row = find_incident(path, incident_id)
    assert found
    assert incident["Asset ID"] == "R001"

    delete_row(path, row)
    df = pd.read_excel(path)
    assert df.empty

def test_find_nonexistent_incident(setup_incident_paths):
    path = str(setup_incident_paths["incidents_path"])
    found, data, row = find_incident(path, "non-existent-id")
    assert not found
    assert data is None
    assert row is None