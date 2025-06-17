import os
import re
import pandas as pd
import pytest
from datetime import datetime
from pathlib import Path
import query_maintenance

@pytest.fixture(autouse=True)
def tmp_cwd(tmp_path, monkeypatch):
    # Sandbox the cwd under tmp_path and ensure data dir exists
    monkeypatch.chdir(tmp_path)
    (tmp_path / "data").mkdir()
    return tmp_path

# --- parse_date tests ---

def test_parse_date_valid():
    dt = query_maintenance.parse_date("2025-06-14")
    assert isinstance(dt, datetime)
    assert dt.year == 2025 and dt.month == 6 and dt.day == 14

def test_parse_date_invalid():
    with pytest.raises(ValueError) as exc:
        query_maintenance.parse_date("14-06-2025")
    assert "Invalid date format" in str(exc.value)

# --- print_filters tests ---

def test_print_filters_all(capsys):
    query_maintenance.print_filters(None, None, None)
    out = capsys.readouterr().out.splitlines()
    assert out[0] == "Filters applied:"
    assert out[1].strip() == "From: [beginning]"
    assert out[2].strip() == "To: [present]"
    assert out[3].strip() == 'Action: [all]'

def test_print_filters_some(capsys):
    d1 = datetime(2025,1,1)
    d2 = datetime(2025,12,31)
    query_maintenance.print_filters(d1, d2, "inspect")
    out = capsys.readouterr().out.splitlines()
    assert "From: 2025-01-01" in out[1]
    assert "To: 2025-12-31" in out[2]
    assert 'Action: "inspect"' in out[3]

# --- query_maintenance tests ---

def test_query_no_file(tmp_cwd, capsys):
    # No maintenance_history.xlsx present
    result = query_maintenance.query_maintenance()
    out = capsys.readouterr().out
    assert result is False
    assert "not found" in out

def test_query_empty_sheet(tmp_cwd, capsys):
    # Create empty sheet
    path = tmp_cwd / "data" / "maintenance_history.xlsx"
    df_empty = pd.DataFrame(columns=["asset_id","date","cost","action_taken"])
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df_empty.to_excel(w, sheet_name="Maintenance History", index=False)

    result = query_maintenance.query_maintenance()
    out = capsys.readouterr().out
    assert result is False
    assert "sheet is empty" in out

@pytest.fixture
def sample_history(tmp_cwd):
    # Build a simple history with three records
    data = [
        {"asset_id":"A","date":"2025-01-01","cost":10,"action_taken":"Inspection"},
        {"asset_id":"B","date":"2025-06-01","cost":20,"action_taken":"Repair"},
        {"asset_id":"A","date":"2025-03-15","cost":15,"action_taken":"Inspection"}
    ]
    path = tmp_cwd / "data" / "maintenance_history.xlsx"
    df = pd.DataFrame(data)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Maintenance History", index=False)
    return path, pd.DataFrame(data)

def test_query_filters_and_no_export(sample_history, capsys):
    # Filter from 2025-02-01 to 2025-06-30, action "Inspect"
    path, _ = sample_history
    result = query_maintenance.query_maintenance(
        from_date="2025-02-01",
        to_date="2025-06-30",
        action="inspect",
        export=False
    )
    out = capsys.readouterr().out
    assert result is True
    # It should find two records: A on 03-15 and B doesn't match action
    assert "Found 1 maintenance record" in out
    assert "Filters applied:" in out
    assert "2025-02-01" in out and "2025-06-30" in out
    # The table grid should appear
    assert "+" in out  # rudimentary check for tabulate grid

def test_query_no_matches(sample_history, capsys):
    # Filter action that doesn't exist
    result = query_maintenance.query_maintenance(action="nonexistent")
    out = capsys.readouterr().out
    assert result is False
    assert "No records found matching the filter criteria" in out

def test_query_export_creates_file(sample_history, monkeypatch, capsys):
    # Freeze datetime so export filename is predictable
    fake_now = datetime(2025,6,14,12,0,0)
    class DummyDateTime(datetime):
        @classmethod
        def now(cls):
            return fake_now
    monkeypatch.setattr(query_maintenance, 'datetime', DummyDateTime)

    result = query_maintenance.query_maintenance(export=True)
    out = capsys.readouterr().out
    assert result is True

    # Expect an export file under data/ with timestamp
    pattern = r"data/maintenance_query_20250614_120000\.xlsx"
    # Find the filename in printed output
    m = re.search(pattern, out)
    assert m, f"Export filename not found in output:\n{out}"

    # And the file exists
    export_path = sample_history[0].parent / m.group(0).split("/")[-1]
    assert export_path.exists()
