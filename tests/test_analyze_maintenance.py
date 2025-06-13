import os
import pandas as pd
import numpy as np
import pytest
from pathlib import Path
import analyze_maintenance

@pytest.fixture
def tmp_cwd(tmp_path, monkeypatch):
    """Sandbox the cwd under tmp_path and ensure data dir exists."""
    monkeypatch.chdir(tmp_path)
    # ensure data/ exists
    (tmp_path / "data").mkdir()
    return tmp_path

# --- Test: no history file present ---

def test_analyze_no_history_file(tmp_cwd, monkeypatch, capsys):
    # Monkey-patch create_maintenance_history_sheet to just create an empty file
    def fake_create(path):
        os.makedirs(Path(path).parent, exist_ok=True)
        df = pd.DataFrame(columns=["dummy"])
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Maintenance History", index=False)
        return True
    monkeypatch.setattr(analyze_maintenance, "create_maintenance_history_sheet", fake_create)

    # Remove any pre-existing history
    history_path = tmp_cwd / "data" / "maintenance_history.xlsx"
    if history_path.exists():
        history_path.unlink()

    result = analyze_maintenance.analyze_maintenance(history_path=str(history_path), export=False)
    captured = capsys.readouterr().out

    assert result is None
    assert "No maintenance history found" in captured
    assert "Creating empty maintenance history file" in captured
    # And the file should now exist
    assert history_path.exists()

# --- Test: existing but empty history sheet ---

def test_analyze_empty_history(tmp_cwd, capsys):
    # create an empty history file (zero rows)
    history_path = tmp_cwd / "data" / "maintenance_history.xlsx"
    df_empty = pd.DataFrame(columns=["asset_id", "date", "cost", "action_taken"])
    with pd.ExcelWriter(history_path, engine="openpyxl") as w:
        df_empty.to_excel(w, sheet_name="Maintenance History", index=False)

    result = analyze_maintenance.analyze_maintenance(history_path=str(history_path), export=False)
    captured = capsys.readouterr().out

    assert result is None
    assert "Maintenance history is empty" in captured

# --- Test: correct analysis without export ---

def test_analyze_with_records_no_export(tmp_cwd, capsys):
    history_path = tmp_cwd / "data" / "maintenance_history.xlsx"
    # Build a small DataFrame with two assets
    data = [
        # asset A: two dates 2025-01-01 and 2025-01-11, costs 100 and 200, actions mix
        {"asset_id": "A", "date": "2025-01-01", "cost": 100.0, "action_taken": "Inspection"},
        {"asset_id": "A", "date": "2025-01-11", "cost": 200.0, "action_taken": "Repair"},
        # asset B: single record on 2025-03-01, cost 50, Replacement
        {"asset_id": "B", "date": "2025-03-01", "cost": 50.0, "action_taken": "Replacement"},
    ]
    df = pd.DataFrame(data)
    with pd.ExcelWriter(history_path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Maintenance History", index=False)

    # Run analysis without exporting
    results = analyze_maintenance.analyze_maintenance(history_path=str(history_path), export=False)
    out = capsys.readouterr().out

    # Should load 3 records
    assert "Loaded 3 maintenance records for analysis." in out

    # Validate returned DataFrame
    # Expect two rows, sorted by frequency descending: A then B
    assert isinstance(results, pd.DataFrame)
    assert list(results["asset_id"]) == ["A", "B"]

    # For asset A:
    row_a = results[results["asset_id"] == "A"].iloc[0]
    # record_count = 2
    assert row_a["record_count"] == 2
    # time_span = 10 days
    assert row_a["time_span_days"] == 10
    # avg_interval = 10
    assert pytest.approx(row_a["avg_interval_days"], rel=1e-6) == 10
    # frequency = (2/10)*365 = 73.0
    assert pytest.approx(row_a["maintenance_frequency"], rel=1e-3) == (2/10)*365
    # total_cost = 300
    assert row_a["total_cost"] == 300
    # inspections = 1, repairs = 1, replacements = 0
    assert row_a["inspections"] == 1
    assert row_a["repairs"] == 1
    assert row_a["replacements"] == 0

    # For asset B:
    row_b = results[results["asset_id"] == "B"].iloc[0]
    # record_count = 1, time_span = 0, frequency = NaN
    assert row_b["record_count"] == 1
    assert row_b["time_span_days"] == 0
    assert np.isnan(row_b["maintenance_frequency"])
    assert row_b["total_cost"] == 50
    assert row_b["inspections"] == 0
    assert row_b["repairs"] == 0
    assert row_b["replacements"] == 1

# --- Test: export adds sheet ---

def test_analyze_export_appends_sheet(tmp_cwd):
    history_path = tmp_cwd / "data" / "maintenance_history.xlsx"
    # minimal non-empty history
    df = pd.DataFrame([
        {"asset_id": "X", "date": "2025-05-01", "cost": 10.0, "action_taken": "Inspection"}
    ])
    with pd.ExcelWriter(history_path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Maintenance History", index=False)

    # Run with export=True
    results = analyze_maintenance.analyze_maintenance(history_path=str(history_path), export=True)
    assert isinstance(results, pd.DataFrame)

    # Reload file and check sheet names
    xl = pd.ExcelFile(history_path)
    assert "Maintenance Analysis" in xl.sheet_names

    # And the analysis sheet has at least the expected columns
    analysis_df = xl.parse("Maintenance Analysis")
    expected_cols = [
        "asset_id","record_count","first_maintenance","last_maintenance",
        "time_span_days","avg_interval_days","min_interval_days",
        "max_interval_days","maintenance_frequency","total_cost",
        "avg_cost_per_maintenance","inspections","repairs","replacements"
    ]
    assert list(analysis_df.columns) == expected_cols
