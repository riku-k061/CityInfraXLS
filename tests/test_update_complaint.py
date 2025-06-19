# tests/test_update_complaint.py

import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

import update_complaint

# --- SANDBOX CWD & DATA DIR ---

@pytest.fixture(autouse=True)
def tmp_cwd(tmp_path, monkeypatch):
    """
    Run each test in its own temp dir with a data/ subfolder.
    """
    monkeypatch.chdir(tmp_path)
    (tmp_path / "data").mkdir()
    return tmp_path

# --- HELPERS ---

def write_schema(schema: dict):
    """Helper to write complaint_schema.json"""
    Path("complaint_schema.json").write_text(json.dumps(schema))

# --- TESTS ---

def test_no_file(tmp_cwd, monkeypatch, capsys):
    """
    If create_complaint_sheet does nothing and complaints.xlsx is absent,
    update_complaint should print a file‐not‐found error and return False.
    """
    monkeypatch.setattr(update_complaint, "create_complaint_sheet", lambda p: None)

    result = update_complaint.update_complaint("any-id", status="Open")
    out = capsys.readouterr().out

    assert result is False
    assert "Error: Complaints file data/complaints.xlsx not found" in out

def test_missing_schema(tmp_cwd, monkeypatch, capsys):
    """
    If complaints.xlsx exists but complaint_schema.json is missing or invalid,
    load_schema will exit(1) with an error message.
    """
    # Stub out sheet creation to produce a valid Excel
    def stub_sheet(path):
        df = pd.DataFrame({"complaint_id": ["X"], "status": ["Open"]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, index=False)
    monkeypatch.setattr(update_complaint, "create_complaint_sheet", stub_sheet)

    with pytest.raises(SystemExit) as exc:
        update_complaint.update_complaint("X", status="Open")
    out = capsys.readouterr().out

    assert exc.value.code == 1
    assert "Error loading schema" in out

def test_invalid_status(tmp_cwd, monkeypatch, capsys):
    """
    If the status passed is not in the schema enum, update_complaint
    should print an invalid‐status error and return False.
    """
    # Write a minimal schema allowing only Open/Closed
    write_schema({
        "properties": {
            "status": {"enum": ["Open", "Closed"]}
        }
    })

    # Stub sheet creation with one existing complaint
    def stub_sheet(path):
        df = pd.DataFrame({"complaint_id": ["id1"], "status": ["Open"]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, index=False)
    monkeypatch.setattr(update_complaint, "create_complaint_sheet", stub_sheet)

    result = update_complaint.update_complaint("id1", status="BadStatus")
    out = capsys.readouterr().out

    assert result is False
    assert "Error: Invalid status. Must be one of: Open, Closed" in out

def test_id_not_found(tmp_cwd, monkeypatch, capsys):
    """
    If the complaint_id does not exist in the sheet, should print
    a not‐found error and return False.
    """
    write_schema({
        "properties": {
            "status": {"enum": ["Open", "Closed"]}
        }
    })

    def stub_sheet(path):
        df = pd.DataFrame({"complaint_id": ["exists"], "status": ["Open"]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, index=False)
    monkeypatch.setattr(update_complaint, "create_complaint_sheet", stub_sheet)

    result = update_complaint.update_complaint("missing", status="Closed")
    out = capsys.readouterr().out

    assert result is False
    assert "Error: Complaint with ID missing not found" in out

def test_append_note(monkeypatch, tmp_cwd, capsys):
    """
    When only a note is supplied, update_complaint should prepend a
    timestamp and write it into 'Resolution Notes'.
    """
    # Freeze datetime.now()
    fake_now = datetime(2025, 6, 17, 12, 0)
    class DummyDT(datetime):
        @classmethod
        def now(cls):
            return fake_now
    monkeypatch.setattr(update_complaint, "datetime", DummyDT)

    write_schema({
        "properties": {
            "status": {"enum": ["Open", "Closed"]}
        }
    })

    def stub_sheet(path):
        # Start with no resolution_notes column
        df = pd.DataFrame({"complaint_id": ["id1"], "status": ["Open"]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, index=False)
    monkeypatch.setattr(update_complaint, "create_complaint_sheet", stub_sheet)

    result = update_complaint.update_complaint("id1", note="First note")
    out = capsys.readouterr().out

    assert result is True
    assert "Complaint id1 updated successfully" in out

    # Read back and check
    df2 = pd.read_excel(tmp_cwd / "data" / "complaints.xlsx")
    note_cell = df2["Resolution Notes"].iloc[0]
    expected = f"[{fake_now.strftime('%Y-%m-%d %H:%M')}] First note"
    assert note_cell == expected
    # Closed At should still be NaN
    assert pd.isna(df2["Closed At"].iloc[0])

def test_status_change_to_closed(monkeypatch, tmp_cwd, capsys):
    """
    Updating status from Open → Closed should set 'Closed At' to now().
    """
    fake_now = datetime(2025, 6, 17, 15, 30)
    class DummyDT(datetime):
        @classmethod
        def now(cls):
            return fake_now
    monkeypatch.setattr(update_complaint, "datetime", DummyDT)

    write_schema({
        "properties": {
            "status": {"enum": ["Open", "Closed"]}
        }
    })

    def stub_sheet(path):
        df = pd.DataFrame({"complaint_id": ["id1"], "status": ["Open"]})
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, index=False)
    monkeypatch.setattr(update_complaint, "create_complaint_sheet", stub_sheet)

    result = update_complaint.update_complaint("id1", status="Closed")
    out = capsys.readouterr().out

    assert result is True
    assert "Complaint id1 updated successfully" in out

    df2 = pd.read_excel(tmp_cwd / "data" / "complaints.xlsx")
    assert df2["Status"].iloc[0] == "Closed"
    # Check Closed At matches fake_now
    closed_val = df2["Closed At"].iloc[0]
    assert pd.to_datetime(closed_val).to_pydatetime() == fake_now

def test_status_change_from_closed(monkeypatch, tmp_cwd, capsys):
    """
    Updating status from Closed → Open should clear 'Closed At'.
    """
    fake_now = datetime(2025, 6, 17, 16, 45)
    class DummyDT(datetime):
        @classmethod
        def now(cls):
            return fake_now
    monkeypatch.setattr(update_complaint, "datetime", DummyDT)

    write_schema({
        "properties": {
            "status": {"enum": ["Open", "Closed"]}
        }
    })

    def stub_sheet(path):
        df = pd.DataFrame({
            "complaint_id": ["id1"],
            "status": ["Closed"],
            # pre-existing Closed At
            "closed_at": [fake_now]
        })
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df.to_excel(w, index=False)
    monkeypatch.setattr(update_complaint, "create_complaint_sheet", stub_sheet)

    result = update_complaint.update_complaint("id1", status="Open")
    out = capsys.readouterr().out

    assert result is True
    assert "Complaint id1 updated successfully" in out

    df2 = pd.read_excel(tmp_cwd / "data" / "complaints.xlsx")
    assert df2["Status"].iloc[0] == "Open"
    # Closed At should now be cleared
    assert pd.isna(df2["Closed At"].iloc[0])
