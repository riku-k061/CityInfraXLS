# tests/test_query_complaints.py

import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytz
import pytest

import query_complaints

# --- FIXTURE: SANDBOX CWD & DATA DIR ---

@pytest.fixture(autouse=True)
def tmp_cwd(tmp_path, monkeypatch):
    """
    Sandbox the working directory under tmp_path and ensure data/ exists.
    """
    monkeypatch.chdir(tmp_path)
    (tmp_path / "data").mkdir()
    return tmp_path

# --- HELPER TO WRITE EXCEL ---

def write_complaints_file(path: Path, df: pd.DataFrame):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Complaints", index=False)

# --- TESTS ---

def test_no_file(tmp_cwd, capsys):
    result = query_complaints.query_complaints()
    out = capsys.readouterr().out
    assert result is None
    assert "Error: Complaints file not found" in out

def test_empty_sheet(tmp_cwd, capsys):
    path = tmp_cwd / "data" / "complaints.xlsx"
    write_complaints_file(path, pd.DataFrame())

    result = query_complaints.query_complaints()
    out = capsys.readouterr().out

    assert isinstance(result, pd.DataFrame)
    assert result.empty
    assert "No complaints found in the database." in out

@pytest.fixture
def sample_complaints(tmp_cwd):
    """
    Create a sample complaints.xlsx with three records.
    Use timezone-naive datetime strings so filtering works.
    """
    data = [
        {
            "complaint_id": "id1",
            "reporter": "Alice",
            "asset_location": "Loc1",
            "department": "Electrical",
            "status": "Open",
            "rating": 3,
            # naive datetime string
            "created_at": "2025-06-01 10:00:00",
            "closed_at": None
        },
        {
            "complaint_id": "id2",
            "reporter": "Bob",
            "asset_location": "Loc2",
            "department": "Water",
            "status": "Closed",
            "rating": 5,
            "created_at": "2025-05-20 09:30:00",
            "closed_at": "2025-06-02 15:00:00"
        },
        {
            "complaint_id": "id3",
            "reporter": "Carol",
            "asset_location": "Loc3",
            "department": "Road",
            "status": "In Progress",
            "rating": 2,
            "created_at": "2025-06-10 08:45:00",
            "closed_at": None
        }
    ]
    df = pd.DataFrame(data)
    path = tmp_cwd / "data" / "complaints.xlsx"
    write_complaints_file(path, df)
    return path, df

def test_query_all(sample_complaints, capsys):
    result = query_complaints.query_complaints()
    out = capsys.readouterr().out

    assert isinstance(result, pd.DataFrame)
    assert len(result) == 3
    assert "No filters applied - showing all complaints" in out
    assert "Found 3 complaint(s):" in out
    assert "+" in out  # basic check for tabulate grid

def test_query_with_filters(sample_complaints, capsys):
    result = query_complaints.query_complaints(
        status="Open",
        department="Electrical",
        min_rating=3,
        date_from="2025-05-01",
        date_to="2025-06-05",
        export=False
    )
    out = capsys.readouterr().out

    # Now only id1 matches
    assert isinstance(result, pd.DataFrame)
    assert len(result) == 1
    assert result.iloc[0]["complaint_id"] == "id1"

    # Printed filters
    assert "Filters applied:" in out
    assert "- Status: Open" in out
    assert "- Department: Electrical" in out
    assert "- Min Rating: 3" in out
    assert "- From: 2025-05-01" in out
    assert "- To: 2025-06-05" in out

def test_query_no_matches(sample_complaints, capsys):
    result = query_complaints.query_complaints(
        status="Closed",
        department="Electrical"
    )
    out = capsys.readouterr().out

    assert isinstance(result, pd.DataFrame)
    assert result.empty
    assert "No complaints match the specified criteria." in out

def test_query_export_creates_file(sample_complaints, monkeypatch, capsys, tmp_cwd):
    # Freeze datetime.now to a known UTC timestamp
    fake_now = datetime(2025, 6, 17, 14, 30, 45, tzinfo=pytz.UTC)
    class DummyDateTime(datetime):
        @classmethod
        def now(cls, tz=None):
            return fake_now

    monkeypatch.setattr(query_complaints, "datetime", DummyDateTime)

    result = query_complaints.query_complaints(export=True)
    out = capsys.readouterr().out

    assert isinstance(result, pd.DataFrame)
    assert len(result) == 3

    expected = f"data/complaints_query_{fake_now.strftime('%Y%m%d%H%M%S')}.xlsx"
    assert re.search(re.escape(expected), out)

    assert (tmp_cwd / expected).exists()
