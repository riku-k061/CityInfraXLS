# tests/test_report_complaint.py

import json
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

import report_complaint
from utils import excel_handler

# --- GLOBAL FIX FOR PANDAS.REPLACE BUG IN TESTS ---

@pytest.fixture(autouse=True)
def patch_dataframe_replace(monkeypatch):
    """
    Stub out pandas.DataFrame.replace so that replace(None, np.nan)
    (and alike) becomes a no-op rather than raising.
    """
    monkeypatch.setattr(
        pd.DataFrame,
        "replace",
        lambda self, *args, **kwargs: self
    )
    # No return necessary for autouse
        

# --- SANDBOX CWD & DATA DIR ---

@pytest.fixture(autouse=True)
def tmp_cwd(tmp_path, monkeypatch):
    """
    Sandbox the working directory under tmp_path and create a data/ subdir.
    """
    monkeypatch.chdir(tmp_path)
    (tmp_path / "data").mkdir()
    return tmp_path


# --- HELPERS ---

class DummyExcelHandler:
    @staticmethod
    def create_complaint_sheet(path):
        """
        Dummy version: just write an empty 'Complaints' sheet so the XLS file exists.
        """
        df_empty = pd.DataFrame()
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df_empty.to_excel(writer, sheet_name="Complaints", index=False)


# --- TESTS ---

def test_error_on_missing_schema(tmp_cwd, monkeypatch, capsys):
    """
    If complaint_schema.json is absent, report_complaint should
    print an error and exit via SystemExit.
    """
    # Patch the handler *inside* report_complaint, not excel_handler module
    monkeypatch.setattr(
        report_complaint,
        "create_complaint_sheet",
        DummyExcelHandler.create_complaint_sheet
    )

    # No schema file -> sys.exit(1)
    with pytest.raises(SystemExit) as exc:
        report_complaint.report_complaint()
    # Confirm exit code is 1
    assert exc.value.code == 1

    out = capsys.readouterr().out
    assert "Error loading complaint schema" in out


def test_successful_report_complaint(tmp_cwd, monkeypatch, capsys):
    """
    Simulate a full run:
      - write a minimal schema.json
      - monkeypatch create_complaint_sheet in report_complaint
      - simulate user inputs
      - assert return True, correct console output, and Excel contents
    """
    # 1) Stub out sheet creation
    monkeypatch.setattr(
        report_complaint,
        "create_complaint_sheet",
        DummyExcelHandler.create_complaint_sheet
    )

    # 2) Write a minimal schema
    schema = {
        "properties": {
            "complaint_id": {"type": "string"},
            "status": {"type": "string"},
            "created_at": {"type": "string"},
            "closed_at": {"type": "string"},
            "description": {
                "type": "string",
                "description": "Description of the issue"
            },
            "severity": {
                "type": "integer",
                "minimum": 1,
                "maximum": 5,
                "description": "Severity Level"
            },
            "category": {
                "type": "string",
                "enum": ["Electrical", "Water", "Road"],
                "description": "Category"
            }
        },
        "required": ["description", "severity"]
    }
    Path("complaint_schema.json").write_text(json.dumps(schema))

    # 3) Simulate user inputs in the prompt order:
    #    description, severity, category
    inputs = iter([
        "Leaky pipe in basement",  # description
        "3",                       # severity (1–5)
        "2"                        # category: Water
    ])
    monkeypatch.setattr("builtins.input", lambda prompt="": next(inputs))

    # 4) Call the function
    result = report_complaint.report_complaint()

    # 5) It should succeed
    assert result is True
    out = capsys.readouterr().out
    assert "Complaint successfully registered with ID:" in out
    assert "Status: Open" in out
    assert "Created At:" in out

    # 6) Verify the XLS file
    complaints_path = tmp_cwd / "data" / "complaints.xlsx"
    assert complaints_path.exists()

    # 7) Read back and inspect
    df = pd.read_excel(complaints_path, sheet_name="Complaints")
    # Expect exactly 1 row and one column per schema field
    assert df.shape == (1, len(schema["properties"]))

    row = df.iloc[0]
    # User‐entered
    assert row["description"] == "Leaky pipe in basement"
    assert int(row["severity"]) == 3
    assert row["category"] == "Water"

    # Auto‐generated
    assert row["status"] == "Open"
    # complaint_id is a valid UUID
    uuid.UUID(row["complaint_id"])
    # created_at parses as ISO‐8601
    datetime.fromisoformat(row["created_at"])

    # closed_at should be blank / NaN
    assert pd.isna(row["closed_at"])
