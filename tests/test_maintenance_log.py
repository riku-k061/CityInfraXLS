import os
import json
import uuid
import datetime
import pandas as pd
import pytest
from pathlib import Path
import maintenance_log

# --- Fixtures & Helpers ---

@pytest.fixture
def tmp_cwd(tmp_path, monkeypatch):
    """Helper to sandbox the current working directory."""
    monkeypatch.chdir(tmp_path)
    return tmp_path

@pytest.fixture
def sample_schema(tmp_path):
    schema = {
        "properties": {
            "asset_id": {"type": "string"},
            "action_taken": {"type": "string", "enum": ["Inspect", "Repair"]},
            "performed_by": {"type": "string"},
            "cost": {"type": "number"},
            "date": {"type": "string", "format": "date"},
            "notes": {"type": "string"}
        },
        "required": ["asset_id", "action_taken", "performed_by", "date"]
    }
    p = tmp_path / "maintenance_schema.json"
    p.write_text(json.dumps(schema))
    return schema

# --- load_schema tests ---

def test_load_schema_success(tmp_cwd, sample_schema):
    # writes sample_schema to maintenance_schema.json in cwd
    loaded = maintenance_log.load_schema()
    assert loaded == sample_schema

def test_load_schema_missing(tmp_cwd):
    # no file present
    result = maintenance_log.load_schema()
    assert result is None

def test_load_schema_invalid_json(tmp_cwd):
    # write a broken JSON
    p = tmp_cwd / "maintenance_schema.json"
    p.write_text("{ not: valid json }")
    result = maintenance_log.load_schema()
    assert result is None

# --- validate_input tests ---

def test_validate_input_empty_required(monkeypatch):
    # set a minimal schema with required
    maintenance_log.schema = {"required": ["foo"]}
    ok, err = maintenance_log.validate_input("", "foo", {"type": "string"})
    assert not ok
    assert "foo is required" in err

def test_validate_input_empty_optional(monkeypatch):
    maintenance_log.schema = {"required": []}
    ok, err = maintenance_log.validate_input("", "bar", {"type": "string"})
    assert ok and err is None

def test_validate_input_enum():
    schema = {"type": "string", "enum": ["A", "B"]}
    ok, err = maintenance_log.validate_input("C", "field", schema)
    assert not ok
    assert "not a valid option" in err

def test_validate_input_date():
    schema = {"type": "string", "format": "date"}
    ok, err = maintenance_log.validate_input("2025-06-14", "d", schema)
    assert ok
    # bad format
    ok2, err2 = maintenance_log.validate_input("14-06-2025", "d", schema)
    assert not ok2
    assert "YYYY-MM-DD" in err2

def test_validate_input_number():
    schema = {"type": "number"}
    ok, _ = maintenance_log.validate_input("3.14", "n", schema)
    assert ok
    ok2, err2 = maintenance_log.validate_input("abc", "n", schema)
    assert not ok2
    assert "must be a number" in err2

# --- get_validated_input tests ---

def test_get_validated_input_retries(monkeypatch, capsys):
    # supply two inputs: first invalid number, then valid
    prompts = []
    def fake_input(prompt):
        prompts.append(prompt)
        return "oops" if len(prompts)==1 else "42"
    monkeypatch.setattr('builtins.input', fake_input)
    maintenance_log.schema = {"required": []}
    # field_schema expects number
    val = maintenance_log.get_validated_input("Enter:", "n", {"type": "number"})
    assert val == 42.0
    # ensure it printed an error once
    captured = capsys.readouterr()
    assert "must be a number" in captured.out

# --- log_maintenance tests ---

def test_log_maintenance_no_schema(tmp_cwd, capsys):
    # no schema file -> early return False
    res = maintenance_log.log_maintenance()
    assert res is False
    out = capsys.readouterr().out
    assert "Cannot log maintenance without schema" in out

def test_log_maintenance_success(tmp_cwd, sample_schema, monkeypatch, capsys):
    # prepare schema
    # patch create_maintenance_history_sheet to always succeed
    created_path = tmp_cwd / "data" / "maintenance_history.xlsx"
    def fake_create(path):
        # make parent dir
        os.makedirs(Path(path).parent, exist_ok=True)
        # create an empty sheet with headers
        df = pd.DataFrame(columns=list(sample_schema["properties"].keys()))
        with pd.ExcelWriter(path, engine='openpyxl') as w:
            df.to_excel(w, sheet_name="Maintenance History", index=False)
        return True
    monkeypatch.setattr(maintenance_log, 'create_maintenance_history_sheet', fake_create)

    # fix uuid for predictability
    monkeypatch.setattr(uuid, 'uuid4', lambda: uuid.UUID(int=0x1234))
    # craft a sequence of inputs:
    # asset_id -> "A1"
    # action choice -> "2"  (which is "Repair")
    # performed_by -> "TechX"
    # cost -> "99.99"
    # date -> "2025-06-14"
    # notes -> "All good"
    answers = iter(["A1", "2", "TechX", "99.99", "2025-06-14", "All good"])
    monkeypatch.setattr('builtins.input', lambda prompt="": next(answers))

    # run
    res = maintenance_log.log_maintenance()
    assert res is True

    # verify printed record ID
    out = capsys.readouterr().out
    assert "Maintenance record logged successfully with ID:" in out

    # read back the file and check contents
    df = pd.read_excel(created_path, sheet_name="Maintenance History")
    assert df.shape[0] == 1
    row = df.iloc[0]
    # expected columns in order
    expected_cols = list(sample_schema["properties"].keys())
    assert list(df.columns) == expected_cols
    # values
    assert row["asset_id"] == "A1"
    assert row["action_taken"] == "Repair"
    assert row["performed_by"] == "TechX"
    assert pytest.approx(row["cost"], 0.01) == 99.99
    assert row["date"] == "2025-06-14"
    assert row["notes"] == "All good"
