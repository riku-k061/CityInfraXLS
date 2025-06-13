import os
import sys
import pandas as pd
import pytest
from datetime import datetime

# Make sure the project root is on the import path
sys.path.insert(0, os.getcwd())

import assign_task
import update_task
import delete_task

@pytest.fixture(autouse=True)
def tmp_cwd(monkeypatch, tmp_path):
    """
    Switch CWD to a fresh temp directory and create a `data/` subfolder.
    All scripts under test use relative paths into `data/`.
    """
    monkeypatch.chdir(tmp_path)
    (tmp_path / "data").mkdir()
    return tmp_path

def create_incidents_file(tmp_path):
    """Helper: write a single open incident to data/incidents.xlsx"""
    df = pd.DataFrame([{
        "Incident ID": "INC-123",
        "Severity": "High issue",
        "Status": "Open"
    }])
    df.to_excel(tmp_path / "data" / "incidents.xlsx", index=False)

def create_contractors_file(tmp_path):
    """Helper: write a single contractor to data/contractors.xlsx"""
    df = pd.DataFrame([{
        "contractor_id": "CTR-456",
        "name": "Alice",
        "specialties": [["Electrical"]],
        "rating": [4.5]
    }])
    # Note: pandas will write lists as strings; the script only checks for the ID
    df.to_excel(tmp_path / "data" / "contractors.xlsx", index=False)

def test_assign_task_creates_and_appends(tmp_path):
    # prepare prerequisites
    create_incidents_file(tmp_path)
    create_contractors_file(tmp_path)

    # call assign_task; it will create data/tasks.xlsx
    new_task = assign_task.assign_task("INC-123", "CTR-456", details="Fix wiring")

    # verify return value
    assert new_task["Incident ID"] == "INC-123"
    assert new_task["Contractor ID"] == "CTR-456"
    assert new_task["Status"] == "Assigned"
    assert "Task ID" in new_task

    # verify file content
    tasks_df = pd.read_excel(tmp_path / "data" / "tasks.xlsx")
    assert len(tasks_df) == 1
    row = tasks_df.iloc[0]
    assert row["Incident ID"] == "INC-123"
    assert row["Contractor ID"] == "CTR-456"
    assert row["Status"] == "Assigned"
    assert row["Details"] == "Fix wiring"

def test_update_task_changes_status_and_appends_note(tmp_path):
    # create an initial tasks.xlsx
    df = pd.DataFrame([{
        "Task ID": "TASK-001",
        "Incident ID": "INC-123",
        "Contractor ID": "CTR-456",
        "Assigned At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Status": "Assigned",
        "Status Updated At": "",
        "Details": ""
    }])
    df.to_excel(tmp_path / "data" / "tasks.xlsx", index=False)

    # perform the update
    success = update_task.update_task("TASK-001", "Completed", change_note="All done")
    assert success

    # reload and inspect
    updated = pd.read_excel(tmp_path / "data" / "tasks.xlsx")
    assert updated.at[0, "Status"] == "Completed"

    details = updated.at[0, "Details"]
    assert "Status changed from 'Assigned' to 'Completed'" in details
    assert "All done" in details
    assert pd.notna(updated.at[0, "Status Updated At"])

def test_delete_task_removes_entry_and_creates_backup(tmp_path):
    # create a tasks.xlsx with two entries
    df = pd.DataFrame([
        {"Task ID": "TASK-001", "Incident ID": "INC-1", "Contractor ID": "C1", "Assigned At": "", "Status": "Assigned", "Status Updated At": "", "Details": ""},
        {"Task ID": "TASK-002", "Incident ID": "INC-2", "Contractor ID": "C2", "Assigned At": "", "Status": "Assigned", "Status Updated At": "", "Details": ""},
    ])
    df.to_excel(tmp_path / "data" / "tasks.xlsx", index=False)

    # delete the first task (force skips prompt)
    success = delete_task.delete_task("TASK-001", force=True)
    assert success

    # backup should have been created under data/backups/
    backups = list((tmp_path / "data" / "backups").iterdir())
    assert len(backups) == 1
    assert backups[0].name.startswith("tasks_")

    # the remaining file should only contain the second task
    remaining = pd.read_excel(tmp_path / "data" / "tasks.xlsx")
    assert len(remaining) == 1
    assert remaining.at[0, "Task ID"] == "TASK-002"
