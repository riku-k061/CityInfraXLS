# tests/test_condition_scoring.py

import json
import pandas as pd
import pytest
from pathlib import Path

import utils.excel_handler  # ensure in path
from utils.excel_handler import load_workbook  # unused but forces import

from utils.excel_handler import Workbook  # fix loading

# assume your functions are in the top‐level module:
import utils.excel_handler as eh  
# If you placed calculate_condition_scores etc in a separate module,
# adjust the import accordingly, e.g.:
# from condition_scoring import calculate_condition_scores, get_risk_color, hex_to_rgb

from utils.excel_handler import (
    calculate_condition_scores,
    get_risk_color,
    hex_to_rgb
)

# -- get_risk_color & hex_to_rgb --

def test_get_risk_color_known():
    assert get_risk_color("Excellent") == "#00FF00"
    assert get_risk_color("Unknown") == "#CCCCCC"

def test_hex_to_rgb():
    assert hex_to_rgb("#FF0000") == (255,0,0)
    assert hex_to_rgb("00FF00") == (0,255,0)

# -- calculate_condition_scores --

@pytest.fixture
def scoring_config(tmp_path):
    cfg = {
        "scoring_parameters": {
            "incident_count": {"critical_threshold": 10},
            "inspection_rating": {"scale": {"min": 1, "max": 5}}
        },
        "scoring_thresholds": {
            "excellent": {"min": 90},
            "good": {"min": 75},
            "fair": {"min": 50},
            "poor": {"min": 25},
            "critical": {"min": 0}
        }
    }
    path = tmp_path / "condition_scoring.json"
    path.write_text(json.dumps(cfg))
    return str(path)

@pytest.fixture
def sample_assets(tmp_path):
    df = pd.DataFrame([
        {"ID":"A1","Surface Type":"Road","Location":"L1","inspection_rating":5},
        {"ID":"A2","Surface Type":"Bridge","Location":"L2"}
    ])
    path = tmp_path / "assets.xlsx"
    df.to_excel(str(path), index=False)
    return str(path)

@pytest.fixture
def sample_incidents(tmp_path):
    df = pd.DataFrame([
        {"Asset ID":"A1","Some":"x"},
        {"Asset ID":"A1","Some":"y"},
        {"Asset ID":"A2","Some":"z"}
    ])
    path = tmp_path / "incidents.xlsx"
    df.to_excel(str(path), index=False)
    return str(path)

@pytest.fixture
def sample_tasks(tmp_path):
    # not used in current code
    path = tmp_path / "tasks.xlsx"
    pd.DataFrame().to_excel(str(path), index=False)
    return str(path)

def test_calculate_condition_scores_minimal(scoring_config, sample_assets, sample_incidents, sample_tasks):
    df = calculate_condition_scores(
        assets_path=sample_assets,
        incidents_path=sample_incidents,
        tasks_path=sample_tasks,
        config_path=scoring_config
    )
    # Should have two rows (A1, A2)
    assert set(df['Asset ID']) == {"A1","A2"}
    # A1 has 2 incidents → incident_score=100*(1-2/10)=80
    row1 = df[df['Asset ID']=="A1"].iloc[0]
    assert pytest.approx(row1['Incident Score'], rel=1e-3) == 80.0
    # A1 inspection_rating=5 → score=100*((5-1)/(5-1))=100
    assert pytest.approx(row1['Inspection Score']) == 100.0
    # weight-adjusted = (80*0.5 + 100*0.5)=90
    assert pytest.approx(row1['Weight-Adjusted Score']) == 90.0
    # risk category: excellent (>=90)
    assert row1['Risk Category'] == 'Excellent'

def test_calculate_condition_scores_default_inspection(scoring_config, sample_assets, sample_incidents, sample_tasks):
    # remove inspection_rating column
    # write new assets.xlsx without inspection_rating
    df = pd.read_excel(sample_assets)
    df = df.drop(columns=['inspection_rating'])
    df.to_excel(sample_assets, index=False)
    df_out = calculate_condition_scores(
        assets_path=sample_assets,
        incidents_path=sample_incidents,
        tasks_path=sample_tasks,
        config_path=scoring_config
    )
    # All Inspection Score values should be the default 75
    assert all(df_out['Inspection Score'] == 75)