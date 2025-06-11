import os
import json
import pytest
from openpyxl import Workbook
from utils.excel_handler import load_workbook, save_workbook, init_workbook, create_sheets_from_schema

@pytest.fixture
def sample_headers():
    """Sample headers for testing"""
    return ["ID", "Name", "Location", "Status", "Last Updated"]

@pytest.fixture
def sample_schema():
    """Sample schema for testing sheet creation"""
    return {
        "Road": ["ID", "Name", "Location", "Length", "Width", "Surface Type", "Condition", "Installation Date"],
        "Bridge": ["ID", "Name", "Location", "Length", "Width", "Material", "Condition", "Installation Date"],
        "Park": ["ID", "Name", "Location", "Area", "Facilities", "Condition", "Installation Date"]
    }

@pytest.fixture
def schema_file(tmp_path, sample_schema):
    """Create a temporary schema file"""
    schema_path = tmp_path / "test_schema.json"
    with open(schema_path, "w") as f:
        json.dump(sample_schema, f)
    return schema_path

def test_init_workbook_creates_new(tmp_path, sample_headers):
    """Test init_workbook creates a new workbook with correct headers"""
    # Setup
    test_path = tmp_path / "test_new.xlsx"
    
    # Execute
    wb = init_workbook(test_path, sample_headers)
    
    # Verify
    assert os.path.exists(test_path)
    
    # Check headers
    ws = wb.active
    for col_idx, header in enumerate(sample_headers, start=1):
        assert ws.cell(row=1, column=col_idx).value == header
    
    # Close workbook
    wb.close()

def test_init_workbook_reuses_existing(tmp_path, sample_headers):
    """Test init_workbook reuses an existing workbook"""
    # Setup - create initial workbook
    test_path = tmp_path / "test_reuse.xlsx"
    wb1 = init_workbook(test_path, sample_headers)
    
    # Add some data to distinguish it
    ws1 = wb1.active
    ws1.cell(row=2, column=1, value="TEST_DATA")
    wb1.save(test_path)
    wb1.close()
    
    # Execute - call init_workbook again
    wb2 = init_workbook(test_path, sample_headers)
    
    # Verify it's the same file with our test data
    ws2 = wb2.active
    assert ws2.cell(row=2, column=1).value == "TEST_DATA"
    
    # Close workbook
    wb2.close()

def test_load_workbook_existing(tmp_path):
    """Test loading an existing workbook"""
    # Setup - create a test workbook
    test_path = tmp_path / "test_load.xlsx"
    wb_original = Workbook()
    ws = wb_original.active
    ws.cell(row=1, column=1, value="TEST_HEADER")
    ws.cell(row=2, column=1, value="TEST_VALUE")
    wb_original.save(test_path)
    wb_original.close()
    
    # Execute
    wb_loaded = load_workbook(test_path)
    
    # Verify
    ws_loaded = wb_loaded.active
    assert ws_loaded.cell(row=1, column=1).value == "TEST_HEADER"
    assert ws_loaded.cell(row=2, column=1).value == "TEST_VALUE"
    
    # Close workbook
    wb_loaded.close()

def test_load_workbook_missing(tmp_path):
    """Test load_workbook raises when file is missing"""
    # Setup
    nonexistent_path = tmp_path / "nonexistent.xlsx"
    
    # Execute and verify
    with pytest.raises(Exception):
        load_workbook(nonexistent_path)

def test_save_workbook_roundtrip(tmp_path):
    """Test save_workbook by round-tripping a workbook"""
    # Setup
    test_path = tmp_path / "test_save.xlsx"
    wb_original = Workbook()
    ws = wb_original.active
    ws.cell(row=1, column=1, value="ORIGINAL")
    
    # Execute - save the workbook
    save_workbook(wb_original, test_path)
    wb_original.close()
    
    # Load it back
    wb_loaded = load_workbook(test_path)
    ws_loaded = wb_loaded.active
    
    # Verify original data
    assert ws_loaded.cell(row=1, column=1).value == "ORIGINAL"
    
    # Modify and save again
    ws_loaded.cell(row=1, column=1, value="MODIFIED")
    ws_loaded.cell(row=2, column=1, value="NEW_ROW")
    save_workbook(wb_loaded, test_path)
    wb_loaded.close()
    
    # Load again and verify changes
    wb_final = load_workbook(test_path)
    ws_final = wb_final.active
    assert ws_final.cell(row=1, column=1).value == "MODIFIED"
    assert ws_final.cell(row=2, column=1).value == "NEW_ROW"
    
    # Close workbook
    wb_final.close()

def test_create_sheets_from_schema_new(tmp_path, schema_file, sample_schema):
    """Test create_sheets_from_schema creating a new workbook"""
    # Setup
    test_path = tmp_path / "test_schema_new.xlsx"
    
    # Execute
    wb = create_sheets_from_schema(schema_file, test_path)
    
    # Verify
    try:
        # Check that all sheets from schema exist
        for sheet_name, expected_headers in sample_schema.items():
            assert sheet_name in wb.sheetnames
            
            # Check headers in each sheet
            ws = wb[sheet_name]
            for col_idx, header in enumerate(expected_headers, start=1):
                assert ws.cell(row=1, column=col_idx).value == header
    finally:
        # Close workbook
        wb.close()

def test_create_sheets_from_schema_update(tmp_path, schema_file, sample_schema):
    """Test create_sheets_from_schema updating an existing workbook"""
    # Setup - create initial workbook with partial schema
    test_path = tmp_path / "test_schema_update.xlsx"
    
    # Create a partial workbook with just Road sheet
    wb_initial = Workbook()
    ws = wb_initial.active
    ws.title = "Road"
    
    # Add headers (intentionally different from schema)
    initial_headers = ["ID", "Name", "Old_Field"]
    for col_idx, header in enumerate(initial_headers, start=1):
        ws.cell(row=1, column=col_idx, value=header)
    
    # Add some data
    ws.cell(row=2, column=1, value="ROAD-001")
    ws.cell(row=2, column=2, value="Main Street")
    
    # Save initial workbook
    wb_initial.save(test_path)
    wb_initial.close()
    
    # Execute - update with full schema
    wb_updated = create_sheets_from_schema(schema_file, test_path)
    
    try:
        # Verify
        # 1. All sheets from schema should exist
        for sheet_name in sample_schema.keys():
            assert sheet_name in wb_updated.sheetnames
        
        # 2. Headers in Road sheet should be updated to match schema
        road_ws = wb_updated["Road"]
        for col_idx, header in enumerate(sample_schema["Road"], start=1):
            assert road_ws.cell(row=1, column=col_idx).value == header
        
        # 3. Data in first sheet should still be there
        assert road_ws.cell(row=2, column=1).value == "ROAD-001"
        assert road_ws.cell(row=2, column=2).value == "Main Street"
    finally:
        # Close workbook
        wb_updated.close()

def test_create_sheets_from_schema_preserves_existing_data(tmp_path, schema_file, sample_schema):
    """Test that create_sheets_from_schema preserves existing data when updating headers"""
    # Setup - create initial workbook with data
    test_path = tmp_path / "test_preserve_data.xlsx"
    
    # Create initial workbook with Bridge sheet and some data
    wb_initial = Workbook()
    ws = wb_initial.active
    ws.title = "Bridge"
    
    # Add some headers (subset of schema headers)
    initial_headers = ["ID", "Name", "Location"]
    for col_idx, header in enumerate(initial_headers, start=1):
        ws.cell(row=1, column=col_idx, value=header)
    
    # Add multiple rows of data
    test_data = [
        ["B-001", "North Bridge", "River Crossing"],
        ["B-002", "South Bridge", "Highway 101"]
    ]
    
    for row_idx, row_data in enumerate(test_data, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    # Save initial workbook
    wb_initial.save(test_path)
    wb_initial.close()
    
    # Execute - update with full schema
    wb_updated = create_sheets_from_schema(schema_file, test_path)
    
    try:
        # Verify
        bridge_ws = wb_updated["Bridge"]
        
        # Check new headers are in place
        for col_idx, header in enumerate(sample_schema["Bridge"], start=1):
            assert bridge_ws.cell(row=1, column=col_idx).value == header
        
        # Verify original data is preserved
        for row_idx, row_data in enumerate(test_data, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                assert bridge_ws.cell(row=row_idx, column=col_idx).value == value
    finally:
        # Close workbook
        wb_updated.close()