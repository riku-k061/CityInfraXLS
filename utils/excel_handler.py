# utils/excel_handler.py

import os
import json
import logging
from openpyxl import Workbook, load_workbook as openpyxl_load

# Configure basic logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),  # Log to console
        logging.FileHandler('cityinfraxls.log')  # Log to file
    ]
)

logger = logging.getLogger('excel_handler')

def load_workbook(path):
    """
    Load an Excel workbook from the given file path.
    
    Args:
        path (str): Path to the Excel file
        
    Returns:
        openpyxl.Workbook: Loaded workbook object
    """
    logger.info(f"Loading workbook from {path}")
    try:
        wb = openpyxl_load(path)
        return wb
    except Exception as e:
        logger.error(f"Failed to load workbook from {path}: {str(e)}")
        raise

def save_workbook(wb, path):
    """
    Save a workbook to the specified path.
    
    Args:
        wb (openpyxl.Workbook): Workbook to save
        path (str): Path where to save the workbook
        
    Returns:
        None
    """
    logger.info(f"Saving workbook to {path}")
    try:
        wb.save(path)
    except Exception as e:
        logger.error(f"Failed to save workbook to {path}: {str(e)}")
        raise

def init_workbook(path, headers):
    """
    Initialize a new workbook with the specified headers if it doesn't exist.
    
    Args:
        path (str): Path where to create/check the workbook
        headers (list): List of header names for the first row
        
    Returns:
        openpyxl.Workbook: The initialized or existing workbook
    """
    if os.path.exists(path):
        logger.info(f"Workbook already exists at {path}, loading existing file")
        return load_workbook(path)
    
    logger.info(f"Creating new workbook at {path} with headers: {headers}")
    wb = Workbook()
    ws = wb.active
    
    # Add headers to the first row
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=header)
    
    # Save the new workbook
    save_workbook(wb, path)
    
    return wb

def create_sheets_from_schema(schema_path, workbook_path):
    """
    Ensure the workbook has all sheets defined in the schema with correct headers.
    
    Args:
        schema_path (str): Path to the JSON schema file
        workbook_path (str): Path to the Excel workbook file
        
    Returns:
        openpyxl.Workbook: The updated workbook
    """
    # Load schema
    logger.info(f"Loading asset schema from {schema_path}")
    try:
        with open(schema_path, 'r') as f:
            schema = json.load(f)
    except Exception as e:
        logger.error(f"Failed to load schema from {schema_path}: {str(e)}")
        raise
    
    # Load or create workbook
    if os.path.exists(workbook_path):
        wb = load_workbook(workbook_path)
    else:
        logger.info(f"Creating new workbook at {workbook_path}")
        wb = Workbook()
        # Remove the default sheet (will add sheets from schema)
        default_sheet = wb.active
        wb.remove(default_sheet)
    
    # Process each asset type in the schema
    for asset_type, headers in schema.items():
        logger.info(f"Processing sheet for asset type: {asset_type}")
        
        # Check if sheet exists
        if asset_type in wb.sheetnames:
            ws = wb[asset_type]
            
            # Check if headers match
            existing_headers = [cell.value for cell in ws[1]]
            
            if existing_headers != headers:
                logger.info(f"Headers for {asset_type} don't match schema, updating headers")
                # Clear the first row and update with schema headers
                for col_idx, header in enumerate(headers, start=1):
                    ws.cell(row=1, column=col_idx, value=header)
        else:
            # Create new sheet with headers
            logger.info(f"Creating new sheet for {asset_type}")
            ws = wb.create_sheet(asset_type)
            
            # Add headers
            for col_idx, header in enumerate(headers, start=1):
                ws.cell(row=1, column=col_idx, value=header)
    
    # Save the updated workbook
    save_workbook(wb, workbook_path)
    
    return wb