# utils/excel_handler.py

import os
import json
import logging
import pandas as pd
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

def create_sheets_from_schema(schema_path, output_path, sheet_name=None):
    """
    Create Excel sheets based on a JSON schema.
    
    Args:
        schema_path (str): Path to the JSON schema file
        output_path (str): Path to save the Excel file
        sheet_name (str, optional): Name of the sheet. Defaults to None.
    """
    # Create directory if it doesn't exist
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Load schema
    with open(schema_path, 'r') as f:
        schema = json.load(f)
    
    if not sheet_name:
        # Convert output_path to sheet_name (e.g., data/assets.xlsx -> assets)
        sheet_name = os.path.basename(output_path).split('.')[0]
    
    # Get properties from schema
    properties = schema.get('properties', {})
    
    # Create DataFrame with columns from schema properties
    df = pd.DataFrame(columns=list(properties.keys()))
    
    # Save empty DataFrame to Excel
    df.to_excel(output_path, index=False, sheet_name=sheet_name)
    print(f"Created {output_path} with columns: {', '.join(df.columns)}")

def create_tasks_sheet(output_path="data/tasks.xlsx", sheet_name="tasks"):
    """
    Create a tasks Excel sheet with predefined columns.
    
    Args:
        output_path (str, optional): Path to save the Excel file. Defaults to "data/tasks.xlsx".
        sheet_name (str, optional): Name of the sheet. Defaults to "tasks".
    """
    # Create directory if it doesn't exist
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Define columns for tasks sheet
    columns = ["Task ID", "Incident ID", "Contractor ID", "Assigned At", "Status", "Details"]
    
    # Create DataFrame with defined columns
    df = pd.DataFrame(columns=columns)
    
    # Save empty DataFrame to Excel
    df.to_excel(output_path, index=False, sheet_name=sheet_name)
    print(f"Created {output_path} with columns: {', '.join(columns)}")