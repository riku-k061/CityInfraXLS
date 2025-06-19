# utils/excel_handler.py

import os
import json
import logging
import pandas as pd
from openpyxl import Workbook, load_workbook as openpyxl_load
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import datetime
import numpy as np

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

def create_maintenance_history_sheet(path='data/maintenance_history.xlsx'):
    """
    Create a new maintenance history Excel sheet with headers based on the maintenance schema.
    
    Args:
        path (str): Path to the Excel file to create or update
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Load the maintenance schema to get the headers
        with open('maintenance_schema.json', 'r') as schema_file:
            schema = json.load(schema_file)
        
        # Extract the properties to use as headers
        headers = list(schema['properties'].keys())
        
        # Create a new DataFrame with the headers but no data
        df = pd.DataFrame(columns=headers)
        
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(path), exist_ok=True)
        
        # Create a new workbook
        wb = Workbook()
        # Get the active worksheet and rename it
        ws = wb.active
        ws.title = "Maintenance History"
        
        # Add headers to the sheet
        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)
        
        # Save the workbook
        wb.save(path)
        print(f"Successfully created maintenance history sheet at {path}")
        return True
        
    except Exception as e:
        print(f"Error creating maintenance history sheet: {e}")
        return False
    
def create_complaint_sheet(path):
    """
    Creates a complaint sheet at the specified path if it doesn't exist already.
    If the file exists but doesn't have a Complaints sheet, adds the sheet.
    Ensures the sheet has the correct headers based on the complaint schema.
    
    Args:
        path (str): Path to the Excel file
    
    Returns:
        bool: True if created or already exists with correct structure, False otherwise
    """
    # Load complaint schema
    with open('complaint_schema.json', 'r') as schema_file:
        schema = json.load(schema_file)
    
    # Extract headers from schema properties
    headers = list(schema['properties'].keys())
    
    # If file doesn't exist, create new file with headers
    if not os.path.exists(path):
        df = pd.DataFrame(columns=headers)
        df.to_excel(path, sheet_name='Complaints', index=False)
        return True
    
    try:
        # Check if Complaints sheet exists
        wb = load_workbook(path)
        if 'Complaints' not in wb.sheetnames:
            # If file exists but Complaints sheet doesn't, create it
            with pd.ExcelWriter(path, engine='openpyxl', mode='a') as writer:
                df = pd.DataFrame(columns=headers)
                df.to_excel(writer, sheet_name='Complaints', index=False)
        else:
            # If sheet exists, verify headers
            df = pd.read_excel(path, sheet_name='Complaints')
            existing_headers = df.columns.tolist()
            
            # Check if headers match
            if set(existing_headers) != set(headers):
                # Create a backup
                backup_path = path.replace('.xlsx', '_backup.xlsx')
                wb.save(backup_path)
                
                # Recreate with correct headers, preserving data for matching columns
                with pd.ExcelWriter(path, engine='openpyxl', mode='a') as writer:
                    new_df = pd.DataFrame(columns=headers)
                    
                    # Copy data for columns that exist in both
                    for col in set(existing_headers).intersection(set(headers)):
                        if len(df) > 0:
                            new_df[col] = df[col]
                            
                    # Delete the original sheet
                    wb = load_workbook(path)
                    if 'Complaints' in wb.sheetnames:
                        wb.remove(wb['Complaints'])
                        wb.save(path)
                    
                    # Write the corrected sheet
                    new_df.to_excel(writer, sheet_name='Complaints', index=False)
        
        return True
    
    except (InvalidFileException, Exception) as e:
        print(f"Error creating complaint sheet: {str(e)}")
        return False
    
def calculate_condition_scores(assets_path, incidents_path, tasks_path, config_path="condition_scoring.json"):
    """
    Calculate condition scores for all assets in one pass by loading data, computing metrics,
    applying weights from the schema, and returning the complete results dataframe.
    
    Args:
        assets_path (str): Path to the assets Excel file
        incidents_path (str): Path to the incidents Excel file
        tasks_path (str): Path to the maintenance tasks Excel file
        config_path (str): Path to the condition scoring configuration
        
    Returns:
        pandas.DataFrame: Complete condition scores dataframe ready for Excel output
    """
    # Load the condition scoring configuration
    with open(config_path, 'r') as f:
        scoring_config = json.load(f)
    
    # Load all required datasets in a single pass
    assets_df = pd.read_excel(assets_path)
    incidents_df = pd.read_excel(incidents_path)
    
    # Create base dataframe with asset metadata
    condition_df = pd.DataFrame()
    condition_df['Asset ID'] = assets_df['ID']
    condition_df['Asset Name'] = ""
    condition_df['Asset Type'] = assets_df['Surface Type']
    condition_df['Location'] = assets_df['Location']
    
    
    # ---------------------------------------------
    # 3. Calculate Incident Count Score
    # ---------------------------------------------
    # Count incidents per asset
    incident_counts = incidents_df.groupby('Asset ID').size().to_dict()
    
    # Add incident count to condition_df
    condition_df['Incident Count'] = condition_df['Asset ID'].map(incident_counts).fillna(0).astype(int)
    
    # Get the incident threshold from config
    critical_incident_count = scoring_config['scoring_parameters']['incident_count'].get('critical_threshold', 10)
    
    # Normalize incident count to score (fewer incidents = higher score)
    condition_df['Incident Score'] = condition_df['Incident Count'].apply(
        lambda x: max(0, 100 * (1 - x / critical_incident_count))
    )
    
    # ---------------------------------------------
    # 4. Calculate Inspection Score (if available)
    # ---------------------------------------------
    # If we have inspection data, we would process it here
    # For now, use a placeholder or try to extract from assets_df if it exists
    if 'inspection_rating' in assets_df.columns:
        condition_df['Inspection Rating'] = assets_df['inspection_rating']
        
        # Get the scale from config
        min_rating = scoring_config['scoring_parameters']['inspection_rating']['scale']['min']
        max_rating = scoring_config['scoring_parameters']['inspection_rating']['scale']['max']
        scale_range = max_rating - min_rating
        
        # Normalize inspection rating to score
        condition_df['Inspection Score'] = condition_df['Inspection Rating'].apply(
            lambda x: 100 * ((x - min_rating) / scale_range) if pd.notnull(x) else 50
        )
    else:
        # Use a default inspection score
        condition_df['Inspection Score'] = 75  # Reasonably good default
    
    # ---------------------------------------------
    # 5. Calculate Combined Weight-Adjusted Score
    # ---------------------------------------------
    # Apply weights and calculate the final score
    condition_df['Weight-Adjusted Score'] = (
        condition_df['Incident Score'] * 0.5 +
        condition_df['Inspection Score'] * 0.5
    ).round(1)
    
    # ---------------------------------------------
    # 6. Determine Risk Category based on thresholds
    # ---------------------------------------------
    thresholds = scoring_config['scoring_thresholds']
    
    def get_risk_category(score):
        """Map the score to a risk category based on thresholds"""
        if score >= thresholds['excellent']['min']:
            return 'Excellent'
        elif score >= thresholds['good']['min']:
            return 'Good'
        elif score >= thresholds['fair']['min']:
            return 'Fair'
        elif score >= thresholds['poor']['min']:
            return 'Poor'
        else:
            return 'Critical'
    
    condition_df['Risk Category'] = condition_df['Weight-Adjusted Score'].apply(get_risk_category)
    
    
    # Sort by score (worst to best) to prioritize assets needing attention
    condition_df = condition_df.sort_values('Weight-Adjusted Score')
    
    return condition_df
    
def create_condition_scores_sheet(output_path, assets_path, incidents_path, tasks_path, 
                                 config_path="condition_scoring_schema.json", 
                                 export_mapping=True):
    """
    Create a Condition Scores report with asset condition metrics, data gaps tracking,
    and optional GIS mapping export.
    
    Args:
        output_path (str): Path where the condition scores sheet will be saved
        assets_path (str): Path to the assets Excel file
        incidents_path (str): Path to the incidents Excel file
        tasks_path (str): Path to the maintenance tasks Excel file
        config_path (str): Path to the condition scoring config file
        export_mapping (bool): Whether to export GeoJSON mapping data
        
    Returns:
        dict: Paths to the created files
    """
    # Calculate all condition scores and identify data gaps
    condition_df = calculate_condition_scores(
        assets_path, incidents_path, tasks_path, config_path
    )
    
    # Create the directory if it doesn't exist
    output_dir = Path(output_path)
    output_dir.mkdir(exist_ok=True)
    
    # Create the main Excel report as before...
    report_path = output_dir / "asset_condition_scores.xlsx"
    
    # [... existing code for creating the condition scores Excel sheets ...]
    
    # Export GIS mapping data if requested
    mapping_files = {}
    if export_mapping:
        # Load assets data with coordinates
        assets_df = pd.read_excel(assets_path)
        
        # Create GIS mapping exports
        geojson_path, mapping_excel_path = export_risk_mapping(
            condition_df, 
            assets_df,
            output_dir,
            "asset_risk_mapping"
        )
        
        mapping_files = {
            'geojson': geojson_path,
            'excel': mapping_excel_path
        }
    
    # Return paths to all created files
    return {
        'report': str(report_path),
        'mapping': mapping_files
    }

def create_asset_condition_report(data_dir="data", output_dir="reports"):
    """
    Creates a comprehensive asset condition report using the condition scoring schema.
    
    Args:
        data_dir (str): Directory containing the data files
        output_dir (str): Directory where the report will be saved
        
    Returns:
        str: Path to the created report
    """
    # Ensure output directory exists
    Path(output_dir).mkdir(exist_ok=True)
    
    # Define paths
    assets_path = Path(data_dir) / "assets.xlsx"
    incidents_path = Path(data_dir) / "incidents.xlsx"
    tasks_path = Path(data_dir) / "tasks.xlsx"
    config_path = Path("condition_scoring_schema.json")
    output_path = Path(output_dir)
    
    # Create the condition scores sheet
    report_path = create_condition_scores_sheet(
        output_path, 
        assets_path, 
        incidents_path, 
        tasks_path, 
        config_path
    )
    
    print(f"Asset condition report created successfully: {report_path}")
    return report_path

def generate_scheduled_condition_reports(frequency="monthly", data_dir="data", output_dir="reports"):
    """
    Set up scheduled generation of condition score reports
    
    Args:
        frequency (str): Frequency of report generation (daily, weekly, monthly)
        data_dir (str): Directory containing the data files
        output_dir (str): Directory where reports will be saved
        
    Returns:
        None
    """
    # This is a placeholder for scheduling functionality
    # In a real implementation, this would set up a scheduled task
    timestamp = datetime.datetime.now().strftime("%Y%m%d")
    report_dir = Path(output_dir) / f"{frequency}_reports_{timestamp}"
    report_dir.mkdir(exist_ok=True)
    
    # Generate the report
    report_path = create_asset_condition_report(data_dir, report_dir)
    
    print(f"Scheduled {frequency} report generated at: {report_path}")
    return report_path

def add_mapping_to_condition_report(condition_report_path, geojson_path):
    """
    Add a GIS Mapping sheet to an existing condition scores report and embed the GeoJSON data.
    
    Args:
        condition_report_path (str): Path to the existing condition scores Excel report
        geojson_path (str): Path to the GeoJSON file to embed
        
    Returns:
        str: Path to the updated report
    """
    # Load the existing workbook
    wb = load_workbook(condition_report_path)
    
    # Check if a mapping sheet already exists, remove it if it does
    if "GIS Mapping" in wb.sheetnames:
        del wb["GIS Mapping"]
    
    # Create a new mapping sheet
    ws = wb.create_sheet("GIS Mapping")
    
    # Load the GeoJSON data
    with open(geojson_path, 'r') as f:
        geojson_data = json.load(f)
    
    # Extract the features
    features = geojson_data.get('features', [])
    
    # Create a DataFrame for visible data
    visible_data = []
    for feature in features:
        props = feature['properties']
        coords = feature['geometry']['coordinates']
        visible_data.append({
            'Asset ID': props.get('asset_id', ''),
            'Asset Name': props.get('name', ''),
            'Risk Category': props.get('risk_category', ''),
            'Score': props.get('score', 0),
            'Longitude': coords[0],
            'Latitude': coords[1]
        })
    
    visible_df = pd.DataFrame(visible_data)
    
    # Add visible data to the sheet
    for r_idx, row in enumerate(dataframe_to_rows(visible_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            
            # Format header row
            if r_idx == 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Format Risk Category cells
            elif c_idx == visible_df.columns.get_loc('Risk Category') + 1 and r_idx > 1:
                color = get_risk_color(value)
                cell.fill = PatternFill(start_color=color[1:], end_color=color[1:], fill_type="solid")
    
    # Add GeoJSON data in a hidden column
    geojson_str = json.dumps(geojson_data)
    geojson_cell = ws.cell(row=1, column=len(visible_df.columns) + 2, value="GeoJSON_Data")
    geojson_cell.font = Font(bold=True)
    
    # Split the GeoJSON string into chunks
    chunk_size = 30000
    geojson_chunks = [geojson_str[i:i+chunk_size] for i in range(0, len(geojson_str), chunk_size)]
    
    for i, chunk in enumerate(geojson_chunks):
        ws.cell(row=i+2, column=len(visible_df.columns) + 2, value=chunk)
    
    # Hide the GeoJSON data column
    col_letter = get_column_letter(len(visible_df.columns) + 2)
    ws.column_dimensions[col_letter].hidden = True
    
    # Add title and instructions
    ws.insert_rows(1, 2)
    title_cell = ws.cell(row=1, column=1, value="Asset Risk GIS Mapping Data")
    title_cell.font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(visible_df.columns))
    title_cell.alignment = Alignment(horizontal="center")
    
    note_cell = ws.cell(row=2, column=1, value="GeoJSON data is embedded in a hidden column for GIS integration")
    note_cell.font = Font(italic=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(visible_df.columns))
    note_cell.alignment = Alignment(horizontal="center")
    
    # Add a map preview link (using OpenStreetMap)
    map_row = len(visible_data) + 5
    map_header = ws.cell(row=map_row, column=1, value="Map Preview")
    map_header.font = Font(bold=True, size=12)
    
    map_instr = ws.cell(row=map_row+1, column=1, 
                        value="Open the GeoJSON file in http://geojson.io or any GIS software to visualize the risk mapping")
    ws.merge_cells(start_row=map_row+1, start_column=1, end_row=map_row+1, end_column=len(visible_df.columns))
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Save the workbook
    wb.save(condition_report_path)
    
    return condition_report_path

def export_risk_mapping(condition_df, assets_df, output_dir, filename_prefix="asset_risk"):
    """
    Export asset risk data for GIS mapping by:
    1. Generating a GeoJSON file with location coordinates and risk categories
    2. Embedding the GeoJSON in the Excel report for easy GIS integration
    
    Args:
        condition_df (DataFrame): DataFrame with asset condition scores and risk categories
        assets_df (DataFrame): DataFrame with asset data including location coordinates
        output_dir (str): Directory where output files will be saved
        filename_prefix (str): Prefix for output filenames
        
    Returns:
        tuple: (geojson_path, excel_path) - Paths to the generated GeoJSON and Excel files
    """
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Merge condition data with asset coordinates
    # Assuming assets_df has columns: asset_id, latitude, longitude (or similar)
    geo_columns = ['latitude', 'longitude']
    
    # Check if we have alternate column names for coordinates
    if not all(col in assets_df.columns for col in geo_columns):
        # Try alternative column names
        alt_geo_columns = {
            'latitude': ['lat', 'y_coord', 'y'],
            'longitude': ['lon', 'long', 'x_coord', 'x']
        }
        
        for standard, alternatives in alt_geo_columns.items():
            for alt in alternatives:
                if alt in assets_df.columns:
                    assets_df[standard] = assets_df[alt]
                    break
    
    # Check if we have the required columns now
    if not all(col in assets_df.columns for col in geo_columns):
        raise ValueError(f"Asset data is missing required location columns: {geo_columns}")
    
    # Merge on asset ID
    mapping_df = pd.merge(
        condition_df[['Asset ID', 'Asset Name', 'Asset Type', 'Location', 'Weight-Adjusted Score', 'Risk Category']],
        assets_df[['asset_id', 'latitude', 'longitude']],
        left_on='Asset ID',
        right_on='asset_id',
        how='left'
    )
    
    # Drop duplicate column
    mapping_df = mapping_df.drop(columns=['asset_id'])
    
    # Check for missing coordinates and log warning
    missing_coords = mapping_df[mapping_df['latitude'].isna() | mapping_df['longitude'].isna()]
    if not missing_coords.empty:
        print(f"Warning: {len(missing_coords)} assets are missing coordinates and will not be included in the GeoJSON export.")
        # Filter out assets with missing coordinates
        mapping_df = mapping_df.dropna(subset=['latitude', 'longitude'])
    
    # Convert to GeoJSON format
    features = []
    
    for _, row in mapping_df.iterrows():
        # Get color code based on risk category
        color = get_risk_color(row['Risk Category'])
        
        # Create GeoJSON feature
        feature = {
            "type": "Feature",
            "properties": {
                "asset_id": row['Asset ID'],
                "name": row['Asset Name'],
                "type": row['Asset Type'],
                "location": row['Location'],
                "score": float(row['Weight-Adjusted Score']),
                "risk_category": row['Risk Category'],
                "color": color,
                "marker-color": color,  # For GeoJSON.io compatibility
                "marker-size": "medium"
            },
            "geometry": {
                "type": "Point",
                "coordinates": [float(row['longitude']), float(row['latitude'])]
            }
        }
        features.append(feature)
    
    # Create GeoJSON structure
    geojson = {
        "type": "FeatureCollection",
        "features": features
    }
    
    # Save GeoJSON to file
    geojson_path = os.path.join(output_dir, f"{filename_prefix}.geojson")
    with open(geojson_path, 'w') as f:
        json.dump(geojson, f, indent=2)
    
    # Create Excel file with embedded GeoJSON
    excel_path = os.path.join(output_dir, f"{filename_prefix}.xlsx")
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Risk Mapping Data"
    
    # Add visible data
    for r_idx, row in enumerate(dataframe_to_rows(mapping_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            
            # Format header row
            if r_idx == 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Format Risk Category cells with conditional coloring
            elif c_idx == mapping_df.columns.get_loc('Risk Category') + 1 and r_idx > 1:
                color = get_risk_color(value)
                # Convert hex color to RGB for Excel
                rgb = hex_to_rgb(color)
                cell.fill = PatternFill(start_color=color[1:], end_color=color[1:], fill_type="solid")
    
    # Add GeoJSON data in a hidden column
    geojson_str = json.dumps(geojson)
    geojson_cell = ws.cell(row=1, column=len(mapping_df.columns) + 2, value="GeoJSON_Data")
    geojson_cell.font = Font(bold=True)
    
    # We'll split the GeoJSON string into chunks to avoid Excel cell size limitations
    # Excel has a limit of about 32,767 characters per cell
    chunk_size = 30000
    geojson_chunks = [geojson_str[i:i+chunk_size] for i in range(0, len(geojson_str), chunk_size)]
    
    for i, chunk in enumerate(geojson_chunks):
        ws.cell(row=i+2, column=len(mapping_df.columns) + 2, value=chunk)
    
    # Add instructions for GIS integration
    instructions_row = len(mapping_df) + 5
    instruction_cell = ws.cell(row=instructions_row, column=1, value="GIS Integration Instructions")
    instruction_cell.font = Font(bold=True, size=12)
    ws.merge_cells(start_row=instructions_row, start_column=1, end_row=instructions_row, end_column=6)
    
    instructions = [
        f"1. The GeoJSON file has been saved to: {geojson_path}",
        "2. For GIS software integration:",
        "   - Import the GeoJSON directly into QGIS, ArcGIS, or similar software",
        "   - Use the 'color' property for symbology",
        "3. For web mapping:",
        "   - Use the GeoJSON with Leaflet, OpenLayers, or Mapbox",
        "   - The 'marker-color' property is compatible with GeoJSON.io",
        "4. To extract the GeoJSON from this Excel file:",
        "   - The complete GeoJSON is in the hidden column to the right",
        "   - Use the concatenated values to recreate the GeoJSON file"
    ]
    
    for i, line in enumerate(instructions):
        ws.cell(row=instructions_row + 1 + i, column=1, value=line)
        ws.merge_cells(start_row=instructions_row + 1 + i, start_column=1, 
                      end_row=instructions_row + 1 + i, end_column=6)
    
    # Add a summary of risk categories
    summary_row = instructions_row + len(instructions) + 3
    summary_cell = ws.cell(row=summary_row, column=1, value="Risk Category Summary")
    summary_cell.font = Font(bold=True, size=12)
    ws.merge_cells(start_row=summary_row, start_column=1, end_row=summary_row, end_column=3)
    
    # Add headers for summary
    ws.cell(row=summary_row + 1, column=1, value="Risk Category").font = Font(bold=True)
    ws.cell(row=summary_row + 1, column=2, value="Count").font = Font(bold=True)
    ws.cell(row=summary_row + 1, column=3, value="Color").font = Font(bold=True)
    
    # Add summary data
    category_counts = mapping_df['Risk Category'].value_counts().to_dict()
    categories = ['Excellent', 'Good', 'Fair', 'Poor', 'Critical']
    
    for i, category in enumerate(categories):
        row_num = summary_row + 2 + i
        count = category_counts.get(category, 0)
        color = get_risk_color(category)
        
        ws.cell(row=row_num, column=1, value=category)
        ws.cell(row=row_num, column=2, value=count)
        color_cell = ws.cell(row=row_num, column=3, value=color)
        color_cell.fill = PatternFill(start_color=color[1:], end_color=color[1:], fill_type="solid")
    
    # Auto adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Hide the GeoJSON column
    col_letter = get_column_letter(len(mapping_df.columns) + 2)
    ws.column_dimensions[col_letter].hidden = True
    
    # Save the Excel file
    wb.save(excel_path)
    
    return geojson_path, excel_path

def get_risk_color(risk_category):
    """
    Get the color code for a given risk category.
    
    Args:
        risk_category (str): Risk category name
        
    Returns:
        str: Hex color code
    """
    risk_colors = {
        'Excellent': '#00FF00',  # Bright Green
        'Good': '#CCFFCC',       # Light Green
        'Fair': '#FFFF99',       # Yellow
        'Poor': '#FF9999',       # Light Red
        'Critical': '#FF0000'    # Bright Red
    }
    
    return risk_colors.get(risk_category, '#CCCCCC')  # Default to gray if not found

def hex_to_rgb(hex_color):
    """
    Convert hex color to RGB.
    
    Args:
        hex_color (str): Hex color code (with or without #)
        
    Returns:
        tuple: (R, G, B) values
    """
    # Remove the # if present
    hex_color = hex_color.lstrip('#')
    
    # Convert to RGB
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))