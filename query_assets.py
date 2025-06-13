# query_assets.py

import os
import argparse
import pandas as pd
from pathlib import Path
from datetime import datetime
import json
from tabulate import tabulate

# Constants
ASSETS_PATH = 'data/assets.xlsx'
SCHEMA_PATH = 'asset_schema.json'

def load_schema():
    """Load the asset schema from JSON"""
    with open(SCHEMA_PATH, 'r') as f:
        return json.load(f)

def find_date_column(asset_type, schema):
    """Determine which column contains date information for this asset type"""
    date_keywords = ['Date', 'Installed', 'Built', 'Maintenance', 'Inspection']
    
    # Get the fields for this asset type
    fields = schema.get(asset_type, [])
    
    # Look for field names that contain date-related keywords
    for field in fields:
        for keyword in date_keywords:
            if keyword.lower() in field.lower():
                return field
    
    # Default to 'Installation Date' if available, otherwise None
    return 'Installation Date' if 'Installation Date' in fields else None

def apply_filters(df, asset_type, location, installed_after, schema):
    """Apply all filters to a dataframe and return filtered results"""
    original_row_count = len(df)
    filtered_row_count = original_row_count
    
    # Apply location filter if specified
    if location:
        # Look for any column containing "Location" in its name
        location_cols = [col for col in df.columns if 'Location' in col]
        if location_cols:
            location_col = location_cols[0]  # Use the first location column
            mask = df[location_col].str.contains(location, case=False, na=False)
            df = df[mask]
            filtered_row_count = len(df)
            print(f"Location filter: {original_row_count - filtered_row_count} rows filtered out")
        else:
            print(f"Warning: No location column found in {asset_type} sheet")
    
    # Apply date filter if specified
    if installed_after and not df.empty:
        date_col = find_date_column(asset_type, schema)
        
        if date_col and date_col in df.columns:
            # Parse the filter date
            filter_date = datetime.strptime(installed_after, '%Y-%m-%d').date()
            
            before_filter_count = len(df)
            
            # Convert date column to datetime, invalid values become NaT
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            
            # Count rows with invalid dates (NaT values)
            invalid_dates_count = df[date_col].isna().sum()
            
            if invalid_dates_count > 0:
                print(f"Warning: {invalid_dates_count} rows have invalid date formats in '{date_col}' column")
            
            # Filter out rows with NaT values
            df = df[~df[date_col].isna()]
            
            # Now safely convert to date for comparison (no NaT values remain)
            if not df.empty:
                df[date_col] = df[date_col].dt.date
                
                # Filter by date
                pre_date_filter_count = len(df)
                df = df[df[date_col] >= filter_date]
                date_filtered = pre_date_filter_count - len(df)
            else:
                date_filtered = 0
            
            # Total rows filtered by date operation
            total_date_filtered = before_filter_count - len(df)
            print(f"Date filter: {total_date_filtered} rows filtered out ({invalid_dates_count} invalid dates, {date_filtered} before {installed_after})")
        else:
            print(f"Warning: No suitable date column found for filtering")
    
    return df

def query_assets(asset_type=None, location=None, installed_after=None, export_path=None):
    """Query assets based on filters and display/export results"""
    # Check if assets file exists
    if not os.path.exists(ASSETS_PATH):
        print(f"Error: Assets file not found at {ASSETS_PATH}")
        return
    
    # Load schema to get available asset types
    schema = load_schema()
    available_types = list(schema.keys())
    
    # Early validation of asset type to avoid unnecessary processing
    if asset_type and asset_type not in available_types:
        print(f"Error: Invalid asset type '{asset_type}'")
        print(f"Available types: {', '.join(available_types)}")
        return
    
    # Determine which sheets to query
    sheets_to_query = [asset_type] if asset_type else available_types
    
    try:
        print(f"Loading asset data from {ASSETS_PATH}...")
        
        # Use ExcelFile for more efficient reading of multiple sheets
        with pd.ExcelFile(ASSETS_PATH) as excel_file:
            # Check that the requested sheets actually exist
            missing_sheets = [sheet for sheet in sheets_to_query if sheet not in excel_file.sheet_names]
            if missing_sheets:
                print(f"Warning: The following sheets were not found in the workbook: {', '.join(missing_sheets)}")
                sheets_to_query = [sheet for sheet in sheets_to_query if sheet in excel_file.sheet_names]
                if not sheets_to_query:
                    print("No valid sheets to query.")
                    return
            
            # Initialize results
            all_results = pd.DataFrame()
            
            # Process each sheet
            for sheet in sheets_to_query:
                print(f"Processing {sheet} assets...")
                
                # Read the sheet into a dataframe
                df = pd.read_excel(excel_file, sheet_name=sheet)
                
                if df.empty:
                    print(f"Sheet {sheet} is empty.")
                    continue
                
                # Add asset type column if querying multiple types
                if not asset_type:
                    df['Asset Type'] = sheet
                
                # Apply all filters
                filtered_df = apply_filters(df, sheet, location, installed_after, schema)
                
                # Append to results
                all_results = pd.concat([all_results, filtered_df])
        
        # Handle results
        if all_results.empty:
            print("No assets match the specified criteria.")
            return
        
        # Display results
        print("\nQuery Results:")
        print(tabulate(all_results, headers='keys', tablefmt='grid', showindex=False))
        print(f"\nTotal: {len(all_results)} assets found")
        
        # Export if requested
        if export_path:
            try:
                export_dir = os.path.dirname(export_path)
                if export_dir:
                    os.makedirs(export_dir, exist_ok=True)
                    
                all_results.to_excel(export_path, index=False)
                print(f"\nResults exported to {export_path}")
            except Exception as e:
                print(f"Error exporting results: {str(e)}")
                
    except Exception as e:
        print(f"Error querying assets: {str(e)}")

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Query CityInfraXLS assets')
    parser.add_argument('--type', help='Filter by asset type (Road, Park, Bridge, Streetlight, etc.)')
    parser.add_argument('--location', help='Filter by location (case insensitive, partial match)')
    parser.add_argument('--installed-after', help='Filter by installation date (YYYY-MM-DD)')
    parser.add_argument('--export', help='Export results to Excel file at specified path')
    
    args = parser.parse_args()
    
    # Check if we have at least one filter
    if not (args.type or args.location or args.installed_after):
        print("Please specify at least one filter: --type, --location, or --installed-after")
        parser.print_help()
        return
    
    # Validate date format if provided
    if args.installed_after:
        try:
            datetime.strptime(args.installed_after, '%Y-%m-%d')
        except ValueError:
            print("Error: Date format must be YYYY-MM-DD")
            return
    
    # Run the query
    query_assets(
        asset_type=args.type, 
        location=args.location, 
        installed_after=args.installed_after, 
        export_path=args.export
    )

if __name__ == "__main__":
    print("=== CityInfraXLS - Asset Query Tool ===")
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {str(e)}")