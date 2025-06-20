# report_complaint.py

import os
import sys
import uuid
import pandas as pd
import json
from datetime import datetime
import pytz
import numpy as np

# Import the helper function
from utils.excel_handler import create_complaint_sheet

def report_complaint():
    """
    Interactive command-line interface for reporting a new infrastructure complaint.
    Prompts user for all required fields, generates ID and timestamps automatically,
    and saves the complaint to the complaints Excel file.
    
    Enhancements:
    - Dynamically reads field order from schema
    - Preserves schema column order in Excel
    - Properly handles null closed_at values
    - Drops any extraneous columns not in schema
    """
    # Ensure the complaints file exists with correct structure
    complaints_file = "data/complaints.xlsx"
    create_complaint_sheet(complaints_file)
    
    # Load complaint schema for field validation and order
    try:
        with open('complaint_schema.json', 'r') as schema_file:
            schema = json.load(schema_file)
    except Exception as e:
        print(f"Error loading complaint schema: {str(e)}")
        sys.exit(1)
    
    # Get schema field order - will be used to reindex DataFrame
    schema_fields = list(schema['properties'].keys())
    
    # Fields to skip in prompt (auto-generated)
    skip_fields = ['complaint_id', 'created_at', 'closed_at']
    
    # Initialize complaint with auto-generated fields
    complaint = {
        "complaint_id": str(uuid.uuid4()),
        "status": "Open",
        "created_at": datetime.now(pytz.UTC).isoformat(),
        "closed_at": None  # Will be converted to NaN before saving
    }
    
    # Prompt for user input for each field in schema order
    print("=== Report New Infrastructure Complaint ===")
    
    for field in schema_fields:
        # Skip auto-generated fields
        if field in skip_fields:
            continue
            
        # Special case for status which has a default
        if field == 'status':
            continue  # We're defaulting to "Open"
        
        # Get field properties
        field_props = schema['properties'][field]
        field_type = field_props.get('type', 'string')
        field_desc = field_props.get('description', field)
        
        # Handle different field types
        if field_type == 'integer':
            min_val = field_props.get('minimum', float('-inf'))
            max_val = field_props.get('maximum', float('inf'))
            
            valid_input = False
            while not valid_input:
                try:
                    value = input(f"{field_desc} ({min_val}-{max_val}): ").strip()
                    # Allow empty input for non-required fields
                    if not value and field not in schema.get('required', []):
                        complaint[field] = None
                        valid_input = True
                    else:
                        value = int(value)
                        if min_val <= value <= max_val:
                            complaint[field] = value
                            valid_input = True
                        else:
                            print(f"Value must be between {min_val} and {max_val}.")
                except ValueError:
                    print(f"Please enter a valid integer.")
        
        # Handle string fields
        elif field_type == 'string':
            # Check if field has enum values
            enum_values = field_props.get('enum', [])
            if enum_values:
                # Present as a menu if it has enum values
                print(f"{field_desc}:")
                for idx, option in enumerate(enum_values, 1):
                    print(f"  {idx}. {option}")
                
                valid_choice = False
                while not valid_choice:
                    choice = input(f"Select {field} (1-{len(enum_values)}): ").strip()
                    try:
                        choice_idx = int(choice) - 1
                        if 0 <= choice_idx < len(enum_values):
                            complaint[field] = enum_values[choice_idx]
                            valid_choice = True
                        else:
                            print(f"Please enter a number between 1 and {len(enum_values)}.")
                    except ValueError:
                        print("Please enter a valid number.")
            else:
                # Regular string input
                value = input(f"{field_desc}: ").strip()
                # Keep prompting if it's a required field and empty
                while not value and field in schema.get('required', []):
                    print(f"{field_desc} is required.")
                    value = input(f"{field_desc}: ").strip()
                complaint[field] = value
        
        # Handle other types as needed
        # (could expand for boolean, array, etc.)
        else:
            complaint[field] = input(f"{field}: ").strip()
    
    # Save the complaint to Excel
    try:
        # Load existing complaints
        try:
            df = pd.read_excel(complaints_file, sheet_name='Complaints')
        except Exception:
            # If sheet is empty or doesn't exist yet
            df = pd.DataFrame(columns=schema_fields)
        
        # Create a new row for the complaint
        new_row = pd.DataFrame([complaint])
        
        # Append new complaint
        df = pd.concat([df, new_row], ignore_index=True)
        
        # Convert 'None' values to NaN (which Excel will display as empty cells)
        df = df.replace('None', np.nan)
        df = df.replace(None, np.nan)
        
        # Reindex to ensure columns match schema order and drop any extra columns
        df = df.reindex(columns=schema_fields)
        
        # Save back to Excel
        with pd.ExcelWriter(complaints_file, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='Complaints', index=False)
        
        print(f"\nComplaint successfully registered with ID: {complaint['complaint_id']}")
        print(f"Status: {complaint['status']}")
        print(f"Created At: {complaint['created_at']}")
        return True
    
    except Exception as e:
        print(f"Error saving complaint: {str(e)}")
        return False

if __name__ == "__main__":
    report_complaint()