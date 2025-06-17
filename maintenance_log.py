# maintenance_log.py

import os
import json
import uuid
import datetime
import pandas as pd
from utils.excel_handler import create_maintenance_history_sheet

def load_schema(schema_path='maintenance_schema.json'):
    """Load the maintenance schema from a JSON file."""
    try:
        with open(schema_path, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Schema file not found at {schema_path}")
        return None
    except json.JSONDecodeError:
        print(f"Error decoding JSON from {schema_path}")
        return None

def validate_input(value, field_name, field_schema):
    """Validate user input against schema requirements."""
    field_type = field_schema.get('type')
    
    # Handle empty input for required fields
    if value.strip() == "":
        if field_name in schema.get('required', []):
            return False, f"{field_name} is required"
        return True, None  # Empty value for optional field is allowed
    
    # Type validation
    if field_type == 'string':
        # Handle enum validation
        if 'enum' in field_schema:
            if value not in field_schema['enum']:
                return False, f"{value} is not a valid option. Choose from: {', '.join(field_schema['enum'])}"
        
        # Handle date format
        if field_schema.get('format') == 'date':
            try:
                datetime.datetime.strptime(value, '%Y-%m-%d')
            except ValueError:
                return False, "Date must be in YYYY-MM-DD format"
    
    elif field_type == 'number':
        try:
            float(value)
        except ValueError:
            return False, f"{field_name} must be a number"
    
    return True, None

def get_validated_input(prompt, field_name, field_schema):
    """Get user input with validation until valid input is provided."""
    while True:
        user_input = input(prompt)
        is_valid, error_message = validate_input(user_input, field_name, field_schema)
        
        if is_valid:
            # Convert to appropriate type if needed
            if field_schema.get('type') == 'number':
                return float(user_input) if user_input.strip() else None
            return user_input
        else:
            print(error_message)

def log_maintenance():
    """Log a new maintenance record."""
    # Load the schema
    global schema
    schema = load_schema()
    if not schema:
        print("Cannot log maintenance without schema")
        return False
    
    # Define the expected column order based on schema
    expected_columns = list(schema['properties'].keys())
    
    # Check if the maintenance history file exists, create it if not
    excel_path = 'data/maintenance_history.xlsx'
    if not os.path.exists(excel_path):
        success = create_maintenance_history_sheet(excel_path)
        if not success:
            print("Failed to create maintenance history sheet")
            return False
    
    # Prepare the new record
    record = {}
    
    # Generate UUID for record_id
    record['record_id'] = str(uuid.uuid4())
    print(f"Record ID: {record['record_id']} (generated automatically)")
    
    # Get asset_id
    record['asset_id'] = get_validated_input(
        "Enter Asset ID: ",
        'asset_id',
        schema['properties']['asset_id']
    )
    
    # Get action_taken (enum)
    action_options = schema['properties']['action_taken']['enum']
    print("Available actions:")
    for i, action in enumerate(action_options, 1):
        print(f"{i}. {action}")
    
    while True:
        try:
            choice = int(input("Select action (enter number): "))
            if 1 <= choice <= len(action_options):
                record['action_taken'] = action_options[choice - 1]
                break
            else:
                print("Invalid choice, please select a valid number")
        except ValueError:
            print("Please enter a number")
    
    # Get performed_by
    record['performed_by'] = get_validated_input(
        "Enter name of person/team who performed the maintenance: ",
        'performed_by',
        schema['properties']['performed_by']
    )
    
    # Get cost (optional)
    cost_input = get_validated_input(
        "Enter cost (leave blank if unknown): ",
        'cost',
        schema['properties']['cost']
    )
    record['cost'] = float(cost_input)
    
    # Get date
    date_input = get_validated_input(
        "Enter date (YYYY-MM-DD): ",
        'date',
        schema['properties']['date']
    )
    record['date'] = date_input
    
    # Get notes (optional)
    record['notes'] = get_validated_input(
        "Enter notes (optional): ",
        'notes',
        schema['properties']['notes']
    )
    
    # Read existing data - explicitly target the Maintenance History sheet
    try:
        # Read existing data
        df = pd.read_excel(excel_path, sheet_name="Maintenance History")
        
        # Fix column ordering issue - ensure all schema columns exist
        # Add any missing columns with NaN values
        for col in expected_columns:
            if col not in df.columns:
                df[col] = pd.NA
                
        # Reorder columns according to schema and drop any columns not in schema
        df = df.reindex(columns=expected_columns)
    except Exception as e:
        print(f"Note: Starting with a new sheet due to: {e}")
        # If reading fails, create a new DataFrame with correct column order
        df = pd.DataFrame(columns=expected_columns)
    
    # Append new record, ensuring it follows the expected column order
    new_record_df = pd.DataFrame([record], columns=expected_columns)
    df = pd.concat([df, new_record_df], ignore_index=True)
    
    # Save updated data
    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name="Maintenance History", index=False)
        print(f"Maintenance record logged successfully with ID: {record['record_id']}")
        return True
    except Exception as e:
        print(f"Error saving maintenance record: {e}")
        return False

if __name__ == "__main__":
    log_maintenance()