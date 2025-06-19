# update_complaint.py

"""
CityInfraXLS - Update Complaint Script
Updates existing complaint records in the system
"""

import os
import sys
import uuid
import argparse
import json
import pandas as pd
from datetime import datetime
from pathlib import Path
sys.path.append(str(Path(__file__).parent))
from utils.excel_handler import create_complaint_sheet

# Constants
COMPLAINTS_EXCEL = "data/complaints.xlsx"
SCHEMA_PATH = "complaint_schema.json"


def load_schema():
    """Load and validate complaint schema"""
    try:
        with open(SCHEMA_PATH, 'r') as schema_file:
            schema = json.load(schema_file)
            return schema
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error loading schema: {str(e)}")
        sys.exit(1)


def update_complaint(complaint_id, status=None, note=None):
    """Updates a complaint record with new status and appends notes"""
    
    # Ensure the complaints Excel file and sheet exist
    create_complaint_sheet(COMPLAINTS_EXCEL)
    
    if not os.path.exists(COMPLAINTS_EXCEL):
        print(f"Error: Complaints file {COMPLAINTS_EXCEL} not found")
        return False
        
    # Load the schema to validate status
    schema = load_schema()
    valid_statuses = schema.get('properties', {}).get('status', {}).get('enum', ['Open', 'In Progress', 'Closed'])
    
    # Validate status if provided
    if status and status not in valid_statuses:
        print(f"Error: Invalid status. Must be one of: {', '.join(valid_statuses)}")
        return False
    
    try:
        # Load the Excel file
        try:
            df = pd.read_excel(COMPLAINTS_EXCEL)
            
            # Normalize column names to lowercase with underscores
            df.columns = [col.lower().replace(' ', '_') for col in df.columns]
            
            # Add missing columns if needed
            if 'resolution_notes' not in df.columns:
                df['resolution_notes'] = None
            
            if 'closed_at' not in df.columns:
                df['closed_at'] = None
                
        except Exception as e:
            print(f"Error reading complaints Excel: {str(e)}")
            return False
        
        # Find the complaint to update
        complaint_index = df[df['complaint_id'] == complaint_id].index
        if len(complaint_index) == 0:
            print(f"Error: Complaint with ID {complaint_id} not found")
            return False
        
        # Get the first matching record
        complaint_index = complaint_index[0]
        
        # Handle note appending (with timestamp)
        if note:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
            new_note = f"[{timestamp}] {note}"
            
            # Append to existing notes if there are any
            existing_notes = df.at[complaint_index, 'resolution_notes']
            if pd.notnull(existing_notes) and existing_notes:
                df.at[complaint_index, 'resolution_notes'] = f"{existing_notes}\n{new_note}"
            else:
                df.at[complaint_index, 'resolution_notes'] = new_note
        
        # Update status if provided
        if status:
            current_status = df.at[complaint_index, 'status']
            df.at[complaint_index, 'status'] = status
            
            # Set closed_at timestamp when status becomes Closed
            if status == "Closed" and current_status != "Closed":
                df.at[complaint_index, 'closed_at'] = datetime.now()
                
            # Clear closed_at when status changes from Closed to another status
            elif status != "Closed" and current_status == "Closed":
                df.at[complaint_index, 'closed_at'] = None
        
        # Convert column names back to title case for writing
        df.columns = [col.replace('_', ' ').title() for col in df.columns]
        
        # Save changes back to the Excel file
        df.to_excel(COMPLAINTS_EXCEL, index=False)
        
        print(f"Complaint {complaint_id} updated successfully")
        return True
        
    except Exception as e:
        print(f"Error updating complaint: {str(e)}")
        return False


def main():
    """Main function to process command-line arguments"""
    parser = argparse.ArgumentParser(description='Update a complaint record')
    parser.add_argument('--id', required=True, help='Complaint ID to update')
    parser.add_argument('--status', help='New status for the complaint')
    parser.add_argument('--note', help='Resolution note to add to the complaint')
    
    args = parser.parse_args()
    
    # Ensure at least one update parameter is provided
    if not args.status and not args.note:
        print("Error: At least one of --status or --note must be specified")
        sys.exit(1)
    
    # Update the complaint
    success = update_complaint(args.id, args.status, args.note)
    
    # Return appropriate exit code
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()