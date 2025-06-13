# assign_task.py

import os
import sys
import json
import uuid
import pandas as pd
import argparse
from datetime import datetime

# Add parent directory to path to import modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.excel_handler import create_sheets_from_schema, create_tasks_sheet

def load_schema(schema_path):
    """Load JSON schema from file."""
    with open(schema_path, 'r') as f:
        return json.load(f)

def validate_contractor(contractor_id, contractors_df, contractors_schema):
    """Validate that contractor exists and meets requirements."""
    if contractor_id not in contractors_df['contractor_id'].values:
        return False, "Contractor ID not found in database."
    return True, "Contractor validation successful."

def load_open_incidents():
    """Load open incidents from incidents.xlsx."""
    try:
        incidents_df = pd.read_excel("data/incidents.xlsx")
        # Filter for open incidents (assuming 'Status' column with 'Open' value)
        open_incidents = incidents_df[incidents_df['Status'] == 'Open']
        return open_incidents
    except Exception as e:
        print(f"Error loading incidents: {e}")
        return pd.DataFrame()

def load_contractors():
    """Load contractors from contractors.xlsx."""
    try:
        return pd.read_excel("data/contractors.xlsx")
    except Exception as e:
        print(f"Error loading contractors: {e}")
        return pd.DataFrame()

def display_open_incidents(incidents_df):
    """Display open incidents for selection."""
    if incidents_df.empty:
        print("No open incidents found.")
        return None
    
    print("\nOpen Incidents:")
    for idx, row in incidents_df.iterrows():
        print(f"{idx+1}. ID: {row['Incident ID']} - {row['Severity'][:50]}...")
    
    while True:
        try:
            choice = int(input("\nSelect incident number (or 0 to cancel): "))
            if choice == 0:
                return None
            if 1 <= choice <= len(incidents_df):
                return incidents_df.iloc[choice-1]['Incident ID']
            print("Invalid choice. Please try again.")
        except ValueError:
            print("Please enter a valid number.")

def display_contractors(contractors_df):
    """Display contractors for selection."""
    if contractors_df.empty:
        print("No contractors found.")
        return None
    
    print("\nAvailable Contractors:")
    for idx, row in contractors_df.iterrows():
        print(f"{idx+1}. ID: {row['contractor_id']} - {row['name']} - Specialties: {', '.join(row['specialties'])} - Rating: {row['rating']}")
    
    while True:
        try:
            choice = int(input("\nSelect contractor number (or 0 to cancel): "))
            if choice == 0:
                return None
            if 1 <= choice <= len(contractors_df):
                return contractors_df.iloc[choice-1]['contractor_id']
            print("Invalid choice. Please try again.")
        except ValueError:
            print("Please enter a valid number.")

def assign_task(incident_id, contractor_id, details=""):
    """Assign a task to a contractor for an incident."""
    tasks_file = "data/tasks.xlsx"
    
    # Check if tasks file exists, if not create it
    if not os.path.exists(tasks_file):
        create_tasks_sheet(tasks_file)
    
    # Load existing tasks
    try:
        tasks_df = pd.read_excel(tasks_file)
    except Exception as e:
        print(f"Error loading tasks: {e}")
        tasks_df = pd.DataFrame(columns=["Task ID", "Incident ID", "Contractor ID", "Assigned At", "Status", "Details"])
    
    # Create new task
    new_task = {
        "Task ID": str(uuid.uuid4()),
        "Incident ID": incident_id,
        "Contractor ID": contractor_id,
        "Assigned At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Status": "Assigned",
        "Details": details
    }
    
    # Append new task
    tasks_df = tasks_df._append(new_task, ignore_index=True)
    
    # Save updated tasks
    tasks_df.to_excel(tasks_file, index=False)
    
    print(f"\nTask assigned successfully!")
    print(f"Task ID: {new_task['Task ID']}")
    print(f"Incident ID: {new_task['Incident ID']}")
    print(f"Contractor ID: {new_task['Contractor ID']}")
    print(f"Assigned At: {new_task['Assigned At']}")
    
    return new_task

def main():
    parser = argparse.ArgumentParser(description="Assign maintenance tasks to contractors")
    parser.add_argument("--contractor-id", help="Contractor ID to assign the task")
    parser.add_argument("--incident-id", help="Incident ID to assign")
    parser.add_argument("--details", help="Additional task details", default="")
    args = parser.parse_args()
    
    # Ensure required files exist
    if not os.path.exists("data/contractors.xlsx"):
        create_sheets_from_schema("contractors_schema.json", "data/contractors.xlsx")
        print("Created data/contractors.xlsx. Please add contractor data before proceeding.")
        sys.exit(1)
    
    if not os.path.exists("data/incidents.xlsx"):
        print("No incidents file found. Please register incidents before assigning tasks.")
        sys.exit(1)
    
    # Load contractors schema
    contractors_schema = load_schema("contractors_schema.json")
    
    # Load contractors and open incidents
    contractors_df = load_contractors()
    if contractors_df.empty:
        print("No contractors found. Please add contractors first.")
        sys.exit(1)
    
    open_incidents_df = load_open_incidents()
    if open_incidents_df.empty:
        print("No open incidents found. Please register incidents first.")
        sys.exit(1)
    
    # Get incident ID from args or prompt user
    incident_id = args.incident_id
    if not incident_id:
        incident_id = display_open_incidents(open_incidents_df)
        if not incident_id:
            print("Task assignment cancelled.")
            sys.exit(0)
    else:
        # Validate provided incident ID
        if incident_id not in open_incidents_df['ID'].values:
            print(f"Error: Incident ID {incident_id} not found or not open.")
            sys.exit(1)
    
    # Get contractor ID from args or prompt user
    contractor_id = args.contractor_id
    if not contractor_id:
        contractor_id = display_contractors(contractors_df)
        if not contractor_id:
            print("Task assignment cancelled.")
            sys.exit(0)
    
    # Validate contractor
    is_valid, message = validate_contractor(contractor_id, contractors_df, contractors_schema)
    if not is_valid:
        print(f"Error: {message}")
        sys.exit(1)
    
    # Get task details
    details = args.details
    if not details:
        details = input("Enter task details (optional): ")
    
    # Assign the task
    task = assign_task(incident_id, contractor_id, details)

if __name__ == "__main__":
    main()