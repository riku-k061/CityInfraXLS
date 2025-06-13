# update_task.py

import os
import sys
import pandas as pd
import argparse
from datetime import datetime

# Add parent directory to path to import modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.excel_handler import create_tasks_sheet

def load_tasks():
    """Load tasks from tasks.xlsx, creating the file if it doesn't exist."""
    tasks_file = "data/tasks.xlsx"
    
    # Create tasks sheet if it doesn't exist
    if not os.path.exists(tasks_file):
        print(f"Tasks file not found, creating {tasks_file}...")
        create_tasks_sheet(tasks_file)
        # Return empty DataFrame with correct columns
        return pd.DataFrame(columns=["Task ID", "Incident ID", "Contractor ID", 
                                    "Assigned At", "Status", "Status Updated At", "Details"])
    
    try:
        tasks_df = pd.read_excel(tasks_file)
        
        # Check for required columns
        required_columns = ["Task ID", "Status", "Details"]
        missing_columns = [col for col in required_columns if col not in tasks_df.columns]
        
        if missing_columns:
            print(f"Error: Required columns missing from tasks file: {', '.join(missing_columns)}")
            print("Please ensure the tasks file has the correct structure.")
            return None
        
        # Add Status Updated At column if it doesn't exist
        if "Status Updated At" not in tasks_df.columns:
            tasks_df["Status Updated At"] = None
            
        return tasks_df
    except Exception as e:
        print(f"Error loading tasks: {e}")
        return None

def display_tasks(tasks_df):
    """Display tasks for selection."""
    if tasks_df is None or tasks_df.empty:
        print("No tasks found.")
        return None
    
    print("\nTasks:")
    for idx, row in tasks_df.iterrows():
        print(f"{idx+1}. Task ID: {row['Task ID']} - Incident ID: {row['Incident ID']} - Status: {row['Status']}")
    
    while True:
        try:
            choice = int(input("\nSelect task number (or 0 to cancel): "))
            if choice == 0:
                return None
            if 1 <= choice <= len(tasks_df):
                return tasks_df.iloc[choice-1]['Task ID']
            print("Invalid choice. Please try again.")
        except ValueError:
            print("Please enter a valid number.")

def update_task(task_id, new_status, change_note=""):
    """Update the status of a task and append change note to Details."""
    tasks_file = "data/tasks.xlsx"
    
    # Load existing tasks
    tasks_df = load_tasks()
    if tasks_df is None:
        return False
    
    # If no tasks exist yet, we can't update
    if tasks_df.empty:
        print("No tasks found to update.")
        return False
    
    # Check if task exists
    if task_id not in tasks_df['Task ID'].values:
        print(f"Error: Task ID {task_id} not found.")
        return False
    
    # Get task index
    task_idx = tasks_df.index[tasks_df['Task ID'] == task_id].tolist()[0]
    current_status = tasks_df.at[task_idx, 'Status']
    
    # Skip update if status is the same
    if current_status == new_status:
        print(f"Task is already in '{new_status}' status. No update needed.")
        return True
    
    # Format timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Format change note with timestamp
    status_note = f"[{timestamp}] Status changed from '{current_status}' to '{new_status}'"
    if change_note:
        status_note += f": {change_note}"
    
    # Update Details field - append the new note
    if pd.isna(tasks_df.at[task_idx, 'Details']) or tasks_df.at[task_idx, 'Details'] == "":
        tasks_df.at[task_idx, 'Details'] = status_note
    else:
        tasks_df.at[task_idx, 'Details'] = f"{tasks_df.at[task_idx, 'Details']}\n{status_note}"
    
    # Update Status field
    tasks_df.at[task_idx, 'Status'] = new_status
    
    # Update Status Updated At field
    tasks_df.at[task_idx, 'Status Updated At'] = timestamp
    
    # Save updated tasks with error handling
    try:
        tasks_df.to_excel(tasks_file, index=False)
        print(f"\nTask updated successfully!")
        print(f"Task ID: {task_id}")
        print(f"New Status: {new_status}")
        print(f"Updated At: {timestamp}")
        if change_note:
            print(f"Change Note: {change_note}")
        return True
    except PermissionError:
        print(f"Error: Could not save to {tasks_file}. The file may be open in another application.")
        print("Please close the file and try again.")
        return False
    except Exception as e:
        print(f"Error saving task update: {e}")
        return False

def main():
    valid_statuses = ['Assigned', 'In Progress', 'Completed']
    
    parser = argparse.ArgumentParser(description="Update task status in CityInfraXLS")
    parser.add_argument("--task-id", help="ID of the task to update")
    parser.add_argument("--status", choices=valid_statuses, help="New status for the task")
    parser.add_argument("--note", help="Additional note about the status change", default="")
    args = parser.parse_args()
    
    # Load tasks (creates tasks file if it doesn't exist)
    tasks_df = load_tasks()
    if tasks_df is None:
        # Error message already printed in load_tasks()
        sys.exit(1)
    
    # If no tasks exist yet and we're trying to update
    if tasks_df.empty:
        print("No tasks found. Please assign tasks before updating.")
        sys.exit(1)
    
    # Get task ID from args or prompt user
    task_id = args.task_id
    if not task_id:
        task_id = display_tasks(tasks_df)
        if not task_id:
            print("Task update cancelled.")
            sys.exit(0)
    else:
        # Validate provided task ID
        if task_id not in tasks_df['Task ID'].values:
            print(f"Error: Task ID {task_id} not found.")
            sys.exit(1)
    
    # Get new status from args or prompt user
    new_status = args.status
    if not new_status:
        print("\nAvailable statuses:")
        for i, status in enumerate(valid_statuses):
            print(f"{i+1}. {status}")
        
        while True:
            try:
                choice = int(input("\nSelect new status (or 0 to cancel): "))
                if choice == 0:
                    print("Task update cancelled.")
                    sys.exit(0)
                if 1 <= choice <= len(valid_statuses):
                    new_status = valid_statuses[choice-1]
                    break
                print("Invalid choice. Please try again.")
            except ValueError:
                print("Please enter a valid number.")
    
    # Get change note from args or prompt user
    change_note = args.note
    if not change_note:
        change_note = input("Enter note about status change (optional): ")
    
    # Update the task
    success = update_task(task_id, new_status, change_note)
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()