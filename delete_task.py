# delete_task.py

import os
import sys
import pandas as pd
import argparse
import logging
import shutil
from datetime import datetime

# Add parent directory to path to import modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.excel_handler import create_tasks_sheet

# Configure logging
logging.basicConfig(
    filename="cityinfraxls.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger("CityInfraXLS")

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
        return tasks_df
    except Exception as e:
        print(f"Error loading tasks: {e}")
        logger.error(f"Error loading tasks: {e}")
        return None

def create_backup(file_path):
    """Create a timestamped backup of the given file."""
    # Create backups directory if it doesn't exist
    backup_dir = os.path.join(os.path.dirname(file_path), "backups")
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
    
    # Generate timestamp and backup filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = os.path.basename(file_path)
    name, ext = os.path.splitext(file_name)
    backup_path = os.path.join(backup_dir, f"{name}_{timestamp}{ext}")
    
    # Create backup if original file exists
    if os.path.exists(file_path):
        try:
            shutil.copy2(file_path, backup_path)
            logger.info(f"Created backup: {backup_path}")
            return backup_path
        except Exception as e:
            logger.error(f"Failed to create backup: {e}")
            return None
    else:
        logger.warning(f"No backup created: {file_path} doesn't exist")
        return None

def display_task(task_df):
    """Display task details in a formatted way."""
    print("\nTask Details:")
    print("-" * 50)
    for col in task_df.columns:
        value = task_df[col].iloc[0]
        if pd.isna(value):
            value = "N/A"
        print(f"{col}: {value}")
    print("-" * 50)

def delete_task(task_id, force=False):
    """Delete a task from the tasks workbook."""
    tasks_file = "data/tasks.xlsx"
    
    # Load existing tasks
    tasks_df = load_tasks()
    if tasks_df is None or tasks_df.empty:
        print("No tasks found.")
        return False
    
    # Check if task exists
    if task_id not in tasks_df['Task ID'].values:
        print(f"Error: Task ID {task_id} not found.")
        return False
    
    # Get task information for display
    task_info = tasks_df[tasks_df['Task ID'] == task_id]
    display_task(task_info)
    
    # Confirm deletion if not forced
    if not force:
        confirmation = input("\nAre you sure you want to delete this task? (yes/no): ").lower()
        if confirmation != 'yes':
            print("Task deletion cancelled.")
            return False
    
    # Create backup before deletion
    backup_file = create_backup(tasks_file)
    if backup_file:
        print(f"Backup created: {backup_file}")
    
    # Delete the task
    try:
        # Get the index and delete
        task_idx = tasks_df.index[tasks_df['Task ID'] == task_id].tolist()[0]
        tasks_df = tasks_df.drop(task_idx)
        
        # Save updated file
        tasks_df.to_excel(tasks_file, index=False)
        
        # Log the deletion
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logger.info(f"Task {task_id} deleted at {timestamp}")
        
        print(f"\nTask {task_id} successfully deleted.")
        return True
    except PermissionError:
        print(f"Error: Could not save to {tasks_file}. The file may be open in another application.")
        print("Please close the file and try again.")
        logger.error(f"Permission error while deleting task {task_id}")
        return False
    except Exception as e:
        print(f"Error deleting task: {e}")
        logger.error(f"Error deleting task {task_id}: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Delete a task from CityInfraXLS")
    parser.add_argument("--task-id", help="ID of the task to delete")
    parser.add_argument("--force", action="store_true", help="Skip confirmation prompt")
    args = parser.parse_args()
    
    # Load tasks
    tasks_df = load_tasks()
    if tasks_df is None:
        sys.exit(1)
    
    if tasks_df.empty:
        print("No tasks found.")
        sys.exit(0)
    
    # Get task ID from args or prompt user
    task_id = args.task_id
    if not task_id:
        print("\nAvailable tasks:")
        for idx, row in tasks_df.iterrows():
            print(f"{idx+1}. Task ID: {row['Task ID']} - Incident ID: {row.get('Incident ID', 'N/A')} - Status: {row.get('Status', 'N/A')}")
        
        while True:
            try:
                choice = int(input("\nSelect task number to delete (or 0 to cancel): "))
                if choice == 0:
                    print("Task deletion cancelled.")
                    sys.exit(0)
                if 1 <= choice <= len(tasks_df):
                    task_id = tasks_df.iloc[choice-1]['Task ID']
                    break
                print("Invalid choice. Please try again.")
            except ValueError:
                print("Please enter a valid number.")
    
    # Delete the task
    success = delete_task(task_id, args.force)
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()