# export_budget_alerts.py
import os
import csv
import pandas as pd
import openpyxl
import time
from datetime import datetime
from pathlib import Path
import logging
import shutil

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='budget_alerts_export.log'
)
logger = logging.getLogger('budget_alerts')

def export_alerts_to_csv(
    source_excel='data/budget_allocations.xlsx', 
    output_csv='data/exports/budget_alerts.csv',
    backup=True
):
    """
    Export only the budget alert rows from the Alerts sheet to a standalone CSV file
    for the notification system.
    
    Parameters:
        source_excel: Path to the Excel file containing the Alerts sheet
        output_csv: Path where the CSV file should be saved
        backup: Whether to create a backup of any existing CSV before overwriting
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Ensure the export directory exists
        os.makedirs(os.path.dirname(output_csv), exist_ok=True)
        
        # Check if the Alerts sheet exists in the workbook
        workbook = openpyxl.load_workbook(source_excel, read_only=True)
        if 'Alerts' not in workbook.sheetnames:
            logger.error(f"No Alerts sheet found in {source_excel}")
            return False
            
        # Read the Alerts sheet
        alerts_df = pd.read_excel(source_excel, sheet_name='Alerts')
        
        # Extract only the required columns
        if all(col in alerts_df.columns for col in ['department', 'project_id', 'allocated_amount', 'remaining_budget', 'overrun_amount', 'status']):
            # Filter rows - typically these would be rows with critical status
            critical_alerts = alerts_df[alerts_df['status'].isin(['Over Budget', 'At Risk'])]
            
            # Select only the needed columns for the notification system
            export_df = critical_alerts[['department', 'project_id', 'remaining_budget', 'overrun_amount', 'status']]
            
            # Add timestamp column
            export_df['export_timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            # Create backup if requested and file exists
            if backup and os.path.exists(output_csv):
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_path = f"{os.path.splitext(output_csv)[0]}_{timestamp}_backup.csv"
                shutil.copy2(output_csv, backup_path)
                logger.info(f"Created backup at {backup_path}")
            
            # Export to CSV
            export_df.to_csv(output_csv, index=False, quoting=csv.QUOTE_NONNUMERIC)
            logger.info(f"Successfully exported {len(export_df)} alert records to {output_csv}")
            
            # Create a sync marker file to indicate last sync time
            with open(f"{os.path.splitext(output_csv)[0]}_last_sync.txt", 'w') as f:
                f.write(f"Last synchronized: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            return True
        else:
            logger.error("Required columns not found in Alerts sheet")
            return False
            
    except Exception as e:
        logger.error(f"Error exporting alerts: {str(e)}")
        return False

def setup_auto_export(
    source_excel='data/budget_allocations.xlsx',
    output_csv='data/exports/budget_alerts.csv',
    check_interval=300  # 5 minutes in seconds
):
    """
    Set up a continuous monitoring process that checks for changes in the Alerts sheet
    and automatically exports updates to CSV when changes are detected.
    
    Parameters:
        source_excel: Path to the Excel file containing the Alerts sheet
        output_csv: Path where the CSV file should be saved
        check_interval: How often to check for changes (in seconds)
    """
    logger.info(f"Starting budget alerts auto-export monitoring (interval: {check_interval}s)")
    
    last_modified = 0
    
    try:
        while True:
            current_modified = os.path.getmtime(source_excel)
            
            # If the file has been modified since last check
            if current_modified > last_modified:
                logger.info(f"Changes detected in {source_excel}")
                if export_alerts_to_csv(source_excel, output_csv):
                    last_modified = current_modified
                    logger.info("Alerts successfully exported after change detection")
                else:
                    logger.error("Failed to export alerts after change detection")
            
            time.sleep(check_interval)
            
    except KeyboardInterrupt:
        logger.info("Auto-export monitoring stopped by user")
    except Exception as e:
        logger.error(f"Error in auto-export monitoring: {str(e)}")

def main():
    """
    Main function to handle command line arguments and run the appropriate export mode.
    """
    import argparse
    
    parser = argparse.ArgumentParser(description='Export budget alerts to CSV for notification system')
    parser.add_argument('--source', default='data/budget_allocations.xlsx', help='Source Excel file path')
    parser.add_argument('--output', default='data/exports/budget_alerts.csv', help='Output CSV file path')
    parser.add_argument('--monitor', action='store_true', help='Run in continuous monitoring mode')
    parser.add_argument('--interval', type=int, default=300, help='Checking interval in seconds (for monitor mode)')
    
    args = parser.parse_args()
    
    if args.monitor:
        setup_auto_export(args.source, args.output, args.interval)
    else:
        success = export_alerts_to_csv(args.source, args.output)
        if success:
            print(f"Exported alerts to {args.output}")
        else:
            print("Failed to export alerts. See log for details.")

if __name__ == "__main__":
    main()