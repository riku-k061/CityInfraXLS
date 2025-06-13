import os
import sys
import argparse
import pandas as pd
from datetime import datetime
import logging
from tabulate import tabulate

# Configure logging
logging.basicConfig(
    filename='data/maintenance_query.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def parse_date(date_str):
    """
    Parse date string in YYYY-MM-DD format.
    
    Args:
        date_str (str): Date string in YYYY-MM-DD format
    
    Returns:
        datetime: Parsed datetime object
    
    Raises:
        ValueError: If date format is invalid
    """
    try:
        return datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        raise ValueError(f"Invalid date format: {date_str}. Expected format: YYYY-MM-DD")

def query_maintenance(from_date=None, to_date=None, action=None, export=False):
    """
    Query maintenance records by date range and action taken.
    
    Args:
        from_date (str): Start date in YYYY-MM-DD format
        to_date (str): End date in YYYY-MM-DD format
        action (str): Filter by action taken
        export (bool): Whether to export results to Excel
    
    Returns:
        bool: True if successful, False otherwise
    """
    excel_path = "data/maintenance_history.xlsx"
    
    # Check if file exists
    if not os.path.exists(excel_path):
        print(f"Error: {excel_path} not found.")
        return False
    
    try:
        # Load workbook into pandas dataframe
        df = pd.read_excel(excel_path, sheet_name="Maintenance History")
        
        # Check if dataframe is empty
        if df.empty:
            print("Maintenance History sheet is empty.")
            return False
        
        # Ensure date column is datetime type
        df['date'] = pd.to_datetime(df['date'])
        
        # Apply date range filter
        if from_date:
            from_date = parse_date(from_date)
            df = df[df['date'] >= from_date]
        
        if to_date:
            to_date = parse_date(to_date)
            df = df[df['date'] <= to_date]
        
        # Apply action filter (case-insensitive partial match)
        if action:
            df = df[df['action_taken'].str.contains(action, case=False, na=False)]
        
        # Check if any records found
        if df.empty:
            print("No records found matching the filter criteria.")
            return False
        
        # Format date column for display
        df_display = df.copy()
        df_display['date'] = df_display['date'].dt.strftime('%Y-%m-%d')
        
        # Print results
        record_count = len(df)
        print(f"\nFound {record_count} maintenance record{'s' if record_count != 1 else ''}")
        print_filters(from_date, to_date, action)
        
        # Print table
        print("\n" + tabulate(df_display, headers='keys', tablefmt='grid', showindex=False))
        
        # Export if requested
        if export:
            export_path = f"data/maintenance_query_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df.to_excel(export_path, index=False)
            print(f"\nExported query results to {export_path}")
            logging.info(f"Exported {record_count} records to {export_path}")
        
        return True
    
    except Exception as e:
        print(f"Error querying maintenance records: {e}")
        logging.error(f"Error querying maintenance records: {e}")
        return False

def print_filters(from_date, to_date, action):
    """Print the filters applied to the query"""
    print("Filters applied:")
    
    if from_date:
        print(f"  From: {from_date.strftime('%Y-%m-%d')}")
    else:
        print("  From: [beginning]")
        
    if to_date:
        print(f"  To: {to_date.strftime('%Y-%m-%d')}")
    else:
        print("  To: [present]")
        
    if action:
        print(f"  Action: \"{action}\"")
    else:
        print("  Action: [all]")

def main():
    """
    Process command-line arguments and call the query function.
    """
    parser = argparse.ArgumentParser(
        description="Query maintenance records by date range and action type."
    )
    parser.add_argument("--from", dest="from_date", 
                        help="Start date in YYYY-MM-DD format")
    parser.add_argument("--to", dest="to_date", 
                        help="End date in YYYY-MM-DD format")
    parser.add_argument("--action", 
                        help="Filter by action taken (case-insensitive partial match)")
    parser.add_argument("--export", action="store_true", 
                        help="Export results to Excel")
    
    args = parser.parse_args()
    
    # Validate dates if provided
    if args.from_date:
        try:
            parse_date(args.from_date)
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(1)
    
    if args.to_date:
        try:
            parse_date(args.to_date)
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(1)
    
    # Ensure at least one filter is provided
    if not any([args.from_date, args.to_date, args.action]):
        print("Warning: No filters specified. This will return all maintenance records.")
        confirmation = input("Do you want to continue? (y/N): ")
        if confirmation.lower() not in ['y', 'yes']:
            print("Query cancelled.")
            sys.exit(0)
    
    success = query_maintenance(args.from_date, args.to_date, args.action, args.export)
    
    # Provide appropriate exit code for scripting purposes
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()