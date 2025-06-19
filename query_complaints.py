# query_complaints.py

import os
import sys
import pandas as pd
import argparse
from datetime import datetime
import pytz
from tabulate import tabulate

def query_complaints(
    status=None, 
    department=None, 
    min_rating=None,
    date_from=None, 
    date_to=None, 
    export=False
):
    """
    Query complaints based on filters and optionally export results.
    
    Args:
        status (str): Filter by complaint status (Open/In Progress/Closed)
        department (str): Filter by department
        min_rating (int): Filter by minimum rating (1-5)
        date_from (str): Start date for filtering (YYYY-MM-DD)
        date_to (str): End date for filtering (YYYY-MM-DD)
        export (bool): Whether to export results to Excel file
        
    Returns:
        DataFrame: Filtered complaints data
    """
    # Check if complaints file exists
    complaints_file = "data/complaints.xlsx"
    if not os.path.exists(complaints_file):
        print(f"Error: Complaints file not found at {complaints_file}")
        return None
    
    try:
        # Load complaints data
        df = pd.read_excel(complaints_file, sheet_name='Complaints')
        
        if df.empty:
            print("No complaints found in the database.")
            return df
            
        # Make a copy of original dataframe for filtering
        filtered_df = df.copy()
        
        # Apply filters
        filters_applied = []
        
        # Filter by status
        if status:
            filtered_df = filtered_df[filtered_df['status'] == status]
            filters_applied.append(f"Status: {status}")
            
        # Filter by department
        if department:
            filtered_df = filtered_df[filtered_df['department'] == department]
            filters_applied.append(f"Department: {department}")
            
        # Filter by minimum rating
        if min_rating:
            filtered_df = filtered_df[filtered_df['rating'] >= min_rating]
            filters_applied.append(f"Min Rating: {min_rating}")
            
        # Filter by date range (created_at)
        if date_from:
            filtered_df['created_at'] = pd.to_datetime(filtered_df['created_at'])
            from_date = pd.to_datetime(date_from)
            filtered_df = filtered_df[filtered_df['created_at'] >= from_date]
            filters_applied.append(f"From: {date_from}")
            
        if date_to:
            filtered_df['created_at'] = pd.to_datetime(filtered_df['created_at'])
            to_date = pd.to_datetime(date_to)
            filtered_df = filtered_df[filtered_df['created_at'] <= to_date]
            filters_applied.append(f"To: {date_to}")
        
        # Print filter information
        if filters_applied:
            print("Filters applied:")
            for filter_info in filters_applied:
                print(f"- {filter_info}")
        else:
            print("No filters applied - showing all complaints")
            
        # Display results
        if filtered_df.empty:
            print("\nNo complaints match the specified criteria.")
            return filtered_df
            
        # Format dates for display
        display_df = filtered_df.copy()
        if 'created_at' in display_df.columns:
            display_df['created_at'] = pd.to_datetime(display_df['created_at']).dt.strftime('%Y-%m-%d %H:%M')
        if 'closed_at' in display_df.columns:
            # Convert NaT to empty string for display
            display_df['closed_at'] = pd.to_datetime(display_df['closed_at'])
            display_df['closed_at'] = display_df['closed_at'].fillna('')
            display_df.loc[display_df['closed_at'] != '', 'closed_at'] = display_df.loc[display_df['closed_at'] != '', 'closed_at'].dt.strftime('%Y-%m-%d %H:%M')
        
        print(f"\nFound {len(filtered_df)} complaint(s):")
        
        # Only show key columns in tabulated output for readability
        display_columns = ['complaint_id', 'reporter', 'asset_location', 'department', 'status', 'rating', 'created_at']
        
        print(tabulate(
            display_df[display_columns].head(50), 
            headers='keys', 
            tablefmt='grid', 
            showindex=False
        ))
        
        if len(filtered_df) > 50:
            print(f"\nNote: Only showing first 50 of {len(filtered_df)} results")
        
        # Export if requested
        if export and not filtered_df.empty:
            timestamp = datetime.now(pytz.UTC).strftime('%Y%m%d%H%M%S')
            export_path = f"data/complaints_query_{timestamp}.xlsx"
            
            filtered_df.to_excel(export_path, index=False)
            print(f"\nExported {len(filtered_df)} complaints to {export_path}")
        
        return filtered_df
        
    except Exception as e:
        print(f"Error querying complaints: {str(e)}")
        return None

def main():
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description='Query and filter infrastructure complaints')
    
    parser.add_argument('--status', choices=['Open', 'In Progress', 'Closed'], 
                        help='Filter by complaint status')
    parser.add_argument('--department', type=str, help='Filter by department')
    parser.add_argument('--min-rating', type=int, choices=range(1, 6),
                        help='Filter by minimum rating (1-5)')
    parser.add_argument('--from', dest='date_from', type=str, 
                        help='Start date for filtering (YYYY-MM-DD)')
    parser.add_argument('--to', dest='date_to', type=str,
                        help='End date for filtering (YYYY-MM-DD)')
    parser.add_argument('--export', action='store_true',
                        help='Export results to Excel file')
    
    args = parser.parse_args()
    
    # Call the query function with provided arguments
    query_complaints(
        status=args.status,
        department=args.department,
        min_rating=args.min_rating,
        date_from=args.date_from,
        date_to=args.date_to,
        export=args.export
    )

if __name__ == "__main__":
    main()