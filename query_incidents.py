# query_incidents.py

"""
Script for querying and analyzing infrastructure incidents.
Provides filtering, grouping, statistics, and export capabilities.
"""

import os
import sys
import argparse
import logging
import pandas as pd
import numpy as np
import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import matplotlib.pyplot as plt
from tabulate import tabulate

# Configure logging
logging.basicConfig(
    filename='cityinfraxls.log',
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger('CityInfraXLS')

def parse_args():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(description='Query and analyze infrastructure incidents')
    
    parser.add_argument('--overdue', action='store_true', 
                        help='Show only incidents where Elapsed Hours > SLA Deadline')
    
    parser.add_argument('--group-by', choices=['severity', 'status', 'type'],
                        help='Group results by the specified field')
    
    parser.add_argument('--stats', action='store_true',
                        help='Display statistics including counts and SLA compliance rates')
    
    parser.add_argument('--export', type=str, metavar='FILENAME',
                        help='Export results to the specified Excel file')
    
    parser.add_argument('--days', type=int, default=None,
                        help='Filter to incidents reported within the last N days')
    
    return parser.parse_args()

def load_incidents_data() -> pd.DataFrame:
    """
    Load incidents data from Excel and prepare for analysis
    
    Returns:
        DataFrame containing incident data with calculated fields
    """
    incident_file = 'data/incidents.xlsx'
    
    if not os.path.exists(incident_file):
        logger.error(f"Incidents file not found: {incident_file}")
        print(f"ERROR: Incidents file not found at {incident_file}")
        sys.exit(1)
    
    try:
        # Load the data
        df = pd.read_excel(incident_file)
        
        # Handle empty dataframe
        if df.empty:
            logger.info("No incidents found in the incidents file")
            print("No incidents found in the database.")
            sys.exit(0)
        
        # Convert date columns to datetime if they're not already
        for col in ['Reported At', 'SLA Deadline']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Calculate elapsed hours if not already present
        if 'Elapsed Hours' not in df.columns or df['Elapsed Hours'].isna().all():
            current_time = datetime.datetime.now()
            df['Elapsed Hours'] = (current_time - df['Reported At']).dt.total_seconds() / 3600
        
        # Determine SLA compliance
        df['Is Overdue'] = df['Elapsed Hours'] > (df['SLA Deadline'] - df['Reported At']).dt.total_seconds() / 3600
        df['SLA Compliant'] = ~df['Is Overdue'] | (df['Status'] == 'Closed')
        
        return df
        
    except Exception as e:
        logger.error(f"Error loading incidents data: {e}")
        print(f"ERROR: Failed to load incidents data - {e}")
        sys.exit(1)

def filter_data(df: pd.DataFrame, args: argparse.Namespace) -> pd.DataFrame:
    """
    Filter the data based on command line arguments
    
    Args:
        df: DataFrame with incidents data
        args: Command line arguments
        
    Returns:
        Filtered DataFrame
    """
    filtered_df = df.copy()
    filter_applied = False
    
    # Apply overdue filter if requested
    if args.overdue:
        filtered_df = filtered_df[filtered_df['Is Overdue']]
        filter_applied = True
        logger.info(f"Filtered to show only overdue incidents: {len(filtered_df)} found")
    
    # Apply days filter if requested
    if args.days:
        cutoff_date = datetime.datetime.now() - datetime.timedelta(days=args.days)
        filtered_df = filtered_df[filtered_df['Reported At'] >= cutoff_date]
        filter_applied = True
        logger.info(f"Filtered to incidents from last {args.days} days: {len(filtered_df)} found")
    
    if filter_applied and filtered_df.empty:
        print("No incidents match the specified filters.")
        sys.exit(0)
        
    return filtered_df

def calculate_statistics(df: pd.DataFrame, group_by: Optional[str] = None) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Calculate statistics for the incidents data
    
    Args:
        df: DataFrame with incidents data
        group_by: Optional field to group by
        
    Returns:
        Tuple of (summary_stats, detailed_stats) DataFrames
    """
    # Define the grouping field or use a dummy grouper if none specified
    if group_by and group_by in df.columns:
        grouper = df[group_by]
    else:
        df['All Incidents'] = 'All'
        grouper = df['All Incidents']
        group_by = 'All Incidents'
    
    # Calculate basic statistics
    stats = {
        'Total Incidents': grouper.count(),
        'Open Incidents': df[df['Status'] == 'Open'].groupby(grouper).size(),
        'Closed Incidents': df[df['Status'] == 'Closed'].groupby(grouper).size(),
        'Overdue Incidents': df[df['Is Overdue']].groupby(grouper).size(),
        'SLA Compliant (%)': df.groupby(grouper)['SLA Compliant'].mean() * 100
    }
    
    # Convert to DataFrame
    summary_stats = pd.DataFrame(stats)
    summary_stats.fillna(0, inplace=True)
    summary_stats = summary_stats.astype({
        'Total Incidents': int,
        'Open Incidents': int,
        'Closed Incidents': int,
        'Overdue Incidents': int,
    })
    
    # Calculate mean response time and other metrics
    detailed_stats = df.groupby(grouper).agg({
        'Elapsed Hours': ['count', 'mean', 'min', 'max'],
        'SLA Compliant': ['mean', 'sum']
    })
    
    detailed_stats.columns = [
        'Count', 'Avg Hours', 'Min Hours', 'Max Hours', 
        'SLA Compliance Rate', 'Compliant Count'
    ]
    
    detailed_stats['Avg Hours'] = detailed_stats['Avg Hours'].round(2)
    detailed_stats['SLA Compliance Rate'] = (detailed_stats['SLA Compliance Rate'] * 100).round(2)
    detailed_stats['Compliant Count'] = detailed_stats['Compliant Count'].astype(int)
    
    return summary_stats, detailed_stats

def create_dashboard(df: pd.DataFrame, summary_stats: pd.DataFrame, 
                    detailed_stats: pd.DataFrame, filename: str) -> None:
    """
    Create an Excel dashboard with incident statistics
    
    Args:
        df: Original filtered DataFrame
        summary_stats: Summary statistics DataFrame
        detailed_stats: Detailed statistics DataFrame
        filename: Path to save the Excel file
    """
    logger.info(f"Creating Excel dashboard at {filename}")
    
    # Create a Pandas Excel writer using XlsxWriter
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    
    # Write each DataFrame to a different worksheet
    df.to_excel(writer, sheet_name='Raw Data', index=False)
    summary_stats.to_excel(writer, sheet_name='Summary Stats')
    detailed_stats.to_excel(writer, sheet_name='Detailed Stats')
    
    # Access the XlsxWriter workbook and worksheet objects
    workbook = writer.book
    
    # Create a chart sheet for the summary visualization
    chart_sheet = workbook.add_worksheet('Dashboard')
    
    # Create charts for the dashboard
    create_summary_charts(workbook, chart_sheet, summary_stats)
    
    # Create compliance chart
    create_compliance_chart(workbook, chart_sheet, summary_stats)
    
    # Close the Pandas Excel writer and save the Excel file
    writer.close()
    
    print(f"\nDashboard exported to {filename}")

def create_summary_charts(workbook, chart_sheet, summary_stats):
    """Create incident count chart for the dashboard"""
    # Create a column chart for incident counts
    chart = workbook.add_chart({'type': 'column'})
    
    # Configure the series for the chart from the summary stats
    num_rows = len(summary_stats)
    
    # Add series for each incident type
    chart.add_series({
        'name': 'Total Incidents',
        'categories': ['Summary Stats', 1, 0, num_rows, 0],
        'values': ['Summary Stats', 1, 1, num_rows, 1],
    })
    
    chart.add_series({
        'name': 'Open Incidents',
        'categories': ['Summary Stats', 1, 0, num_rows, 0],
        'values': ['Summary Stats', 1, 2, num_rows, 2],
    })
    
    chart.add_series({
        'name': 'Overdue Incidents',
        'categories': ['Summary Stats', 1, 0, num_rows, 0],
        'values': ['Summary Stats', 1, 4, num_rows, 4],
    })
    
    # Add chart title and labels
    chart.set_title({'name': 'Incident Counts by Category'})
    chart.set_x_axis({'name': 'Category'})
    chart.set_y_axis({'name': 'Count'})
    
    # Set an Excel chart style
    chart.set_style(11)
    chart.set_size({'width': 720, 'height': 400})
    
    # Insert the chart into the chart sheet
    chart_sheet.insert_chart('A1', chart)

def create_compliance_chart(workbook, chart_sheet, summary_stats):
    """Create SLA compliance chart for the dashboard"""
    # Create a column chart for SLA compliance
    chart = workbook.add_chart({'type': 'column'})
    
    # Configure the series for the chart from the summary stats
    num_rows = len(summary_stats)
    
    # Add series for SLA compliance
    chart.add_series({
        'name': 'SLA Compliance Rate (%)',
        'categories': ['Summary Stats', 1, 0, num_rows, 0],
        'values': ['Summary Stats', 1, 5, num_rows, 5],
        'data_labels': {'value': True, 'num_format': '0.0%'},
    })
    
    # Add chart title and labels
    chart.set_title({'name': 'SLA Compliance Rate by Category'})
    chart.set_x_axis({'name': 'Category'})
    chart.set_y_axis({
        'name': 'Compliance Rate (%)',
        'min': 0,
        'max': 100,
    })
    
    # Set an Excel chart style
    chart.set_style(42)  # Different style for visual distinction
    chart.set_size({'width': 720, 'height': 400})
    
    # Insert the chart into the chart sheet below the first chart
    chart_sheet.insert_chart('A25', chart)

def display_results(df: pd.DataFrame, args: argparse.Namespace) -> None:
    """
    Display the results in the console
    
    Args:
        df: DataFrame with incidents data
        args: Command line arguments
    """
    if args.group_by and args.stats:
        # Calculate and display grouped statistics
        summary_stats, detailed_stats = calculate_statistics(df, args.group_by)
        
        print("\n=== Summary Statistics ===")
        print(tabulate(summary_stats, headers='keys', tablefmt='grid', showindex=True))
        
        print("\n=== Detailed Statistics ===")
        print(tabulate(detailed_stats, headers='keys', tablefmt='grid', showindex=True))
        
    elif args.stats:
        # Calculate and display overall statistics
        summary_stats, detailed_stats = calculate_statistics(df)
        
        print("\n=== Summary Statistics ===")
        print(tabulate(summary_stats, headers='keys', tablefmt='grid', showindex=True))
        
    elif args.group_by:
        # Group and display incidents
        for name, group in df.groupby(args.group_by):
            print(f"\n=== {args.group_by}: {name} ({len(group)} incidents) ===")
            display_df = group[['Incident ID', 'Asset ID', 'Type', 'Severity', 
                               'Reported At', 'Status', 'Elapsed Hours']]
            print(tabulate(display_df, headers='keys', tablefmt='grid', showindex=False))
    else:
        # Display all incidents
        display_df = df[['Incident ID', 'Asset ID', 'Type', 'Severity', 
                         'Reported At', 'Status', 'Elapsed Hours']]
        print("\n=== Incident List ===")
        print(tabulate(display_df, headers='keys', tablefmt='grid', showindex=False))
    
    # Export if requested
    if args.export:
        if args.stats or args.group_by:
            summary_stats, detailed_stats = calculate_statistics(df, args.group_by if args.group_by else None)
            create_dashboard(df, summary_stats, detailed_stats, args.export)
        else:
            # Simple export without dashboard
            df.to_excel(args.export, index=False)
            print(f"\nData exported to {args.export}")

def main():
    """Main function to run the incident query script"""
    # Parse command line arguments
    args = parse_args()
    
    # Load incidents data
    df = load_incidents_data()
    
    # Apply filters
    filtered_df = filter_data(df, args)
    
    # Display results
    display_results(filtered_df, args)
    
    # Log the query
    filter_desc = []
    if args.overdue:
        filter_desc.append("overdue only")
    if args.days:
        filter_desc.append(f"last {args.days} days")
    if args.group_by:
        filter_desc.append(f"grouped by {args.group_by}")
    if args.stats:
        filter_desc.append("with statistics")
        
    filter_str = ", ".join(filter_desc) if filter_desc else "no filters"
    logger.info(f"Queried incidents with {filter_str}; found {len(filtered_df)} matching incidents")

if __name__ == "__main__":
    main()