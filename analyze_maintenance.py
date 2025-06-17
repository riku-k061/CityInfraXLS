# analyze_maintenance.py

import os
import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
from utils.excel_handler import create_maintenance_history_sheet

def analyze_maintenance(history_path='data/maintenance_history.xlsx', export=True):
    """
    Analyze maintenance records to identify patterns by asset.
    
    Args:
        history_path (str): Path to the maintenance history Excel file
        export (bool): Whether to export results to Excel
        
    Returns:
        pd.DataFrame: Analysis results
    """
    # Check if history file exists
    if not os.path.exists(history_path):
        print(f"No maintenance history found at {history_path}")
        print("Creating empty maintenance history file...")
        create_maintenance_history_sheet(history_path)
        print("Please log some maintenance records before analysis.")
        return None
    
    # Load maintenance history
    try:
        df = pd.read_excel(history_path, sheet_name="Maintenance History")
        if df.empty:
            print("Maintenance history is empty. Please log some records first.")
            return None
        
        print(f"Loaded {len(df)} maintenance records for analysis.")
    except Exception as e:
        print(f"Error loading maintenance history: {e}")
        return None
    
    # Convert date strings to datetime objects
    df['date'] = pd.to_datetime(df['date'])
    
    # Sort by asset_id and date
    df = df.sort_values(['asset_id', 'date'])
    
    # Group by asset_id
    asset_groups = df.groupby('asset_id')
    
    # Initialize results DataFrame
    results = []
    
    # Process each asset
    for asset_id, group in asset_groups:
        # Count total records
        record_count = len(group)
        
        # Calculate date intervals (in days)
        if record_count > 1:
            # Calculate the difference between consecutive dates
            dates = group['date'].sort_values().reset_index(drop=True)
            intervals = [(dates[i] - dates[i-1]).days for i in range(1, len(dates))]
            avg_interval = sum(intervals) / len(intervals)
            min_interval = min(intervals) if intervals else np.nan
            max_interval = max(intervals) if intervals else np.nan
        else:
            avg_interval = np.nan
            min_interval = np.nan
            max_interval = np.nan
        
        # Get first and last maintenance dates
        first_date = group['date'].min()
        last_date = group['date'].max()
        
        # Calculate total time span
        time_span = (last_date - first_date).days if record_count > 1 else 0
        
        # Calculate frequency (records per year)
        if time_span > 0:
            frequency = (record_count / time_span) * 365
        else:
            frequency = np.nan
        
        # Calculate total cost and average cost
        total_cost = group['cost'].sum()
        avg_cost = total_cost / record_count if record_count > 0 else 0
        
        # Count maintenance actions by type
        action_counts = group['action_taken'].value_counts().to_dict()
        inspections = action_counts.get('Inspection', 0)
        repairs = action_counts.get('Repair', 0)
        replacements = action_counts.get('Replacement', 0)
        
        # Add to results
        results.append({
            'asset_id': asset_id,
            'record_count': record_count,
            'first_maintenance': first_date,
            'last_maintenance': last_date,
            'time_span_days': time_span,
            'avg_interval_days': avg_interval,
            'min_interval_days': min_interval,
            'max_interval_days': max_interval,
            'maintenance_frequency': frequency,
            'total_cost': total_cost,
            'avg_cost_per_maintenance': avg_cost,
            'inspections': inspections,
            'repairs': repairs,
            'replacements': replacements
        })
    
    # Create results DataFrame
    results_df = pd.DataFrame(results)
    
    # Sort by maintenance frequency (descending)
    results_df = results_df.sort_values('maintenance_frequency', ascending=False)
    
    # Print top five assets by frequency
    print("\nTop 5 Assets by Maintenance Frequency:")
    print("========================================")
    top_five = results_df.head(5)
    for i, (_, asset) in enumerate(top_five.iterrows(), 1):
        print(f"{i}. Asset: {asset['asset_id']}")
        print(f"   Records: {asset['record_count']}")
        
        if not np.isnan(asset['avg_interval_days']):
            print(f"   Average interval: {asset['avg_interval_days']:.1f} days")
        else:
            print("   Average interval: N/A (only one record)")
            
        print(f"   Total cost: ${asset['total_cost']:.2f}")
        print(f"   Actions: {asset['inspections']} inspections, {asset['repairs']} repairs, {asset['replacements']} replacements")
        print()
    
    # Export results if requested
    if export and not results_df.empty:
        try:
            with pd.ExcelWriter(history_path, engine='openpyxl', mode='a') as writer:
                # Check if "Maintenance Analysis" sheet already exists
                if "Maintenance Analysis" in pd.ExcelFile(history_path).sheet_names:
                    # If exists, we need to remove it first (can't just overwrite with ExcelWriter)
                    book = writer.book
                    if "Maintenance Analysis" in book.sheetnames:
                        std_idx = book.sheetnames.index("Maintenance Analysis")
                        book.remove(book.worksheets[std_idx])
                        
                # Write the new analysis
                results_df.to_excel(writer, sheet_name="Maintenance Analysis", index=False)
                print(f"Analysis exported to {history_path}, sheet 'Maintenance Analysis'")
        except Exception as e:
            print(f"Error exporting analysis: {e}")
    
    return results_df

if __name__ == "__main__":
    results = analyze_maintenance()
    if results is not None:
        print("Analysis complete.")