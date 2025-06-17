# record_budget.py

"""
record_budget.py - Tool for recording new budget allocations in CityInfraXLS

This script prompts users for required budget allocation information,
validates the input against budget_allocation_schema.json, and
appends the new record to the budget allocations Excel sheet.
"""

import os
import json
import argparse
import datetime
from decimal import Decimal, DecimalException
import re
import pandas as pd
from utils.excel_handler import create_sheets_from_schema, load_workbook, save_workbook

def validate_fiscal_year(year_str):
    """Validate fiscal year format YYYY-YYYY."""
    if not re.match(r'^\d{4}-\d{4}$', year_str):
        raise ValueError("Fiscal year must be in format YYYY-YYYY")
    
    start_year, end_year = map(int, year_str.split('-'))
    if end_year != start_year + 1:
        raise ValueError("End year must be one year after start year")
    
    return year_str

def validate_project_id(project_id):
    """Validate project ID format PRJ-XXXXXX."""
    if not re.match(r'^PRJ-\d{6}$', project_id):
        raise ValueError("Project ID must be in format PRJ-XXXXXX where X is a digit")
    return project_id

def load_schema():
    """Load budget allocation schema from JSON file."""
    schema_path = os.path.join(os.path.dirname(__file__), "budget_allocation_schema.json")
    with open(schema_path, 'r') as f:
        return json.load(f)

def prompt_for_department():
    """Prompt user for department name."""
    while True:
        department = input("Department name: ").strip()
        if len(department) < 2 or len(department) > 100:
            print("Department name must be between 2 and 100 characters.")
            continue
        return department

def prompt_for_category(schema):
    """Prompt user for budget category."""
    categories = schema["properties"]["category"]["enum"]
    print("Available categories:")
    for i, category in enumerate(categories, 1):
        print(f"{i}. {category}")
    
    while True:
        choice = input("Select category (number): ").strip()
        try:
            index = int(choice) - 1
            if 0 <= index < len(categories):
                return categories[index]
            else:
                print("Invalid selection. Please try again.")
        except ValueError:
            print("Please enter a number.")

def prompt_for_status(schema):
    """Prompt user for budget status."""
    statuses = schema["properties"]["status"]["enum"]
    print("Available statuses:")
    for i, status in enumerate(statuses, 1):
        print(f"{i}. {status}")
    
    while True:
        choice = input("Select status (number): ").strip()
        try:
            index = int(choice) - 1
            if 0 <= index < len(statuses):
                return statuses[index]
            else:
                print("Invalid selection. Please try again.")
        except ValueError:
            print("Please enter a number.")

def append_to_excel(excel_path, sheet_name, record):
    """
    Append a record to an Excel file.
    
    Args:
        excel_path (str): Path to the Excel file
        sheet_name (str): Name of the sheet
        record (dict): Record to append
    """
    # Check if file exists
    if not os.path.exists(excel_path):
        # Get schema path from the record's type
        schema_path = os.path.join(os.path.dirname(__file__), "budget_allocation_schema.json")
        # Create Excel file with schema
        create_sheets_from_schema(schema_path, excel_path, sheet_name)
    
    # Load existing workbook
    wb = load_workbook(excel_path)
    
    # Get the worksheet
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        raise ValueError(f"Sheet {sheet_name} not found in {excel_path}")
    
    # Get headers from the first row
    headers = [cell.value for cell in ws[1]]
    
    # Create new row with values from record
    new_row = []
    for header in headers:
        new_row.append(record.get(header, ""))
    
    # Append row to worksheet
    ws.append(new_row)
    
    # Save the workbook
    save_workbook(wb, excel_path)
    print(f"Record appended to {excel_path}, sheet {sheet_name}")

def main():
    parser = argparse.ArgumentParser(description="Record a new budget allocation")
    parser.add_argument("--batch", action="store_true", help="Run in batch mode without prompts")
    args = parser.parse_args()
    
    # Define paths
    excel_path = "data/budget_allocations.xlsx"
    schema_path = "budget_allocation_schema.json"
    sheet_name = "Allocations"
    
    # Load schema for validation
    schema = load_schema()
    
    # Initialize budget data with today's date
    budget_data = {
        "allocation_date": datetime.date.today().isoformat()
    }
    
    try:
        # Department
        if not args.batch:
            budget_data["department"] = prompt_for_department()
        else:
            budget_data["department"] = input()
        
        # Fiscal year
        while True:
            try:
                year_input = input("Fiscal year (YYYY-YYYY): ").strip()
                budget_data["fiscal_year"] = validate_fiscal_year(year_input)
                break
            except ValueError as e:
                print(f"Error: {str(e)}")
        
        # Allocated amount
        while True:
            try:
                amount = input("Allocated amount ($): ").strip()
                amount_decimal = Decimal(amount)
                if amount_decimal <= 0:
                    print("Amount must be positive.")
                    continue
                budget_data["allocated_amount"] = float(amount_decimal)
                break
            except (ValueError, DecimalException):
                print("Please enter a valid number.")
        
        # Project ID
        while True:
            try:
                project_id = input("Project ID (PRJ-XXXXXX): ").strip()
                budget_data["project_id"] = validate_project_id(project_id)
                break
            except ValueError as e:
                print(f"Error: {str(e)}")
        
        # Category
        budget_data["category"] = prompt_for_category(schema)
        
        # Status
        budget_data["status"] = prompt_for_status(schema)
        
        # Notes (optional)
        notes = input("Notes (optional): ").strip()
        if notes:
            budget_data["notes"] = notes
        
        # Approving authority
        approving_auth = input("Approving authority: ").strip()
        if approving_auth:
            budget_data["approving_authority"] = approving_auth
        
        # Append to Excel file
        append_to_excel(excel_path, sheet_name, budget_data)
        
        print(f"Budget allocation of ${budget_data['allocated_amount']} "
              f"for {budget_data['department']} successfully recorded.")
              
    except KeyboardInterrupt:
        print("\nOperation canceled.")
        return 1
    except Exception as e:
        print(f"Error: {str(e)}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())