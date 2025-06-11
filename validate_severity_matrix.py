# validate_severity_matrix.py
import json
import os
from typing import Dict, Any

def validate_severity_matrix(file_path: str) -> Dict[str, Dict[str, Any]]:
    """
    Validates the severity matrix JSON file.
    
    Args:
        file_path: Path to the severity matrix JSON file
    
    Returns:
        The validated severity matrix as a dictionary
    
    Raises:
        FileNotFoundError: If the file doesn't exist
        ValueError: If the file is malformed or missing required fields
    """
    # Check if file exists
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Severity matrix file not found at {file_path}")
    
    # Load JSON file
    try:
        with open(file_path, 'r') as file:
            severity_matrix = json.load(file)
    except json.JSONDecodeError:
        raise ValueError(f"Invalid JSON format in severity matrix file: {file_path}")
    
    # Validate structure
    if not isinstance(severity_matrix, dict):
        raise ValueError("Severity matrix must be a dictionary/object")
    
    # Check each severity level
    required_fields = ["hours", "description", "priority"]
    
    for severity_level, details in severity_matrix.items():
        # Ensure each severity level is a dictionary with all required fields
        if not isinstance(details, dict):
            raise ValueError(f"Severity level '{severity_level}' must contain a dictionary of details")
        
        # Check for required fields
        for field in required_fields:
            if field not in details:
                raise ValueError(f"Missing required field '{field}' for severity level '{severity_level}'")
        
        # Validate field types
        if not isinstance(details["hours"], (int, float)):
            raise ValueError(f"'hours' must be a number for severity level '{severity_level}'")
            
        if not isinstance(details["description"], str):
            raise ValueError(f"'description' must be a string for severity level '{severity_level}'")
            
        if not isinstance(details["priority"], int):
            raise ValueError(f"'priority' must be an integer for severity level '{severity_level}'")
            
    return severity_matrix

if __name__ == "__main__":
    # Example usage
    try:
        matrix = validate_severity_matrix("severity_matrix.json")
        print("Severity matrix validation successful!")
        print(f"Found {len(matrix)} severity levels: {', '.join(matrix.keys())}")
    except (FileNotFoundError, ValueError) as e:
        print(f"Validation error: {e}")