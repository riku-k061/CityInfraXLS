# geodata_handler.py
import json
import os
import pandas as pd
from datetime import datetime
import jsonschema
from jsonschema import validate
from pathlib import Path

class GeodataHandler:
    def __init__(self):
        schema_path = 'geodata_schema.json'
        with open(schema_path, 'r') as f:
            self.schema = json.load(f)

        # ensure the data directory exists
        self.data_path = Path(__file__).parent.parent / 'data' / 'asset_geodata.xlsx'
        os.makedirs(self.data_path.parent, exist_ok=True)

        # Create the geodata Excel file if it doesn't yet exist
        if not self.data_path.exists():
            self.create_geodata_file()
    
    def create_geodata_file(self):
        """Create the initial geodata Excel file with proper columns"""
        columns = list(self.schema['properties'].keys())
        df = pd.DataFrame(columns=columns)
        
        # Add data type metadata
        dtype_row = {}
        for col in columns:
            if self.schema['properties'][col]['type'] == 'number':
                dtype_row[col] = 'float'
            elif self.schema['properties'][col].get('format') == 'date-time':
                dtype_row[col] = 'datetime'
            else:
                dtype_row[col] = 'string'
        
        # Create metadata sheet
        metadata = pd.DataFrame([dtype_row], index=['data_type'])
        
        # Create enum constraints sheet
        enum_data = {}
        for col, prop in self.schema['properties'].items():
            if 'enum' in prop:
                enum_data[col] = prop['enum']
        
        # Add constraints as a separate sheet if any exist
        constraints_df = None
        if enum_data:
            # Convert enums to a dataframe format
            max_enum_length = max(len(v) for v in enum_data.values())
            for k, v in enum_data.items():
                v.extend([''] * (max_enum_length - len(v)))
            constraints_df = pd.DataFrame(enum_data)
        
        # Save to Excel with multiple sheets
        with pd.ExcelWriter(self.data_path) as writer:
            df.to_excel(writer, sheet_name='AssetGeodata', index=False)
            metadata.to_excel(writer, sheet_name='Metadata')
            if constraints_df is not None:
                constraints_df.to_excel(writer, sheet_name='AllowedValues', index=False)
    
    def validate_geodata(self, geodata):
        """
        Validate geodata against the schema
        
        Args:
            geodata (dict): Dictionary containing asset geodata
            
        Returns:
            tuple: (bool, str) - (is_valid, error_message)
        """
        try:
            validate(instance=geodata, schema=self.schema)
            return True, ""
        except jsonschema.exceptions.ValidationError as e:
            return False, str(e)
    
    def add_geodata(self, asset_data):
        """
        Add geodata for an asset
        
        Args:
            asset_data (dict): Dictionary containing asset geodata
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        # Validate the data
        is_valid, error = self.validate_geodata(asset_data)
        if not is_valid:
            return False, f"Validation failed: {error}"
        
        # Ensure timestamp is provided or generate one
        if 'geocode_timestamp' not in asset_data:
            asset_data['geocode_timestamp'] = datetime.now().isoformat()
        
        # Set default validation status if not provided
        if 'validation_status' not in asset_data:
            asset_data['validation_status'] = "UNVERIFIED"
        
        # Load existing data
        try:
            df = pd.read_excel(self.data_path, sheet_name='AssetGeodata')
        except:
            # If there's an error loading, create a new file
            self.create_geodata_file()
            df = pd.read_excel(self.data_path, sheet_name='AssetGeodata')
        
        # Check if the asset already has geodata
        existing = df[df['asset_id'] == asset_data['asset_id']]
        if not existing.empty:
            # Update existing record
            for key, value in asset_data.items():
                df.loc[df['asset_id'] == asset_data['asset_id'], key] = value
        else:
            # Add new record
            df = df.append(asset_data, ignore_index=True)
        
        # Save back to Excel
        with pd.ExcelWriter(self.data_path) as writer:
            df.to_excel(writer, sheet_name='AssetGeodata', index=False)
            
            # Preserve other sheets
            for sheet_name in ['Metadata', 'AllowedValues']:
                try:
                    sheet_df = pd.read_excel(self.data_path, sheet_name=sheet_name)
                    sheet_df.to_excel(writer, sheet_name=sheet_name)
                except:
                    pass
        
        return True, "Geodata added successfully"
    
    def get_asset_geodata(self, asset_id):
        """
        Retrieve geodata for a specific asset
        
        Args:
            asset_id (str): The asset ID to look up
            
        Returns:
            dict: Asset geodata or None if not found
        """
        try:
            df = pd.read_excel(self.data_path, sheet_name='AssetGeodata')
            asset_data = df[df['asset_id'] == asset_id]
            
            if asset_data.empty:
                return None
                
            # Convert to dictionary
            return asset_data.iloc[0].to_dict()
        except Exception as e:
            print(f"Error retrieving asset geodata: {e}")
            return None