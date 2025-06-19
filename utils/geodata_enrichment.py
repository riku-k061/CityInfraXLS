# utils/geodata_enrichment.py
import os
import pandas as pd
import json
from pathlib import Path
from datetime import datetime
from utils.geodata_handler import GeodataHandler

class GeodataEnrichmentManager:
    def __init__(self):
        self.geodata_handler = GeodataHandler()
        self.schema_path = Path(__file__).parent.parent / 'geodata_schema.json'
        self.data_path = Path(__file__).parent.parent / 'data' / 'asset_geodata.xlsx'
        
        # Load schema for validation
        with open(self.schema_path, 'r') as f:
            self.schema = json.load(f)
        
    def validate_geodata_sheet(self):
        """
        Validate that the geodata sheet exists and has the correct headers.
        
        Returns:
            tuple: (bool, str) - (is_valid, message)
        """
        try:
            # Check if file exists
            if not os.path.exists(self.data_path):
                # Create new file if it doesn't exist
                self.geodata_handler.create_geodata_file()
                return True, "Created new geodata file with required headers"
            
            # Load the file and validate headers
            df = pd.read_excel(self.data_path, sheet_name='AssetGeodata')
            
            # Get required columns from schema
            required_columns = self.schema['required']
            all_columns = list(self.schema['properties'].keys())
            
            # Check all required columns are present
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                return False, f"Missing required columns: {', '.join(missing_columns)}"
            
            # Check for any unexpected columns
            unexpected_columns = [col for col in df.columns if col not in all_columns]
            if unexpected_columns:
                return False, f"Unexpected columns found: {', '.join(unexpected_columns)}"
                
            return True, "Geodata sheet headers validated successfully"
            
        except Exception as e:
            return False, f"Error validating geodata sheet: {e}"
    
    def check_existing_coordinates(self, asset_id):
        """
        Check if coordinates already exist for an asset
        
        Args:
            asset_id (str): The asset ID to check
            
        Returns:
            tuple: (bool, dict) - (has_coordinates, existing_data)
        """
        try:
            # Get existing asset data
            existing_data = self.geodata_handler.get_asset_geodata(asset_id)
            
            if existing_data is None:
                return False, None
                
            # Check if valid lat/long already exists
            has_coordinates = (
                'latitude' in existing_data and 
                'longitude' in existing_data and
                existing_data['latitude'] is not None and
                existing_data['longitude'] is not None and
                not pd.isna(existing_data['latitude']) and
                not pd.isna(existing_data['longitude'])
            )
                
            return has_coordinates, existing_data
            
        except Exception as e:
            print(f"Error checking coordinates: {e}")
            return False, None
    
    def prepare_for_enrichment(self):
        """
        Prepare the system for geodata enrichment by validating the sheet structure
        
        Returns:
            tuple: (bool, str) - (ready, message)
        """
        # Validate sheet headers
        valid, message = self.validate_geodata_sheet()
        if not valid:
            return False, f"Cannot proceed with enrichment: {message}"
        
        # Load assets to ensure we can cross-reference
        try:
            assets_path = Path(__file__).parent.parent / 'data' / 'assets.xlsx'
            if not os.path.exists(assets_path):
                return False, "Assets file not found. Cannot enrich geodata without assets."
                
            assets_df = pd.read_excel(assets_path)
            if 'ID' not in assets_df.columns:
                return False, "Assets file does not contain an ID column."
                
            return True, "System is ready for geodata enrichment"
            
        except Exception as e:
            return False, f"Error preparing for enrichment: {e}"
    
    def enrich_asset_geodata(self, asset_id, latitude, longitude, source="MANUAL_ENTRY", 
                            validation_status="UNVERIFIED", additional_data=None):
        """
        Enrich an asset with geodata, preserving existing valid coordinates unless force_update is True
        
        Args:
            asset_id (str): Asset ID to enrich
            latitude (float): Latitude coordinate
            longitude (float): Longitude coordinate 
            source (str): Source of the geodata
            validation_status (str): Validation status
            additional_data (dict): Any additional geodata fields
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        # Check if coordinates already exist
        has_coords, existing_data = self.check_existing_coordinates(asset_id)
        
        if has_coords:
            return False, f"Asset {asset_id} already has coordinates. Use update_asset_geodata to modify."
        
        # Prepare data for addition
        geodata = {
            'asset_id': asset_id,
            'latitude': latitude,
            'longitude': longitude,
            'geocode_source': source,
            'geocode_timestamp': datetime.now().isoformat(),
            'validation_status': validation_status
        }
        
        # Add any additional provided data
        if additional_data and isinstance(additional_data, dict):
            for key, value in additional_data.items():
                if key in self.schema['properties']:
                    geodata[key] = value
        
        # Add the geodata
        success, message = self.geodata_handler.add_geodata(geodata)
        return success, message
    
    def update_asset_geodata(self, asset_id, geodata_updates, force_update=False):
        """
        Update an asset's geodata, optionally forcing coordinate updates
        
        Args:
            asset_id (str): Asset ID to update
            geodata_updates (dict): Dictionary of fields to update
            force_update (bool): Whether to force updating coordinates
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        # Check if coordinates already exist and handle based on force_update
        has_coords, existing_data = self.check_existing_coordinates(asset_id)
        
        coordinate_fields = ['latitude', 'longitude']
        updates_coordinates = any(field in geodata_updates for field in coordinate_fields)
        
        if has_coords and updates_coordinates and not force_update:
            # Remove coordinate updates if not forcing
            for field in coordinate_fields:
                if field in geodata_updates:
                    del geodata_updates[field]
            
            if not geodata_updates:  # If no updates remain after removing coordinates
                return False, f"No updates to make for asset {asset_id}. Use force_update=True to override coordinates."
        
        # Prepare the complete update data
        update_data = existing_data.copy() if existing_data else {'asset_id': asset_id}
        update_data.update(geodata_updates)
        
        # Update timestamp for the modification
        update_data['geocode_timestamp'] = datetime.now().isoformat()
        
        # Add/update the geodata
        success, message = self.geodata_handler.add_geodata(update_data)
        return success, message
    
    def bulk_check_geodata_status(self, asset_ids=None):
        """
        Check geodata status for multiple assets
        
        Args:
            asset_ids (list): List of asset IDs to check, or None for all
            
        Returns:
            pd.DataFrame: DataFrame with asset_id and geodata status
        """
        try:
            # Load all geodata
            df = pd.read_excel(self.data_path, sheet_name='AssetGeodata')
            
            if asset_ids:
                # Filter to requested assets
                df = df[df['asset_id'].isin(asset_ids)]
            
            # Check which assets have valid coordinates
            df['has_coordinates'] = (~pd.isna(df['latitude'])) & (~pd.isna(df['longitude']))
            
            # Create status summary
            status_df = df[['asset_id', 'has_coordinates', 'validation_status', 'geocode_source']]
            
            return status_df
            
        except Exception as e:
            print(f"Error checking geodata status: {e}")
            return pd.DataFrame(columns=['asset_id', 'has_coordinates', 'validation_status', 'geocode_source'])