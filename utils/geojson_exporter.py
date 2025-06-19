# utils/geojson_exporter.py
import json
import os
import pandas as pd
from datetime import datetime
from pathlib import Path

class GeoJSONExporter:
    """Creates GeoJSON layers from geodata sheet records that are In Bounds."""
    
    def __init__(self, geodata_file_path, output_dir="data/geojson"):
        """Initialize with path to geodata file and output directory."""
        self.geodata_file_path = geodata_file_path
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def export_in_bounds_features(self, include_attributes=None):
        """
        Export all In Bounds geodata entries to a GeoJSON file.
        
        Args:
            include_attributes: List of additional columns to include as properties
            
        Returns:
            dict: Information about the export (file_path, feature_count)
        """
        # Read geodata Excel file
        geodata_df = pd.read_excel(self.geodata_file_path)
        
        # Filter to only In Bounds entries
        in_bounds_df = geodata_df[geodata_df['validation_status'] == 'IN_BOUNDS']
        
        if in_bounds_df.empty:
            return {"error": "No In Bounds features found"}
            
        # Create GeoJSON structure
        features = []
        for _, row in in_bounds_df.iterrows():
            # Create properties dict with required fields
            properties = {
                "asset_id": row["asset_id"],
                "validation_status": row["validation_status"]
            }
            
            # Add risk category if it exists
            if "risk_category" in row:
                properties["risk_category"] = row["risk_category"]
                
            # Add any additional requested properties
            if include_attributes:
                for attr in include_attributes:
                    if attr in row and not pd.isna(row[attr]):
                        properties[attr] = row[attr]
            
            # Create GeoJSON feature
            feature = {
                "type": "Feature",
                "geometry": {
                    "type": "Point",
                    "coordinates": [float(row["longitude"]), float(row["latitude"])]
                },
                "properties": properties
            }
            features.append(feature)
        
        # Assemble full GeoJSON structure
        geojson = {
            "type": "FeatureCollection",
            "features": features
        }
        
        # Generate filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        filename = f"in_bounds_assets_{timestamp}.geojson"
        file_path = self.output_dir / filename
        
        # Write to file
        with open(file_path, 'w') as f:
            json.dump(geojson, f)
        
        # Update geodata sheet with metadata about this export
        self._update_geodata_sheet(str(file_path), len(features))
        
        return {
            "file_path": str(file_path),
            "feature_count": len(features),
            "timestamp": timestamp
        }
    
    def _update_geodata_sheet(self, geojson_path, feature_count):
        """
        Update the geodata sheet with information about the export.
        
        Adds a new sheet or updates existing sheet with export metadata.
        """
        try:
            # Load the Excel file
            with pd.ExcelFile(self.geodata_file_path) as xls:
                sheet_names = xls.sheet_names
                
                # Check if exports sheet exists
                if "geojson_exports" in sheet_names:
                    exports_df = pd.read_excel(xls, "geojson_exports")
                else:
                    # Create new DataFrame for exports sheet
                    exports_df = pd.DataFrame(columns=[
                        "export_date", "file_path", "feature_count", "schema_version"
                    ])
            
            # Add new export record
            new_export = {
                "export_date": datetime.now(),
                "file_path": geojson_path, 
                "feature_count": feature_count,
                "schema_version": "1.0"  # Track schema version for future compatibility
            }
            exports_df = pd.concat([exports_df, pd.DataFrame([new_export])], ignore_index=True)
            
            # Write back to Excel file, preserving other sheets
            with pd.ExcelWriter(self.geodata_file_path, engine='openpyxl', mode='a') as writer:
                # Remove sheet if it exists
                if "geojson_exports" in writer.book.sheetnames:
                    std = writer.book["geojson_exports"]
                    writer.book.remove(std)
                
                # Write updated exports sheet
                exports_df.to_excel(writer, sheet_name="geojson_exports", index=False)
                
            return True
        except Exception as e:
            print(f"Error updating geodata sheet: {e}")
            return False