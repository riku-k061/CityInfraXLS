# utils/boundary_validator.py
import os
import pandas as pd
import geopandas as gpd
import json
from shapely.geometry import Point, Polygon
from pathlib import Path
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("BoundaryValidator")

class BoundaryValidator:
    def __init__(self, boundary_file=None):
        """
        Initialize the boundary validator
        
        Args:
            boundary_file (str): Path to GIS boundary file (GeoJSON, Shapefile, etc.)
                                If None, will look in standard locations
        """
        self.boundary_file = boundary_file
        self.boundaries = None
        self.boundary_crs = "EPSG:4326"  # WGS84 - standard for GPS coordinates
        
        # Load boundaries
        self._load_boundaries()
    
    def _load_boundaries(self):
        """Load city boundary polygons from GIS resources"""
        if self.boundary_file and os.path.exists(self.boundary_file):
            # Use specified file
            boundary_path = self.boundary_file
        else:
            # Look in standard locations
            possible_paths = [
                Path(__file__).parent.parent / 'data' / 'boundaries' / 'city_boundary.geojson',
                Path(__file__).parent.parent / 'data' / 'boundaries' / 'city_boundary.shp',
                Path(__file__).parent.parent / 'data' / 'city_boundary.geojson',
                Path('city_boundary.geojson')
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    boundary_path = path
                    break
            else:
                # No boundary file found
                logger.error("No boundary file found. Validation will mark all points as out of bounds.")
                self.boundaries = None
                return
        
        try:
            # Load boundary file as GeoDataFrame
            self.boundaries = gpd.read_file(boundary_path)
            
            # Ensure it's using the expected CRS (WGS84)
            if self.boundaries.crs is None:
                logger.warning("Boundary file has no CRS defined, assuming WGS84")
                self.boundaries.crs = self.boundary_crs
            elif self.boundaries.crs != self.boundary_crs:
                logger.info(f"Reprojecting boundaries from {self.boundaries.crs} to {self.boundary_crs}")
                self.boundaries = self.boundaries.to_crs(self.boundary_crs)
            
            logger.info(f"Loaded boundary file with {len(self.boundaries)} polygon(s)")
            
            # If no 'name' column, add one for reference
            if 'name' not in self.boundaries.columns:
                self.boundaries['name'] = [f"Boundary {i+1}" for i in range(len(self.boundaries))]
            
        except Exception as e:
            logger.error(f"Error loading boundary file: {e}")
            self.boundaries = None
    
    def is_point_in_bounds(self, latitude, longitude):
        """
        Check if a point is within any city boundary
        
        Args:
            latitude (float): Latitude coordinate
            longitude (float): Longitude coordinate
            
        Returns:
            tuple: (bool, str) - (in_bounds, boundary_name)
        """
        if self.boundaries is None:
            return False, "No boundary data available"
        
        try:
            # Create point
            point = gpd.GeoDataFrame(
                [{'geometry': Point(longitude, latitude)}],
                crs=self.boundary_crs
            )
            
            # Check if point is within any boundary
            for idx, boundary in self.boundaries.iterrows():
                if point.geometry.iloc[0].within(boundary.geometry):
                    return True, boundary['name']
            
            # If we get here, point is not in any boundary
            return False, "Outside all defined boundaries"
            
        except Exception as e:
            logger.error(f"Error checking point bounds: {e}")
            return False, f"Error: {str(e)}"
    
    def validate_geodata(self, update_sheet=True, create_report=True):
        """
        Validate all coordinates in geodata against boundary polygons
        
        Args:
            update_sheet (bool): Whether to add in_bounds flag to geodata sheet
            create_report (bool): Whether to create data quality report
            
        Returns:
            tuple: (total_checked, in_bounds_count, out_bounds_count)
        """
        try:
            # Load geodata
            geodata_path = Path(__file__).parent.parent / 'data' / 'asset_geodata.xlsx'
            if not os.path.exists(geodata_path):
                logger.error("Geodata file not found")
                return 0, 0, 0
            
            # Load the Excel file
            geodata_df = pd.read_excel(geodata_path, sheet_name='AssetGeodata')
            
            # Skip rows without coordinates
            valid_coords_df = geodata_df.dropna(subset=['latitude', 'longitude'])
            
            if valid_coords_df.empty:
                logger.info("No valid coordinates found in geodata")
                return 0, 0, 0
            
            # Add/initialize in_bounds column
            if 'in_bounds' not in valid_coords_df.columns:
                valid_coords_df['in_bounds'] = None
                
            if 'boundary_name' not in valid_coords_df.columns:
                valid_coords_df['boundary_name'] = None
            
            # Track statistics
            total_checked = len(valid_coords_df)
            in_bounds_count = 0
            out_bounds_count = 0
            
            # Store out-of-bounds assets for report
            out_of_bounds_assets = []
            
            # Process each row
            for idx, row in valid_coords_df.iterrows():
                try:
                    latitude, longitude = row['latitude'], row['longitude']
                    in_bounds, boundary_name = self.is_point_in_bounds(latitude, longitude)
                    
                    # Update dataframe
                    valid_coords_df.loc[idx, 'in_bounds'] = in_bounds
                    valid_coords_df.loc[idx, 'boundary_name'] = boundary_name
                    
                    if in_bounds:
                        in_bounds_count += 1
                    else:
                        out_bounds_count += 1
                        out_of_bounds_assets.append({
                            'asset_id': row['asset_id'],
                            'latitude': latitude,
                            'longitude': longitude,
                            'boundary_name': boundary_name,
                            'geocode_source': row.get('geocode_source', 'UNKNOWN'),
                            'validation_status': row.get('validation_status', 'UNVERIFIED'),
                            'address_string': row.get('address_string', '')
                        })
                        
                except Exception as e:
                    logger.error(f"Error validating asset {row.get('asset_id', 'Unknown')}: {e}")
            
            # Update original dataframe with validated data
            for col in ['in_bounds', 'boundary_name']:
                if col in valid_coords_df.columns:
                    geodata_df.loc[valid_coords_df.index, col] = valid_coords_df[col]
            
            # Save updated geodata if requested
            if update_sheet:
                # Preserve all sheets
                with pd.ExcelWriter(geodata_path) as writer:
                    geodata_df.to_excel(writer, sheet_name='AssetGeodata', index=False)
                    
                    # Copy other sheets
                    for sheet_name in ['Metadata', 'AllowedValues']:
                        try:
                            sheet_df = pd.read_excel(geodata_path, sheet_name=sheet_name)
                            sheet_df.to_excel(writer, sheet_name=sheet_name)
                        except:
                            # Sheet might not exist yet
                            pass
                
                logger.info(f"Updated geodata sheet with boundary validation results")
            
            # Create data quality report if requested
            if create_report and out_of_bounds_assets:
                self._create_data_quality_report(out_of_bounds_assets, geodata_path)
            
            return total_checked, in_bounds_count, out_bounds_count
            
        except Exception as e:
            logger.error(f"Error validating geodata: {e}")
            return 0, 0, 0
    
    def _create_data_quality_report(self, out_of_bounds_assets, geodata_path):
        """Create data quality report for out-of-bounds assets"""
        try:
            # Create report dataframe
            report_df = pd.DataFrame(out_of_bounds_assets)
            
            # Add assets data if available
            assets_path = Path(__file__).parent.parent / 'data' / 'assets.xlsx'
            if os.path.exists(assets_path):
                assets_df = pd.read_excel(assets_path)
                
                # Get relevant columns for joining
                asset_columns = ['asset_id', 'name', 'type', 'location_description']
                available_cols = [col for col in asset_columns if col in assets_df.columns]
                
                if available_cols:
                    report_df = pd.merge(
                        report_df,
                        assets_df[available_cols],
                        on='asset_id',
                        how='left'
                    )
            
            # Add timestamp and notes columns for manual review
            report_df['review_date'] = None
            report_df['reviewed_by'] = None
            report_df['resolution_notes'] = None
            report_df['resolution_status'] = None
            
            # Add validaton timestamp
            report_df['validation_timestamp'] = datetime.now().isoformat()
            
            # Save to Excel - either in same file or separate file
            quality_report_path = Path(__file__).parent.parent / 'data' / 'geodata_quality_report.xlsx'
            
            # Check if report already exists and append
            if os.path.exists(quality_report_path):
                existing_df = pd.read_excel(quality_report_path)
                
                # Remove duplicates (same asset_id) and keep newer
                combined_df = pd.concat([existing_df, report_df])
                combined_df = combined_df.sort_values('validation_timestamp', ascending=False)
                combined_df = combined_df.drop_duplicates(subset=['asset_id'], keep='first')
                
                report_df = combined_df
            
            # Save report
            report_df.to_excel(quality_report_path, index=False)
            logger.info(f"Created data quality report with {len(report_df)} out-of-bounds assets")
            
        except Exception as e:
            logger.error(f"Error creating data quality report: {e}")