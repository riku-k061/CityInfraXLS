# utils/geodata_sync.py
import pandas as pd
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("GeodataSync")

class GeodataAssetSynchronizer:
    """
    Synchronizes geocoded coordinates from the geodata sheet to the main asset registry.
    Ensures that location data is readily available in the primary asset records.
    """
    
    def __init__(self, 
                 geodata_file: str, 
                 asset_registry_file: str,
                 metadata_enabled: bool = True):
        """
        Initialize the synchronizer with file paths.
        
        Args:
            geodata_file: Path to the geodata Excel file
            asset_registry_file: Path to the main asset registry Excel file
            metadata_enabled: Whether to add metadata columns (source, timestamp, etc.)
        """
        self.geodata_file = Path(geodata_file)
        self.asset_registry_file = Path(asset_registry_file)
        self.metadata_enabled = metadata_enabled
        
        # Define required columns
        self.required_geodata_cols = ['asset_id', 'latitude', 'longitude']
        self.optional_metadata_cols = [
            'geocode_source', 'geocode_timestamp', 'validation_status', 
            'accuracy_meters'
        ]
        
        # Statistics for reporting
        self.stats = {
            "total_assets": 0,
            "assets_with_geodata": 0,
            "assets_updated": 0,
            "assets_skipped": 0,
            "assets_not_found": 0
        }
    
    def _validate_files(self) -> Tuple[bool, str]:
        """
        Validate that input files exist and have required structure.
        
        Returns:
            Tuple of (is_valid, error_message)
        """
        # Check file existence
        if not self.geodata_file.exists():
            return False, f"Geodata file not found: {self.geodata_file}"
        
        if not self.asset_registry_file.exists():
            return False, f"Asset registry file not found: {self.asset_registry_file}"
        
        # Check geodata sheet structure
        try:
            geodata_df = pd.read_excel(self.geodata_file, sheet_name="Geodata")
            missing_cols = [col for col in self.required_geodata_cols if col not in geodata_df.columns]
            if missing_cols:
                return False, f"Missing required columns in geodata: {', '.join(missing_cols)}"
        except Exception as e:
            return False, f"Error reading geodata file: {str(e)}"
        
        # Check asset registry structure
        try:
            asset_df = pd.read_excel(self.asset_registry_file)
            if 'asset_id' not in asset_df.columns:
                return False, "Asset registry must have 'asset_id' column"
        except Exception as e:
            return False, f"Error reading asset registry file: {str(e)}"
            
        return True, ""
    
    def _load_data(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Load data from Excel files.
        
        Returns:
            Tuple of (geodata_df, asset_df)
        """
        geodata_df = pd.read_excel(self.geodata_file, sheet_name="Geodata")
        asset_df = pd.read_excel(self.asset_registry_file)
        
        return geodata_df, asset_df
        
    def _get_columns_to_sync(self, geodata_df: pd.DataFrame) -> List[str]:
        """
        Determine which columns should be synchronized to the asset registry.
        
        Args:
            geodata_df: DataFrame containing geodata
            
        Returns:
            List of column names to sync
        """
        # Always include required columns
        columns_to_sync = self.required_geodata_cols.copy()
        
        # Add metadata columns if enabled and available
        if self.metadata_enabled:
            for col in self.optional_metadata_cols:
                if col in geodata_df.columns:
                    columns_to_sync.append(col)
                    
        return columns_to_sync
    
    def _prepare_asset_registry(self, 
                              asset_df: pd.DataFrame, 
                              columns_to_sync: List[str]) -> pd.DataFrame:
        """
        Prepare the asset registry by adding columns if they don't exist.
        
        Args:
            asset_df: DataFrame containing asset registry
            columns_to_sync: List of columns to be synchronized
            
        Returns:
            Updated asset registry DataFrame
        """
        # Create a copy to avoid modifying original
        updated_df = asset_df.copy()
        
        # Add missing columns
        for col in columns_to_sync:
            if col not in updated_df.columns and col != 'asset_id':
                if col in ['latitude', 'longitude', 'accuracy_meters']:
                    # Numeric columns
                    updated_df[col] = None
                else:
                    # String columns
                    updated_df[col] = ""
        
        return updated_df
    
    def _add_sync_metadata(self, asset_df: pd.DataFrame) -> pd.DataFrame:
        """
        Add or update metadata about the synchronization process.
        
        Args:
            asset_df: Asset registry DataFrame
            
        Returns:
            Updated DataFrame with sync metadata
        """
        # Create a copy to avoid modifying original
        updated_df = asset_df.copy()
        
        # Add synchronization metadata columns
        if 'geodata_sync_timestamp' not in updated_df.columns:
            updated_df['geodata_sync_timestamp'] = ""
        
        # Note: We'll update these values only for records that actually get updated
        
        return updated_df
    
    def synchronize(self, 
                   force_update: bool = False, 
                   backup: bool = True) -> Dict[str, int]:
        """
        Synchronize coordinates from geodata sheet to asset registry.
        
        Args:
            force_update: If True, overwrite existing coordinates in asset registry
            backup: If True, create backup of asset registry before modification
            
        Returns:
            Statistics about the synchronization process
        """
        # Validate files
        is_valid, error_msg = self._validate_files()
        if not is_valid:
            logger.error(f"Validation failed: {error_msg}")
            raise ValueError(error_msg)
        
        # Load data
        geodata_df, asset_df = self._load_data()
        self.stats["total_assets"] = len(asset_df)
        
        # Count assets with geodata
        self.stats["assets_with_geodata"] = len(geodata_df[
            geodata_df['latitude'].notna() & geodata_df['longitude'].notna()
        ])
        
        # Get columns to synchronize
        columns_to_sync = self._get_columns_to_sync(geodata_df)
        logger.info(f"Will synchronize columns: {columns_to_sync}")
        
        # Prepare asset registry with required columns
        updated_asset_df = self._prepare_asset_registry(asset_df, columns_to_sync)
        
        # Add sync metadata if needed
        updated_asset_df = self._add_sync_metadata(updated_asset_df)
        
        # Create mapping of asset IDs to indices for efficient lookups
        asset_indices = {
            asset_id: idx for idx, asset_id in enumerate(updated_asset_df['asset_id'])
        }
        
        # Create backup if requested
        if backup:
            backup_path = self.asset_registry_file.with_suffix(f".backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            asset_df.to_excel(backup_path, index=False)
            logger.info(f"Created backup at {backup_path}")
        
        # Track which assets were updated
        updated_assets = set()
        
        # Process geodata and update asset registry
        for _, row in geodata_df.iterrows():
            asset_id = row['asset_id']
            
            # Skip if missing coordinates
            if pd.isna(row['latitude']) or pd.isna(row['longitude']):
                continue
                
            # Find asset in registry
            if asset_id in asset_indices:
                idx = asset_indices[asset_id]
                
                # Check if should update
                should_update = force_update or pd.isna(updated_asset_df.loc[idx, 'latitude'])
                
                if should_update:
                    # Update all columns in the sync list
                    for col in columns_to_sync:
                        if col != 'asset_id':  # Don't modify the asset_id
                            updated_asset_df.loc[idx, col] = row[col]
                    
                    # Add sync metadata
                    updated_asset_df.loc[idx, 'geodata_sync_timestamp'] = datetime.now().isoformat()
                    
                    # Track updated assets
                    updated_assets.add(asset_id)
                else:
                    self.stats["assets_skipped"] += 1
            else:
                self.stats["assets_not_found"] += 1
                logger.warning(f"Asset {asset_id} from geodata not found in asset registry")
        
        self.stats["assets_updated"] = len(updated_assets)
        
        # Write updated asset registry back to Excel
        if self.stats["assets_updated"] > 0:
            updated_asset_df.to_excel(self.asset_registry_file, index=False)
            logger.info(f"Updated {self.stats['assets_updated']} assets in registry")
        else:
            logger.info("No assets needed updating")
        
        return self.stats
    
    def generate_report(self) -> str:
        """
        Generate a human-readable report of the synchronization process.
        
        Returns:
            Formatted report string
        """
        report = [
            "Geodata Synchronization Report",
            "=============================",
            f"Timestamp: {datetime.now().isoformat()}",
            f"Geodata file: {self.geodata_file}",
            f"Asset registry file: {self.asset_registry_file}",
            "",
            "Statistics:",
            f"  - Total assets in registry: {self.stats['total_assets']}",
            f"  - Assets with geodata: {self.stats['assets_with_geodata']}",
            f"  - Assets updated: {self.stats['assets_updated']}",
            f"  - Assets skipped (already had coordinates): {self.stats['assets_skipped']}",
            f"  - Assets not found in registry: {self.stats['assets_not_found']}",
            "",
            "Summary:",
            f"  Successfully synchronized {self.stats['assets_updated']} assets with geodata."
        ]
        
        return "\n".join(report)