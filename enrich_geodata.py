# enrich_geodata.py
#!/usr/bin/env python3
"""
CityInfraXLS Geodata Enrichment Orchestrator

This script orchestrates the complete geodata enrichment workflow:
1. Validates the geodata sheet structure and data
2. Runs batch geocoding with exponential backoff for missing coordinates
3. Performs boundary validation to mark assets as in/out of bounds
4. Exports validated geodata to GeoJSON for mapping applications
5. Propagates coordinates back to the main asset registry
6. Generates a comprehensive summary report in the geodata file

The process is designed to be resilient with proper error handling
and provides detailed logging at each step.
"""

import argparse
import sys
import logging
import time
import json
import os
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple

# Import utility modules from our toolchain
from utils.geodata_handler import GeodataHandler
from utils.boundary_validator import BoundaryValidator
from utils.geojson_exporter import GeoJSONExporter
from utils.geodata_sync import GeodataAssetSynchronizer

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(module)s:%(funcName)s] - %(message)s'
)
logger = logging.getLogger("GeodataEnrichOrchestrator")


class EnrichmentOrchestrator:
    """
    Orchestrates the complete geodata enrichment workflow from validation through
    geocoding, boundary validation, GeoJSON export, and asset registry synchronization.
    """
    
    def __init__(self, 
                geodata_file: str,
                asset_registry_file: str,
                boundary_file: Optional[str] = None,
                geocoder_type: str = "google",
                config_file: str = "geocoding_config.json",
                output_dir: str = "./exports",
                workers: Optional[int] = None,
                max_retries: int = 3):
        """
        Initialize the enrichment orchestrator.
        
        Args:
            geodata_file: Path to the geodata Excel file
            asset_registry_file: Path to the asset registry Excel file
            boundary_file: Optional GeoJSON file with boundary polygons
            geocoder_type: Geocoding provider to use
            config_file: Path to geocoding configuration
            output_dir: Directory for exports
            workers: Number of parallel workers for geocoding
            max_retries: Maximum retry attempts for geocoding
        """
        self.geodata_file = Path(geodata_file)
        self.asset_registry_file = Path(asset_registry_file)
        self.boundary_file = Path(boundary_file) if boundary_file else None
        self.geocoder_type = geocoder_type
        self.config_file = config_file
        self.output_dir = Path(output_dir)
        self.workers = workers
        self.max_retries = max_retries
        
        # Create output directory if it doesn't exist
        self.output_dir.mkdir(exist_ok=True, parents=True)
        
        # Stats collection for each step
        self.stats = {
            "validation": {},
            "geocoding": {},
            "boundary_check": {},
            "geojson_export": {},
            "asset_sync": {},
            "overall": {
                "start_time": None,
                "end_time": None,
                "duration_seconds": None,
                "total_assets_processed": 0,
                "successfully_enriched": 0,
                "errors": []
            }
        }
    
    def _ensure_summary_sheet(self) -> None:
        """Create EnrichmentSummary sheet in the geodata file if it doesn't exist."""
        try:
            with pd.ExcelFile(self.geodata_file) as xls:
                if "EnrichmentSummary" not in xls.sheet_names:
                    # Create new summary sheet
                    summary_df = pd.DataFrame(columns=[
                        "timestamp", "operation_type", "status", "details", "assets_affected"
                    ])
                    
                    with pd.ExcelWriter(self.geodata_file, engine='openpyxl', mode='a') as writer:
                        summary_df.to_excel(writer, sheet_name="EnrichmentSummary", index=False)
        except Exception as e:
            logger.warning(f"Could not ensure summary sheet exists: {str(e)}")
    
    def _add_to_summary(self, 
                       operation_type: str, 
                       status: str, 
                       details: str,
                       assets_affected: int) -> None:
        """
        Add an entry to the EnrichmentSummary sheet.
        
        Args:
            operation_type: Type of operation (validation, geocoding, etc.)
            status: Status of operation (success, error, warning)
            details: Detailed description of operation results
            assets_affected: Number of assets affected by the operation
        """
        try:
            # Read existing summary sheet
            try:
                summary_df = pd.read_excel(self.geodata_file, sheet_name="EnrichmentSummary")
            except:
                # Sheet doesn't exist, create it
                self._ensure_summary_sheet()
                summary_df = pd.DataFrame(columns=[
                    "timestamp", "operation_type", "status", "details", "assets_affected"
                ])
            
            # Add new row
            new_row = {
                "timestamp": datetime.now().isoformat(),
                "operation_type": operation_type,
                "status": status,
                "details": details,
                "assets_affected": assets_affected
            }
            
            summary_df = pd.concat([summary_df, pd.DataFrame([new_row])], ignore_index=True)
            
            # Write back to Excel
            with pd.ExcelWriter(self.geodata_file, engine='openpyxl', mode='a',
                               if_sheet_exists='replace') as writer:
                summary_df.to_excel(writer, sheet_name="EnrichmentSummary", index=False)
                
        except Exception as e:
            logger.error(f"Failed to update summary sheet: {str(e)}")
    
    def validate_geodata_structure(self) -> bool:
        """
        Validate the structure and contents of the geodata file.
        
        Returns:
            True if validation successful, False otherwise
        """
        logger.info("Step 1: Validating geodata structure")
        
        try:
            validator = GeodataHandler(self.geodata_file)
            validation_result = validator.validate_geodata()
            
            if validation_result["valid"]:
                logger.info("✓ Geodata structure validation successful")
                self.stats["validation"] = {
                    "status": "success",
                    "message": "Geodata structure is valid",
                    "assets_checked": validation_result.get("records_checked", 0)
                }
                
                self._add_to_summary(
                    operation_type="validation",
                    status="success",
                    details="Geodata structure is valid and conforms to schema",
                    assets_affected=validation_result.get("records_checked", 0)
                )
                
                return True
            else:
                errors = validation_result.get("errors", [])
                error_msg = "; ".join(errors)
                logger.error(f"✗ Geodata validation failed: {error_msg}")
                
                self.stats["validation"] = {
                    "status": "error",
                    "message": "Validation failed",
                    "errors": errors,
                    "assets_checked": validation_result.get("records_checked", 0)
                }
                
                self._add_to_summary(
                    operation_type="validation",
                    status="error",
                    details=f"Validation errors: {error_msg}",
                    assets_affected=validation_result.get("records_checked", 0)
                )
                
                return False
                
        except Exception as e:
            logger.error(f"Error during geodata validation: {str(e)}")
            
            self.stats["validation"] = {
                "status": "error",
                "message": f"Exception during validation: {str(e)}",
                "errors": [str(e)]
            }
            
            self._add_to_summary(
                operation_type="validation",
                status="error",
                details=f"Exception during validation: {str(e)}",
                assets_affected=0
            )
            
            return False
    
    def perform_geocoding(self) -> bool:
        """
        Run batch geocoding with retry logic for missing coordinates.
        
        Returns:
            True if geocoding completed (even with some failures), False if critical error
        """
        logger.info("Step 2: Batch geocoding missing coordinates")
        
        try:
            geocoder = RobustParallelGeocoder(
                geodata_file=str(self.geodata_file),
                geocoder_type=self.geocoder_type,
                max_workers=self.workers,
                max_retries=self.max_retries,
                config_file=self.config_file
            )
            
            # Run geocoding
            geocoding_stats = geocoder.run_geocoding()
            
            # Update stats
            self.stats["geocoding"] = geocoding_stats
            
            # Generate summary
            summary = geocoder.generate_summary()
            
            # Determine status
            if geocoding_stats.get("successful_geocodes", 0) > 0:
                status = "success"
                if geocoding_stats.get("failed_geocodes", 0) > 0:
                    status = "partial"
                    
                details = (f"Geocoded {geocoding_stats.get('successful_geocodes', 0)} assets successfully, "
                          f"{geocoding_stats.get('failed_geocodes', 0)} failed")
                
                self._add_to_summary(
                    operation_type="geocoding",
                    status=status,
                    details=details,
                    assets_affected=geocoding_stats.get("total_records", 0)
                )
                
                logger.info(f"✓ Batch geocoding completed: {details}")
                return True
                
            elif "error" in geocoding_stats:
                # Critical error occurred
                logger.error(f"✗ Geocoding failed: {geocoding_stats['error']}")
                
                self._add_to_summary(
                    operation_type="geocoding",
                    status="error",
                    details=f"Geocoding error: {geocoding_stats['error']}",
                    assets_affected=0
                )
                
                return False
                
            else:
                # No records needed geocoding
                logger.info("✓ No records required geocoding")
                
                self._add_to_summary(
                    operation_type="geocoding",
                    status="success",
                    details="No records required geocoding",
                    assets_affected=0
                )
                
                return True
                
        except Exception as e:
            logger.error(f"Error during batch geocoding: {str(e)}")
            
            self.stats["geocoding"] = {
                "status": "error",
                "message": f"Exception during geocoding: {str(e)}",
                "error": str(e)
            }
            
            self._add_to_summary(
                operation_type="geocoding",
                status="error",
                details=f"Exception during geocoding: {str(e)}",
                assets_affected=0
            )
            
            return False
    
    def validate_boundaries(self) -> bool:
        """
        Validate whether geocoded points fall within defined boundaries.
        
        Returns:
            True if boundary validation completed, False if critical error
        """
        logger.info("Step 3: Validating coordinates against boundary")
        
        # Skip if no boundary file provided
        if not self.boundary_file or not self.boundary_file.exists():
            logger.info("Skipping boundary validation (no boundary file provided)")
            
            self.stats["boundary_check"] = {
                "status": "skipped",
                "message": "No boundary file provided"
            }
            
            self._add_to_summary(
                operation_type="boundary_validation",
                status="skipped",
                details="No boundary file provided for validation",
                assets_affected=0
            )
            
            return True
            
        try:
            validator = BoundaryValidator(
                geodata_file=str(self.geodata_file),
                boundary_file=str(self.boundary_file)
            )
            
            # Run validation
            validation_results = validator.validate_coordinates()
            
            # Update stats
            self.stats["boundary_check"] = validation_results
            
            in_bounds = validation_results.get("in_bounds", 0)
            out_bounds = validation_results.get("out_of_bounds", 0)
            total = in_bounds + out_bounds
            
            details = f"Found {in_bounds} assets within boundary, {out_bounds} outside boundary"
            logger.info(f"✓ Boundary validation completed: {details}")
            
            self._add_to_summary(
                operation_type="boundary_validation",
                status="success",
                details=details,
                assets_affected=total
            )
            
            return True
            
        except Exception as e:
            logger.error(f"Error during boundary validation: {str(e)}")
            
            self.stats["boundary_check"] = {
                "status": "error",
                "message": f"Exception during boundary validation: {str(e)}",
                "error": str(e)
            }
            
            self._add_to_summary(
                operation_type="boundary_validation",
                status="error",
                details=f"Exception during boundary validation: {str(e)}",
                assets_affected=0
            )
            
            return False
    
    def export_geojson(self) -> bool:
        """
        Export geodata to GeoJSON for mapping.
        
        Returns:
            True if export completed successfully, False otherwise
        """
        logger.info("Step 4: Exporting geodata to GeoJSON")
        
        try:
            exporter = GeoJSONExporter(
                geodata_file=str(self.geodata_file),
                output_dir=str(self.output_dir / "geojson")
            )
            
            # Run export
            export_result = exporter.export_and_record()
            
            # Update stats
            self.stats["geojson_export"] = export_result
            
            details = (f"Exported {export_result.get('feature_count', 0)} features to "
                     f"{export_result.get('layer_name', 'geojson')}") 
            
            logger.info(f"✓ GeoJSON export completed: {details}")
            
            self._add_to_summary(
                operation_type="geojson_export",
                status="success",
                details=details,
                assets_affected=export_result.get("feature_count", 0)
            )
            
            return True
            
        except Exception as e:
            logger.error(f"Error during GeoJSON export: {str(e)}")
            
            self.stats["geojson_export"] = {
                "status": "error",
                "message": f"Exception during GeoJSON export: {str(e)}",
                "error": str(e)
            }
            
            self._add_to_summary(
                operation_type="geojson_export",
                status="error",
                details=f"Exception during GeoJSON export: {str(e)}",
                assets_affected=0
            )
            
            return False
    
    def sync_to_asset_registry(self) -> bool:
        """
        Synchronize coordinates from geodata to asset registry.
        
        Returns:
            True if sync completed successfully, False otherwise
        """
        logger.info("Step 5: Synchronizing coordinates to asset registry")
        
        try:
            synchronizer = GeodataAssetSynchronizer(
                geodata_file=str(self.geodata_file),
                asset_registry_file=str(self.asset_registry_file)
            )
            
            # Run synchronization
            sync_stats = synchronizer.synchronize(force_update=False, backup=True)
            
            # Update stats
            self.stats["asset_sync"] = sync_stats
            
            details = (f"Updated {sync_stats.get('assets_updated', 0)} assets, "
                     f"skipped {sync_stats.get('assets_skipped', 0)}, "
                     f"not found {sync_stats.get('assets_not_found', 0)}")
            
            logger.info(f"✓ Asset synchronization completed: {details}")
            
            self._add_to_summary(
                operation_type="asset_sync",
                status="success",
                details=details,
                assets_affected=sync_stats.get("assets_updated", 0)
            )
            
            return True
            
        except Exception as e:
            logger.error(f"Error during asset synchronization: {str(e)}")
            
            self.stats["asset_sync"] = {
                "status": "error",
                "message": f"Exception during asset synchronization: {str(e)}",
                "error": str(e)
            }
            
            self._add_to_summary(
                operation_type="asset_sync",
                status="error",
                details=f"Exception during asset synchronization: {str(e)}",
                assets_affected=0
            )
            
            return False
    
    def write_final_summary(self) -> None:
        """Generate and write final summary of entire enrichment process."""
        logger.info("Generating final enrichment summary")
        
        try:
            # Generate human-readable summary
            validation_status = self.stats["validation"].get("status", "not run")
            geocoding_status = self.stats["geocoding"].get("status", "not run")
            boundary_status = self.stats["boundary_check"].get("status", "not run")
            export_status = self.stats["geojson_export"].get("status", "not run")
            sync_status = self.stats["asset_sync"].get("status", "not run")
            
            # Determine overall status
            if all(s == "success" for s in [validation_status, geocoding_status, boundary_status, 
                                           export_status, sync_status]):
                overall_status = "success"
                status_msg = "All operations completed successfully"
            elif any(s == "error" for s in [validation_status, geocoding_status, boundary_status, 
                                          export_status, sync_status]):
                overall_status = "error"
                status_msg = "One or more operations failed"
            else:
                overall_status = "partial"
                status_msg = "Some operations completed with warnings"
            
            # Collect metrics
            assets_processed = self.stats["validation"].get("assets_checked", 0)
            geocoded_assets = self.stats["geocoding"].get("successful_geocodes", 0)
            assets_updated = self.stats["asset_sync"].get("assets_updated", 0)
            
            # Create summary message
            summary_message = (
                f"Enrichment completed with status: {overall_status}. {status_msg}.\n"
                f"Processed {assets_processed} assets, geocoded {geocoded_assets}, "
                f"synchronized {assets_updated} to asset registry."
            )
            
            # Add to summary sheet
            self._add_to_summary(
                operation_type="enrichment_complete",
                status=overall_status,
                details=summary_message,
                assets_affected=assets_processed
            )
            
            # Calculate total runtime
            total_runtime = 0
            if self.stats["overall"]["start_time"] and self.stats["overall"]["end_time"]:
                total_runtime = (self.stats["overall"]["end_time"] - 
                                self.stats["overall"]["start_time"]).total_seconds()
                
            # Save detailed report to JSON
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_file = self.output_dir / f"enrichment_report_{timestamp}.json"
            
            # Prepare JSON-serializable stats
            json_stats = self._prepare_stats_for_json(self.stats)
            
            with open(report_file, 'w') as f:
                json.dump(json_stats, f, indent=2)
                
            logger.info(f"Detailed enrichment report saved to {report_file}")
            
        except Exception as e:
            logger.error(f"Error generating final summary: {str(e)}")
    
    def _prepare_stats_for_json(self, stats_dict: Dict) -> Dict:
        """Convert stats dictionary to JSON-serializable format."""
        import copy
        
        json_dict = copy.deepcopy(stats_dict)
        
        # Convert datetime objects to strings
        for key, value in json_dict.items():
            if isinstance(value, dict):
                json_dict[key] = self._prepare_stats_for_json(value)
            elif isinstance(value, datetime):
                json_dict[key] = value.isoformat()
                
        return json_dict
    
    def run_enrichment(self) -> bool:
        """
        Run the complete enrichment workflow.
        
        Returns:
            True if enrichment completed successfully, False if critical errors occurred
        """
        # Record start time
        self.stats["overall"]["start_time"] = datetime.now()
        
        logger.info(f"Starting geodata enrichment workflow for {self.geodata_file}")
        
        # Step 1: Validate geodata file structure
        if not self.validate_geodata_structure():
            logger.error("Geodata validation failed, aborting workflow")
            
            self.stats["overall"]["end_time"] = datetime.now()
            self.stats["overall"]["duration_seconds"] = (
                self.stats["overall"]["end_time"] - self.stats["overall"]["start_time"]
            ).total_seconds()
            
            self.write_final_summary()
            return False
            
        # Step 2: Run batch geocoding for missing coordinates
        geocoding_result = self.perform_geocoding()
        if not geocoding_result:
            logger.error("Geocoding failed with critical error, proceeding with other steps")
            
        # Step 3: Validate coordinates against boundary
        boundary_result = self.validate_boundaries()
        if not boundary_result:
            logger.error("Boundary validation failed, proceeding with other steps")
            
        # Step 4: Export to GeoJSON for mapping
        export_result = self.export_geojson()
        if not export_result:
            logger.error("GeoJSON export failed, proceeding with other steps")
            
        # Step 5: Sync coordinates to asset registry
        sync_result = self.sync_to_asset_registry()
        if not sync_result:
            logger.error("Asset registry synchronization failed")
            
        # Record end time
        self.stats["overall"]["end_time"] = datetime.now()
        self.stats["overall"]["duration_seconds"] = (
            self.stats["overall"]["end_time"] - self.stats["overall"]["start_time"]
        ).total_seconds()
        
        # Generate and write final summary
        self.write_final_summary()
        
        # Determine overall success based on critical operations
        critical_ops = [geocoding_result, sync_result]
        return all(critical_ops)


def main():
    parser = argparse.ArgumentParser(
        description="CityInfraXLS Geodata Enrichment Orchestrator"
    )
    
    parser.add_argument("--geodata-file", required=True,
                       help="Path to geodata Excel file")
    parser.add_argument("--asset-file", required=True,
                       help="Path to asset registry Excel file")
    parser.add_argument("--boundary-file",
                       help="Path to GeoJSON boundary file for validation")
    parser.add_argument("--geocoder", choices=["google", "osm"], default="google",
                       help="Geocoding provider to use")
    parser.add_argument("--config", default="geocoding_config.json",
                       help="Path to geocoding configuration file")
    parser.add_argument("--output-dir", default="./exports",
                       help="Directory for exports")
    parser.add_argument("--workers", type=int,
                       help="Number of parallel workers for geocoding")
    parser.add_argument("--max-retries", type=int, default=3,
                       help="Maximum retry attempts for geocoding API errors")
    parser.add_argument("--verbose", action="store_true",
                       help="Enable verbose logging")
    
    args = parser.parse_args()
    
    # Set log level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
        
    try:
        # Create and run orchestrator
        orchestrator = EnrichmentOrchestrator(
            geodata_file=args.geodata_file,
            asset_registry_file=args.asset_file,
            boundary_file=args.boundary_file,
            geocoder_type=args.geocoder,
            config_file=args.config,
            output_dir=args.output_dir,
            workers=args.workers,
            max_retries=args.max_retries
        )
        
        success = orchestrator.run_enrichment()
        
        if success:
            logger.info("✓ Enrichment workflow completed successfully")
            return 0
        else:
            logger.warning("⚠ Enrichment workflow completed with issues")
            return 1
            
    except Exception as e:
        logger.error(f"Fatal error during enrichment workflow: {str(e)}")
        if args.verbose:
            import traceback
            logger.debug(traceback.format_exc())
        return 2


if __name__ == "__main__":
    sys.exit(main())