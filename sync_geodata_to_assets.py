# sync_geodata_to_assets.py

import argparse
import sys
import logging
from utils.geodata_sync import GeodataAssetSynchronizer

def main():
    parser = argparse.ArgumentParser(
        description='Synchronize coordinates from geodata sheet to asset registry'
    )
    parser.add_argument('--geodata-file', required=True, 
                       help='Path to geodata Excel file')
    parser.add_argument('--asset-file', required=True,
                       help='Path to asset registry Excel file')
    parser.add_argument('--force', action='store_true',
                       help='Force update of existing coordinates in asset registry')
    parser.add_argument('--no-backup', action='store_true',
                       help='Skip creating backup of asset registry before modification')
    parser.add_argument('--no-metadata', action='store_true',
                       help='Skip including metadata columns (source, timestamp, etc.)')
    parser.add_argument('--report-file',
                       help='Path to save synchronization report')
    parser.add_argument('--verbose', action='store_true',
                       help='Enable verbose logging')
    
    args = parser.parse_args()
    
    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    logger = logging.getLogger("GeoSync")
    
    try:
        # Create synchronizer
        sync = GeodataAssetSynchronizer(
            geodata_file=args.geodata_file,
            asset_registry_file=args.asset_file,
            metadata_enabled=not args.no_metadata
        )
        
        # Run synchronization
        stats = sync.synchronize(
            force_update=args.force,
            backup=not args.no_backup
        )
        
        # Generate and display report
        report = sync.generate_report()
        print(report)
        
        # Save report if requested
        if args.report_file:
            with open(args.report_file, 'w') as f:
                f.write(report)
            logger.info(f"Report saved to {args.report_file}")
        
        return 0
        
    except Exception as e:
        logger.error(f"Error during geodata synchronization: {str(e)}")
        if args.verbose:
            logger.exception("Detailed error:")
        return 1

if __name__ == "__main__":
    sys.exit(main())