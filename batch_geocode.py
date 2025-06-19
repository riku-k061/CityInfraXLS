# batch_geocode.py
import argparse
import logging
from utils.batch_geocoder import BatchGeocoder

def main():
    parser = argparse.ArgumentParser(description="Batch geocode assets with missing coordinates")
    
    # API and provider options
    parser.add_argument('--provider', default='nominatim', 
                        choices=['nominatim', 'google', 'mapbox', 'here', 'arcgis', 'test'],
                        help='Geocoding service provider')
    parser.add_argument('--api-key', help='API key for geocoding service')
    
    # Batch processing options
    parser.add_argument('--batch-size', type=int, default=10, 
                        help='Number of locations to geocode in each batch')
    parser.add_argument('--rate-limit', type=float, default=1.5, 
                        help='Seconds between API requests (respect rate limits)')
    parser.add_argument('--max-assets', type=int, 
                        help='Maximum number of assets to process')
    parser.add_argument('--workers', type=int, default=4, 
                        help='Number of worker threads for batch processing')
    
    # Debug options
    parser.add_argument('--debug', action='store_true', 
                        help='Enable debug logging')
    
    args = parser.parse_args()
    
    # Configure logging
    log_level = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("geocoding.log"),
            logging.StreamHandler()
        ]
    )
    
    logger = logging.getLogger("BatchGeocode")
    
    try:
        # Initialize geocoder
        geocoder = BatchGeocoder(
            api_key=args.api_key,
            provider=args.provider,
            rate_limit=args.rate_limit,
            batch_size=args.batch_size
        )
        
        # Run batch geocoding
        success_count, total_count = geocoder.run_geocoding_batch(
            max_assets=args.max_assets,
            workers=args.workers
        )
        
        # Report results
        if total_count == 0:
            print("No assets found needing geocoding.")
        else:
            print(f"Geocoding complete: {success_count} of {total_count} assets successfully geocoded.")
            
            if success_count < total_count:
                print(f"Failed to geocode {total_count - success_count} assets. Check the log for details.")
        
        return 0 if success_count == total_count else 1
        
    except Exception as e:
        logger.error(f"Error in batch geocoding: {e}", exc_info=True)
        print(f"Error: {e}")
        return 1

if __name__ == "__main__":
    exit(main())