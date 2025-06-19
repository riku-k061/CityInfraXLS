# enrich_asset_geodata.py
import argparse
import pandas as pd
from utils.geodata_enrichment import GeodataEnrichmentManager

def main():
    parser = argparse.ArgumentParser(description="Enrich assets with geodata")
    parser.add_argument('--asset_id', help='Asset ID to enrich')
    parser.add_argument('--latitude', type=float, help='Latitude coordinate')
    parser.add_argument('--longitude', type=float, help='Longitude coordinate')
    parser.add_argument('--source', default='MANUAL_ENTRY', 
                        choices=['GPS', 'ADDRESS_LOOKUP', 'MANUAL_ENTRY', 'IMPORTED', 'SURVEY_DATA'],
                        help='Source of the geodata')
    parser.add_argument('--file', help='CSV file with multiple assets to enrich')
    parser.add_argument('--force', action='store_true', help='Force update existing coordinates')
    parser.add_argument('--check', action='store_true', help='Check geodata status without enriching')
    args = parser.parse_args()
    
    # Create the enrichment manager
    manager = GeodataEnrichmentManager()
    
    # Prepare the system for enrichment
    ready, message = manager.prepare_for_enrichment()
    if not ready:
        print(f"Error: {message}")
        return 1
    
    print(message)
    
    if args.check:
        # Just check status
        if args.asset_id:
            # Check single asset
            has_coords, data = manager.check_existing_coordinates(args.asset_id)
            if data:
                print(f"Asset {args.asset_id} geodata status:")
                for key, value in data.items():
                    print(f"  {key}: {value}")
            else:
                print(f"No geodata found for asset {args.asset_id}")
        else:
            # Check all assets
            status_df = manager.bulk_check_geodata_status()
            if status_df.empty:
                print("No geodata records found")
            else:
                print("\nGeodata Status Summary:")
                print(status_df)
                
                # Show statistics
                total = len(status_df)
                with_coords = status_df['has_coordinates'].sum()
                percent = (with_coords / total) * 100 if total > 0 else 0
                
                print(f"\nTotal assets with geodata: {total}")
                print(f"Assets with coordinates: {with_coords} ({percent:.1f}%)")
                print(f"Assets missing coordinates: {total - with_coords}")
                
                # Show validation status breakdown
                print("\nValidation Status Breakdown:")
                status_counts = status_df['validation_status'].value_counts()
                for status, count in status_counts.items():
                    print(f"  {status}: {count}")
        
        return 0
    
    if args.file:
        # Process bulk file
        try:
            df = pd.read_csv(args.file)
            required_cols = ['asset_id', 'latitude', 'longitude']
            missing_cols = [col for col in required_cols if col not in df.columns]
            
            if missing_cols:
                print(f"Error: CSV file missing required columns: {', '.join(missing_cols)}")
                return 1
                
            print(f"Processing {len(df)} assets for geodata enrichment...")
            
            success_count = 0
            for _, row in df.iterrows():
                additional_data = {}
                for col in df.columns:
                    if col not in ['asset_id', 'latitude', 'longitude', 'geocode_source'] and not pd.isna(row[col]):
                        additional_data[col] = row[col]
                
                source = row.get('geocode_source', args.source)
                
                if args.force:
                    success, msg = manager.update_asset_geodata(
                        row['asset_id'], 
                        {
                            'latitude': row['latitude'],
                            'longitude': row['longitude'],
                            'geocode_source': source,
                            **additional_data
                        },
                        force_update=True
                    )
                else:
                    success, msg = manager.enrich_asset_geodata(
                        row['asset_id'],
                        row['latitude'],
                        row['longitude'],
                        source=source,
                        additional_data=additional_data
                    )
                
                if success:
                    success_count += 1
                else:
                    print(f"Warning for {row['asset_id']}: {msg}")
            
            print(f"Successfully enriched {success_count} out of {len(df)} assets")
            
        except Exception as e:
            print(f"Error processing file: {e}")
            return 1
    
    elif args.asset_id and args.latitude is not None and args.longitude is not None:
        # Process single asset
        if args.force:
            success, msg = manager.update_asset_geodata(
                args.asset_id,
                {
                    'latitude': args.latitude,
                    'longitude': args.longitude,
                    'geocode_source': args.source
                },
                force_update=True
            )
        else:
            success, msg = manager.enrich_asset_geodata(
                args.asset_id,
                args.latitude,
                args.longitude,
                source=args.source
            )
            
        if success:
            print(f"Successfully enriched asset {args.asset_id}")
        else:
            print(f"Error: {msg}")
            return 1
    
    else:
        print("Error: Must provide either --file or --asset_id with --latitude and --longitude")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())