# export_geojson.py
import argparse
import os
import sys
import json
from utils.geojson_exporter import GeoJSONExporter

def main():
    parser = argparse.ArgumentParser(description="Export In Bounds geodata to GeoJSON file")
    parser.add_argument(
        "--geodata-file", 
        required=True, 
        help="Path to geodata Excel file"
    )
    parser.add_argument(
        "--output-dir", 
        default="data/geojson", 
        help="Directory for GeoJSON output files"
    )
    parser.add_argument(
        "--attributes",
        nargs="*", 
        help="Additional attributes to include in GeoJSON properties"
    )
    parser.add_argument(
        "--format",
        choices=["pretty", "compact"], 
        default="compact",
        help="Output format (pretty for human-readable, compact for smaller files)"
    )
    
    args = parser.parse_args()
    
    if not os.path.exists(args.geodata_file):
        print(f"Error: Geodata file not found: {args.geodata_file}")
        sys.exit(1)
        
    exporter = GeoJSONExporter(args.geodata_file, args.output_dir)
    result = exporter.export_in_bounds_features(args.attributes)
    
    if "error" in result:
        print(f"Error: {result['error']}")
        sys.exit(1)
        
    print(f"Successfully exported {result['feature_count']} features to {result['file_path']}")
    
    # Display the GeoJSON if pretty format is requested
    if args.format == "pretty" and result['feature_count'] > 0:
        with open(result['file_path'], 'r') as f:
            geojson_data = json.load(f)
            print(json.dumps(geojson_data, indent=2))

if __name__ == "__main__":
    main()