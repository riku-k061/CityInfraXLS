# manage_assets.py - Main entry point for CityInfraXLS

import argparse
import sys
from datetime import datetime

# Import functionality from the modules
from register_asset import register_asset
from query_assets import query_assets
from delete_asset import delete_asset

def setup_parser():
    """Set up the argument parser with subcommands"""
    parser = argparse.ArgumentParser(
        description='CityInfraXLS - Urban Infrastructure Maintenance & Analytics System',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python manage_assets.py register               Register a new asset with interactive prompts
  python manage_assets.py query --type Bridge    Query all bridge assets
  python manage_assets.py delete ABC123          Delete asset with ID ABC123
"""
    )
    
    # Create subparsers for each command
    subparsers = parser.add_subparsers(dest='command', help='Command to execute')
    
    # Register command
    register_parser = subparsers.add_parser('register', 
        help='Register a new asset (interactive mode)')
    
    # Query command
    query_parser = subparsers.add_parser('query', 
        help='Query existing assets')
    query_parser.add_argument('--type', 
        help='Filter by asset type (e.g., Road, Bridge, Park)')
    query_parser.add_argument('--location', 
        help='Filter by location (case insensitive, partial match)')
    query_parser.add_argument('--installed-after', 
        help='Filter by installation date (YYYY-MM-DD)')
    query_parser.add_argument('--export', 
        help='Export results to Excel file at specified path')
    
    # Delete command
    delete_parser = subparsers.add_parser('delete', 
        help='Delete an existing asset')
    delete_parser.add_argument('asset_id', 
        help='ID of the asset to delete')
    delete_parser.add_argument('--force', '-f', 
        action='store_true', 
        help='Delete without confirmation')
    
    return parser

def main():
    """Main entry point for the application"""
    parser = setup_parser()
    args = parser.parse_args()
    
    if args.command is None:
        parser.print_help()
        return 1
    
    print(f"=== CityInfraXLS - {args.command.capitalize()} Mode ===")
    
    try:
        if args.command == 'register':
            register_asset()
        
        elif args.command == 'query':
            # Validate that at least one filter is specified
            if not (args.type or args.location or args.installed_after):
                print("Error: Please specify at least one filter: --type, --location, or --installed-after")
                query_parser = [p for p in parser._subparsers._actions 
                                if isinstance(p, argparse._SubParsersAction)][0]
                query_parser.choices['query'].print_help()
                return 1
            
            # Validate date format if provided
            if args.installed_after:
                try:
                    datetime.strptime(args.installed_after, '%Y-%m-%d')
                except ValueError:
                    print("Error: Date format must be YYYY-MM-DD")
                    return 1
                
            query_assets(
                asset_type=args.type,
                location=args.location,
                installed_after=args.installed_after,
                export_path=args.export
            )
        
        elif args.command == 'delete':
            delete_asset(args.asset_id, confirm=not args.force)
    
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130  # Standard exit code for SIGINT
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())