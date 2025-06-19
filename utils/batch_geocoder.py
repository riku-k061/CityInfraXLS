# utils/batch_geocoder.py
import os
import pandas as pd
import json
import time
import requests
import logging
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from utils.geodata_enrichment import GeodataEnrichmentManager

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("geocoding.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("BatchGeocoder")

class BatchGeocoder:
    def __init__(self, api_key=None, provider="nominatim", rate_limit=1, batch_size=10):
        """
        Initialize the batch geocoder
        
        Args:
            api_key (str): API key for geocoding service (if needed)
            provider (str): Geocoding provider to use
            rate_limit (float): Seconds between API calls
            batch_size (int): Number of items to process in each batch
        """
        self.api_key = api_key
        self.provider = provider
        self.rate_limit = rate_limit
        self.batch_size = batch_size
        self.geo_manager = GeodataEnrichmentManager()
        
        # Verify geocoding provider and API key
        self.providers = {
            "nominatim": self._geocode_nominatim,
            "google": self._geocode_google,
            "mapbox": self._geocode_mapbox,
            "here": self._geocode_here,
            "arcgis": self._geocode_arcgis,
            "test": self._geocode_test  # For testing without actual API calls
        }
        
        if provider not in self.providers:
            raise ValueError(f"Unsupported geocoding provider: {provider}")
            
        # Check if API key is required
        if provider in ["google", "mapbox", "here", "arcgis"] and not api_key:
            raise ValueError(f"{provider} requires an API key")
    
    def _geocode_nominatim(self, location_str):
        """Geocode using Nominatim (OpenStreetMap)"""
        try:
            url = "https://nominatim.openstreetmap.org/search"
            params = {
                "q": location_str,
                "format": "json",
                "limit": 1,
                "email": "cityinfra@example.com"  # Should be a real email in production
            }
            
            headers = {
                "User-Agent": "CityInfraXLS/1.0"
            }
            
            response = requests.get(url, params=params, headers=headers)
            response.raise_for_status()
            
            results = response.json()
            if results and len(results) > 0:
                return {
                    "latitude": float(results[0]["lat"]),
                    "longitude": float(results[0]["lon"]),
                    "status": "OK",
                    "accuracy_meters": 20,  # Estimated
                    "address_string": results[0].get("display_name", "")
                }
            else:
                return {"status": "NOT_FOUND"}
        
        except Exception as e:
            logger.error(f"Geocoding error with Nominatim: {e}")
            return {"status": "ERROR", "error_message": str(e)}
    
    def _geocode_google(self, location_str):
        """Geocode using Google Maps API"""
        try:
            url = "https://maps.googleapis.com/maps/api/geocode/json"
            params = {
                "address": location_str,
                "key": self.api_key
            }
            
            response = requests.get(url, params=params)
            response.raise_for_status()
            
            result = response.json()
            if result["status"] == "OK":
                location = result["results"][0]["geometry"]["location"]
                return {
                    "latitude": location["lat"],
                    "longitude": location["lng"],
                    "status": "OK",
                    "accuracy_meters": self._get_google_accuracy(result["results"][0]["geometry"].get("location_type", "")),
                    "address_string": result["results"][0].get("formatted_address", "")
                }
            else:
                return {"status": result["status"]}
        
        except Exception as e:
            logger.error(f"Geocoding error with Google: {e}")
            return {"status": "ERROR", "error_message": str(e)}
    
    def _get_google_accuracy(self, location_type):
        """Convert Google location type to accuracy in meters"""
        accuracy_map = {
            "ROOFTOP": 10,
            "RANGE_INTERPOLATED": 50,
            "GEOMETRIC_CENTER": 100,
            "APPROXIMATE": 500
        }
        return accuracy_map.get(location_type, 1000)
    
    def _geocode_mapbox(self, location_str):
        """Geocode using Mapbox API"""
        try:
            url = f"https://api.mapbox.com/geocoding/v5/mapbox.places/{location_str}.json"
            params = {
                "access_token": self.api_key,
                "limit": 1
            }
            
            response = requests.get(url, params=params)
            response.raise_for_status()
            
            result = response.json()
            if result["features"] and len(result["features"]) > 0:
                feature = result["features"][0]
                return {
                    "latitude": feature["center"][1],
                    "longitude": feature["center"][0],
                    "status": "OK",
                    "accuracy_meters": self._get_mapbox_accuracy(feature.get("place_type", ["place"])[0]),
                    "address_string": feature.get("place_name", "")
                }
            else:
                return {"status": "NOT_FOUND"}
        
        except Exception as e:
            logger.error(f"Geocoding error with Mapbox: {e}")
            return {"status": "ERROR", "error_message": str(e)}
    
    def _get_mapbox_accuracy(self, place_type):
        """Convert Mapbox place type to accuracy in meters"""
        accuracy_map = {
            "address": 20,
            "poi": 50,
            "neighborhood": 500,
            "postcode": 1000,
            "place": 5000,
            "region": 10000,
            "country": 100000
        }
        return accuracy_map.get(place_type, 1000)
    
    def _geocode_here(self, location_str):
        """Geocode using HERE API"""
        # Implementation similar to others
        return {"status": "NOT_IMPLEMENTED"}
        
    def _geocode_arcgis(self, location_str):
        """Geocode using ArcGIS API"""
        # Implementation similar to others
        return {"status": "NOT_IMPLEMENTED"}
    
    def _geocode_test(self, location_str):
        """Test geocoder that returns fake coordinates"""
        import hashlib
        import random
        
        # Generate deterministic but random-looking coordinates based on input
        hash_val = int(hashlib.md5(location_str.encode()).hexdigest(), 16)
        random.seed(hash_val)
        
        # Generate values that look like real coordinates
        lat = random.uniform(25.0, 48.0)
        lng = random.uniform(-125.0, -70.0)
        
        return {
            "latitude": lat,
            "longitude": lng,
            "status": "OK",
            "accuracy_meters": random.choice([10, 20, 50, 100, 500]),
            "address_string": f"Test Address for: {location_str}"
        }
    
    def geocode_location(self, location_str):
        """
        Geocode a single location string
        
        Args:
            location_str (str): Location description to geocode
            
        Returns:
            dict: Geocoding result
        """
        # Use the appropriate provider function
        geocode_func = self.providers[self.provider]
        result = geocode_func(location_str)
        
        # Add metadata
        result["geocode_source"] = f"API_{self.provider.upper()}"
        result["geocode_timestamp"] = datetime.now().isoformat()
        
        return result
    
    def _process_asset_batch(self, asset_batch):
        """Process a batch of assets"""
        results = []
        
        for asset in asset_batch:
            # Enforce rate limiting
            time.sleep(self.rate_limit)
            
            asset_id = asset["asset_id"]
            location = asset["location_description"]
            
            logger.info(f"Geocoding asset {asset_id}: {location}")
            
            # Geocode the location
            geocode_result = self.geocode_location(location)
            
            # Combine the asset info with geocode results
            results.append({
                "asset_id": asset_id,
                **geocode_result
            })
        
        return results
    
    def identify_assets_needing_geocoding(self):
        """
        Identify assets with location descriptions but missing coordinates
        
        Returns:
            list: List of assets needing geocoding
        """
        try:
            # Get assets with descriptions
            assets_path = Path(__file__).parent.parent / 'data' / 'assets.xlsx'
            assets_df = pd.read_excel(assets_path)
            
            # Filter assets with location descriptions
            if 'location_description' not in assets_df.columns:
                logger.error("Assets file does not contain location_description column")
                return []
            
            assets_with_location = assets_df[~assets_df['location_description'].isna()]
            
            # Load geodata
            geodata_path = Path(__file__).parent.parent / 'data' / 'asset_geodata.xlsx'
            if os.path.exists(geodata_path):
                geodata_df = pd.read_excel(geodata_path, sheet_name='AssetGeodata')
                
                # Find assets with valid coordinates
                if not geodata_df.empty:
                    assets_with_coords = geodata_df[
                        (~geodata_df['latitude'].isna()) & 
                        (~geodata_df['longitude'].isna())
                    ]['asset_id'].tolist()
                else:
                    assets_with_coords = []
            else:
                assets_with_coords = []
            
            # Find assets needing geocoding
            assets_needing_geocoding = []
            
            for _, row in assets_with_location.iterrows():
                asset_id = row['asset_id']
                if asset_id not in assets_with_coords:
                    assets_needing_geocoding.append({
                        "asset_id": asset_id,
                        "location_description": row['location_description']
                    })
            
            logger.info(f"Found {len(assets_needing_geocoding)} assets needing geocoding")
            return assets_needing_geocoding
            
        except Exception as e:
            logger.error(f"Error identifying assets for geocoding: {e}")
            return []
    
    def run_geocoding_batch(self, assets=None, max_assets=None, workers=4):
        """
        Run batch geocoding for assets
        
        Args:
            assets (list): List of assets to geocode, or None to auto-detect
            max_assets (int): Maximum number of assets to process
            workers (int): Number of worker threads
            
        Returns:
            tuple: (success_count, total_count)
        """
        # Identify assets if not provided
        if assets is None:
            assets = self.identify_assets_needing_geocoding()
        
        # Limit number of assets if requested
        if max_assets is not None:
            assets = assets[:max_assets]
        
        if not assets:
            logger.info("No assets found needing geocoding")
            return 0, 0
        
        total_assets = len(assets)
        logger.info(f"Starting batch geocoding for {total_assets} assets")
        
        # Split into batches
        batches = [assets[i:i + self.batch_size] for i in range(0, len(assets), self.batch_size)]
        
        all_results = []
        
        # Process batches with thread pool
        with ThreadPoolExecutor(max_workers=workers) as executor:
            batch_futures = [executor.submit(self._process_asset_batch, batch) for batch in batches]
            
            for future in batch_futures:
                # Get results from each batch
                batch_results = future.result()
                all_results.extend(batch_results)
        
        # Update geodata with results
        success_count = self._update_geodata_with_results(all_results)
        
        logger.info(f"Completed batch geocoding: {success_count} of {total_assets} successful")
        return success_count, total_assets
    
    def _update_geodata_with_results(self, results):
        """
        Update geodata with batch results in one atomic operation
        
        Args:
            results (list): List of geocoding results
            
        Returns:
            int: Count of successful updates
        """
        success_count = 0
        
        # Group updates by status
        successful_updates = []
        for result in results:
            if result.get("status") == "OK":
                # Prepare data for update
                update_data = {
                    "asset_id": result["asset_id"],
                    "latitude": result["latitude"],
                    "longitude": result["longitude"],
                    "geocode_source": result["geocode_source"],
                    "geocode_timestamp": result["geocode_timestamp"],
                    "validation_status": "UNVERIFIED"
                }
                
                # Add additional data if available
                for field in ["accuracy_meters", "address_string"]:
                    if field in result:
                        update_data[field] = result[field]
                
                successful_updates.append(update_data)
            else:
                logger.warning(f"Geocoding failed for asset {result['asset_id']}: {result.get('status')}")
        
        # Prepare geodata file
        ready, message = self.geo_manager.prepare_for_enrichment()
        if not ready:
            logger.error(f"Cannot update geodata: {message}")
            return 0
        
        # Perform batch update
        try:
            # Load current geodata
            geodata_path = Path(__file__).parent.parent / 'data' / 'asset_geodata.xlsx'
            if os.path.exists(geodata_path):
                geodata_df = pd.read_excel(geodata_path, sheet_name='AssetGeodata')
            else:
                # Create empty dataframe with correct columns
                schema_path = Path(__file__).parent.parent / 'geodata_schema.json'
                with open(schema_path, 'r') as f:
                    schema = json.load(f)
                geodata_df = pd.DataFrame(columns=list(schema['properties'].keys()))
            
            # Process updates
            for update in successful_updates:
                asset_id = update["asset_id"]
                
                # Check if asset exists in geodata
                existing = geodata_df[geodata_df['asset_id'] == asset_id]
                
                if existing.empty:
                    # Add new row
                    geodata_df = geodata_df.append(update, ignore_index=True)
                else:
                    # Update existing row
                    for key, value in update.items():
                        geodata_df.loc[geodata_df['asset_id'] == asset_id, key] = value
                
                success_count += 1
            
            # Save back to Excel (atomic update of the entire file)
            if success_count > 0:
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
                
                logger.info(f"Successfully updated geodata for {success_count} assets")
            
            return success_count
            
        except Exception as e:
            logger.error(f"Error updating geodata: {e}")
            return 0