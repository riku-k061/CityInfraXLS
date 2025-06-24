import pytest
import time
import requests
from datetime import datetime
from utils.batch_geocoder import BatchGeocoder

class DummyResponse:
    def __init__(self, json_data, status_code=200):
        self._json = json_data
        self.status_code = status_code

    def json(self):
        return self._json

    def raise_for_status(self):
        if self.status_code != 200:
            raise requests.HTTPError(f"Status {self.status_code}")

@pytest.fixture
def geocoder_test():
    # Use the "test" provider to avoid external calls
    return BatchGeocoder(provider="test", rate_limit=0)

def test_init_invalid_provider():
    with pytest.raises(ValueError):
        BatchGeocoder(provider="no_such_provider")

def test_init_missing_api_key_for_google():
    with pytest.raises(ValueError):
        BatchGeocoder(provider="google")

def test_get_google_accuracy_mapping():
    bg = BatchGeocoder(provider="test")
    assert bg._get_google_accuracy("ROOFTOP") == 10
    assert bg._get_google_accuracy("UNKNOWN_TYPE") == 1000

def test_get_mapbox_accuracy_mapping():
    bg = BatchGeocoder(provider="test")
    assert bg._get_mapbox_accuracy("address") == 20
    assert bg._get_mapbox_accuracy("country") == 100000
    assert bg._get_mapbox_accuracy("foo") == 1000

def test_geocode_location_test_provider_consistency(geocoder_test):
    loc = "123 Main St"
    r1 = geocoder_test.geocode_location(loc)
    r2 = geocoder_test.geocode_location(loc)
    # test provider should be deterministic
    assert r1["status"] == "OK" and r2["status"] == "OK"
    assert r1["latitude"] == r2["latitude"]
    assert r1["longitude"] == r2["longitude"]
    assert r1["geocode_source"] == "API_TEST"
    assert "geocode_timestamp" in r1
    # timestamp should parse as ISO
    datetime.fromisoformat(r1["geocode_timestamp"])

def test_process_asset_batch(geocoder_test):
    batch = [
        {"asset_id": "A1", "location_description": "LocA"},
        {"asset_id": "B2", "location_description": "LocB"}
    ]
    results = geocoder_test._process_asset_batch(batch)
    assert isinstance(results, list) and len(results) == 2
    for res, asset in zip(results, batch):
        assert res["asset_id"] == asset["asset_id"]
        assert res["status"] == "OK"
        assert "latitude" in res and "longitude" in res
        assert res["geocode_source"] == "API_TEST"

def test_run_geocoding_batch_empty(monkeypatch):
    bg = BatchGeocoder(provider="test", rate_limit=0, batch_size=2)
    # No assets needing geocoding
    monkeypatch.setattr(BatchGeocoder, "identify_assets_needing_geocoding", lambda self: [])
    success, total = bg.run_geocoding_batch()
    assert success == 0 and total == 0

def test_run_geocoding_batch_full_cycle(monkeypatch):
    bg = BatchGeocoder(provider="test", rate_limit=0, batch_size=2)
    dummy_assets = [
        {"asset_id": "X1", "location_description": "LocX"},
        {"asset_id": "Y2", "location_description": "LocY"},
        {"asset_id": "Z3", "location_description": "LocZ"}
    ]
    # Stub identification to return our dummy list
    monkeypatch.setattr(BatchGeocoder, "identify_assets_needing_geocoding", lambda self: dummy_assets)
    # Stub update to simply return the number of OK results
    monkeypatch.setattr(BatchGeocoder, "_update_geodata_with_results", lambda self, results: len(results))
    success, total = bg.run_geocoding_batch()
    assert total == 3
    assert success == 3

def test_geocode_nominatim_success(monkeypatch):
    bg = BatchGeocoder(provider="test")
    # Monkey-patch requests.get for nominatim
    dummy = [{
        "lat": "12.34",
        "lon": "56.78",
        "display_name": "Test Place"
    }]
    monkeypatch.setattr(requests, "get", lambda url, params, headers: DummyResponse(dummy))
    out = bg._geocode_nominatim("anything")
    assert out["status"] == "OK"
    assert out["latitude"] == 12.34
    assert out["longitude"] == 56.78
    assert out["address_string"] == "Test Place"

def test_geocode_nominatim_not_found(monkeypatch):
    bg = BatchGeocoder(provider="test")
    monkeypatch.setattr(requests, "get", lambda *args, **kwargs: DummyResponse([], status_code=200))
    out = bg._geocode_nominatim("nothing")
    assert out["status"] == "NOT_FOUND"

def test_geocode_nominatim_error(monkeypatch):
    bg = BatchGeocoder(provider="test")
    def raise_exc(*args, **kwargs):
        raise Exception("Network fail")
    monkeypatch.setattr(requests, "get", raise_exc)
    out = bg._geocode_nominatim("fail")
    assert out["status"] == "ERROR"
    assert "Network fail" in out["error_message"]
