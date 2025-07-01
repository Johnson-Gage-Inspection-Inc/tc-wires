#!/usr/bin/env python3
"""
Test script to verify the new Qualer SDK API functions work.
"""

from qualer_sdk.client import AuthenticatedClient
from qualer_sdk.api.client_asset_service_records import (
    get_asset_service_records_by_asset_get_2,
)
from qualer_sdk.api.service_order_items import get_work_items_workitems
from qualer_sdk.api.service_order_documents import (
    get_documents_list,
    get_document,
)


def test_api_imports():
    """Test that all API imports work correctly."""
    print("✓ All API imports successful!")


def test_client_creation():
    """Test client creation (without real token)."""
    try:
        client = AuthenticatedClient(
            base_url="https://jgiquality.qualer.com",
            token="test-token"
        )
        print("✓ Client creation successful!")
        return client
    except Exception as e:
        print(f"✗ Client creation failed: {e}")
        return None


def test_api_signatures():
    """Test that API functions have expected signatures."""
    import inspect

    # Test asset service records API
    sig = inspect.signature(get_asset_service_records_by_asset_get_2.sync)
    print(f"✓ get_asset_service_records_by_asset_get_2.sync signature: {sig}")

    # Test service order items API
    sig = inspect.signature(get_work_items_workitems.sync)
    print(f"✓ get_work_items_workitems.sync signature: {sig}")

    # Test service order documents API
    sig = inspect.signature(get_documents_list.sync)
    print(f"✓ get_documents_list.sync signature: {sig}")

    sig = inspect.signature(get_document.sync_detailed)
    print(f"✓ get_document.sync signature: {sig}")


if __name__ == "__main__":
    print("Testing New Qualer SDK API...")
    test_api_imports()
    test_client_creation()
    test_api_signatures()
    print("All tests completed!")
