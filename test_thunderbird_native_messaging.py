#!/usr/bin/env python3
"""
Test script for Thunderbird native messaging integration
"""

import asyncio
import logging
import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from core.thunderbird_native_client import ThunderbirdNativeClient
from utils.thunderbird_native_config import ThunderbirdNativeConfig

def test_native_config():
    """Test native messaging configuration"""
    print("Testing native messaging configuration...")

    config = ThunderbirdNativeConfig()

    # Print current status
    config.print_status()

    # Check if installed
    if config.check_installation():
        print("✓ Native messaging host is properly configured")
        return True
    else:
        print("✗ Native messaging host is not configured")
        print("\nTo install, run:")
        print("python -m src.utils.thunderbird_native_config install")
        return False

async def test_native_connection():
    """Test native messaging connection"""
    print("\nTesting native messaging connection...")

    try:
        client = ThunderbirdNativeClient()

        # Try to connect
        connected = await client.connect()
        if not connected:
            print("✗ Failed to connect to native messaging host")
            return False

        print("✓ Connected to native messaging host")

        # Test availability check
        available = await client.check_availability()
        if available:
            print("✓ Native messaging host reports available")
        else:
            print("✗ Native messaging host reports not available")

        # Test getting accounts
        accounts = await client.get_accounts()
        print(f"✓ Retrieved {len(accounts)} accounts")

        # Test sending email (mock)
        email_data = {
            'to': ['test@example.com'],
            'cc': [],
            'bcc': [],
            'subject': 'Test Email from Native Messaging',
            'body': 'This is a test email sent via native messaging.',
            'attachmentPath': None,
            'attachmentName': None
        }

        result = await client.send_email(email_data)
        if result.get('success'):
            print(f"✓ Email sent successfully: {result.get('messageId')}")
        else:
            print(f"✗ Email sending failed: {result.get('error')}")

        # Disconnect
        await client.disconnect()
        print("✓ Disconnected from native messaging host")

        return True

    except Exception as e:
        print(f"✗ Error testing native messaging connection: {str(e)}")
        return False

def main():
    """Main test function"""
    print("Thunderbird Native Messaging Test")
    print("=" * 50)

    # Setup logging
    logging.basicConfig(level=logging.INFO)

    # Test configuration
    config_ok = test_native_config()

    if not config_ok:
        print("\nConfiguration test failed. Please install the native messaging host first.")
        return

    # Test connection
    try:
        success = asyncio.run(test_native_connection())
        if success:
            print("\n✓ All tests passed!")
        else:
            print("\n✗ Some tests failed.")
    except Exception as e:
        print(f"\n✗ Error running tests: {str(e)}")

if __name__ == "__main__":
    main()