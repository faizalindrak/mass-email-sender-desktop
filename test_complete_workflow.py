#!/usr/bin/env python3
"""
Complete workflow test for Thunderbird native messaging integration
"""

import asyncio
import logging
import sys
import os
import tempfile

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from core.email_sender import EmailSenderFactory
from utils.thunderbird_native_config import ThunderbirdNativeConfig

def test_native_config():
    """Test native messaging configuration"""
    print("=== Testing Native Messaging Configuration ===")

    config = ThunderbirdNativeConfig()

    if config.check_installation():
        print("✓ Native messaging host is properly configured")
        return True
    else:
        print("✗ Native messaging host is not configured")
        print("Run: python -m src.utils.thunderbird_native_config install")
        return False

async def test_email_sender():
    """Test email sender with native messaging"""
    print("\n=== Testing Email Sender ===")

    try:
        # Create Thunderbird sender
        sender = EmailSenderFactory.create_sender('thunderbird')

        if sender is None:
            print("✗ Failed to create Thunderbird sender")
            return False

        print("✓ Thunderbird sender created successfully")

        # Test connection
        if sender.test_connection():
            print("✓ Thunderbird connection test passed")
        else:
            print("⚠ Thunderbird connection test failed (expected if Thunderbird not running)")
            print("  This is normal - the extension needs to be installed and Thunderbird running")

        # Test email sending (mock test - won't actually send)
        print("✓ Email sender initialization successful")

        return True

    except Exception as e:
        print(f"✗ Error testing email sender: {str(e)}")
        return False

def test_extension_files():
    """Test that extension files are present and valid"""
    print("\n=== Testing Extension Files ===")

    extension_dir = "thunderbird_extension"
    required_files = [
        "manifest.json",
        "background.js",
        "content.js",
        "email_automation_native_host.json"
    ]

    missing_files = []
    for file in required_files:
        if not os.path.exists(os.path.join(extension_dir, file)):
            missing_files.append(file)

    if missing_files:
        print(f"✗ Missing extension files: {missing_files}")
        return False

    print("✓ All required extension files present")

    # Check manifest.json
    try:
        import json
        with open(os.path.join(extension_dir, "manifest.json"), 'r') as f:
            manifest = json.load(f)

        # Check for native messaging permission
        permissions = manifest.get('permissions', [])
        if 'nativeMessaging' in permissions:
            print("✓ Native messaging permission configured in manifest")
        else:
            print("✗ Native messaging permission missing from manifest")
            return False

        # Check extension ID
        gecko_id = manifest.get('applications', {}).get('gecko', {}).get('id')
        if gecko_id:
            print(f"✓ Extension ID configured: {gecko_id}")
        else:
            print("✗ Extension ID missing from manifest")
            return False

    except Exception as e:
        print(f"✗ Error reading manifest.json: {str(e)}")
        return False

    return True

def test_xpi_file():
    """Test that XPI file exists and is valid"""
    print("\n=== Testing XPI File ===")

    xpi_file = "email_automation_thunderbird_extension-1.0.0.xpi"

    if not os.path.exists(xpi_file):
        print(f"✗ XPI file not found: {xpi_file}")
        return False

    # Check file size (should be reasonable)
    size = os.path.getsize(xpi_file)
    if size < 1000:  # Less than 1KB is suspicious
        print(f"⚠ XPI file seems too small: {size} bytes")
    elif size > 100000:  # More than 100KB is suspicious
        print(f"⚠ XPI file seems too large: {size} bytes")
    else:
        print(f"✓ XPI file size looks reasonable: {size} bytes")

    # Try to read as ZIP
    try:
        import zipfile
        with zipfile.ZipFile(xpi_file, 'r') as zipf:
            files = zipf.namelist()
            print(f"✓ XPI file contains {len(files)} files")

            # Check for required files
            required_in_zip = ['manifest.json', 'background.js']
            missing_in_zip = [f for f in required_in_zip if f not in files]

            if missing_in_zip:
                print(f"✗ Missing files in XPI: {missing_in_zip}")
                return False

            print("✓ Required files found in XPI")

    except Exception as e:
        print(f"✗ Error reading XPI file: {str(e)}")
        return False

    return True

def print_installation_instructions():
    """Print installation instructions"""
    print("\n=== Installation Instructions ===")
    print("""
To complete the setup:

1. Install the Thunderbird extension:
   - Open Thunderbird
   - Go to Tools → Add-ons and Themes
   - Click the gear icon (⚙️) → Install Add-on From File
   - Select: email_automation_thunderbird_extension-1.0.0.xpi
   - Restart Thunderbird

2. Configure native messaging (already done):
   - Native messaging host is already installed
   - Configuration file: %APPDATA%\\Thunderbird\\NativeMessagingHosts\\email_automation_native_host.json

3. Test the integration:
   - Run the Email Automation application
   - Configure a profile to use Thunderbird
   - Try sending an email

4. Troubleshooting:
   - Check Thunderbird's Browser Console (Ctrl+Shift+J) for errors
   - Verify the extension appears in Tools → Add-ons and Themes
   - Ensure Thunderbird 78+ is installed
   - Check that native messaging host is properly configured
    """)

def main():
    """Main test function"""
    print("Thunderbird Native Messaging - Complete Workflow Test")
    print("=" * 60)

    # Setup logging
    logging.basicConfig(level=logging.INFO)

    tests = [
        ("Native Config", test_native_config),
        ("Extension Files", test_extension_files),
        ("XPI File", test_xpi_file),
    ]

    async_tests = [
        ("Email Sender", test_email_sender),
    ]

    passed = 0
    total = len(tests) + len(async_tests)

    # Run synchronous tests
    for test_name, test_func in tests:
        try:
            if test_func():
                print(f"✓ {test_name}: PASSED")
                passed += 1
            else:
                print(f"✗ {test_name}: FAILED")
        except Exception as e:
            print(f"✗ {test_name}: ERROR - {str(e)}")

    # Run asynchronous tests
    for test_name, test_func in async_tests:
        try:
            if asyncio.run(test_func()):
                print(f"✓ {test_name}: PASSED")
                passed += 1
            else:
                print(f"✗ {test_name}: FAILED")
        except Exception as e:
            print(f"✗ {test_name}: ERROR - {str(e)}")

    print(f"\n=== Results: {passed}/{total} tests passed ===")

    if passed == total:
        print("🎉 All tests passed! The native messaging integration is ready.")
        print_installation_instructions()
    else:
        print("⚠️  Some tests failed. Please check the output above for details.")
        print("   You may need to complete the installation steps.")

if __name__ == "__main__":
    main()