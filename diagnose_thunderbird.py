#!/usr/bin/env python3
"""
Thunderbird WebExtension Diagnostic Script
Tests WebExtension installation, connection, and functionality
"""

import os
import sys
import time
import asyncio
import logging
from typing import Dict, Any

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def setup_logging():
    """Setup logging for diagnostic script"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('thunderbird_diagnostic.log')
        ]
    )
    return logging.getLogger(__name__)

def test_webextension_installation() -> Dict[str, Any]:
    """Test WebExtension installation status"""
    print("\n🔍 Testing WebExtension Installation")
    print("=" * 50)

    try:
        from src.utils.thunderbird_extension_manager import ThunderbirdExtensionManager

        manager = ThunderbirdExtensionManager()
        extension_info = manager.get_extension_info()

        if extension_info:
            print(f"✅ Extension info loaded: {extension_info.get('name')} v{extension_info.get('version')}")

        # Test environment validation
        validation = manager.validate_environment()

        results = {
            'extension_info': extension_info,
            'validation': validation,
            'issues': []
        }

        print("\n📋 Environment Validation Results:")
        for key, value in validation.items():
            if key == 'issues':
                if value:
                    print(f"⚠️  Issues found: {len(value)}")
                    for issue in value:
                        print(f"   - {issue}")
                    results['issues'].extend(value)
                else:
                    print("✅ No validation issues")
            else:
                status = "✅" if value else "❌"
                print(f"{status} {key}: {value}")

        # Test extension building
        print("\n🔨 Testing Extension Building...")
        xpi_path = manager.build_extension()
        if xpi_path and os.path.exists(xpi_path):
            file_size = os.path.getsize(xpi_path)
            print(f"✅ Extension built successfully: {xpi_path}")
            print(f"📦 Extension size: {file_size} bytes")
            results['xpi_path'] = xpi_path
        else:
            print("❌ Failed to build extension")
            results['issues'].append("Extension build failed")

        # Test installation status
        print("\n📦 Testing Installation Status...")
        is_installed = manager.is_installed()
        if is_installed:
            print("✅ Extension is installed")
            results['installed'] = True
        else:
            print("❌ Extension is not installed")
            results['installed'] = False
            results['issues'].append("Extension not installed")

        return results

    except Exception as e:
        print(f"❌ Installation test failed: {e}")
        return {'error': str(e), 'issues': [str(e)]}

def test_websocket_connection() -> Dict[str, Any]:
    """Test WebSocket connection"""
    print("\n🌐 Testing WebSocket Connection")
    print("=" * 50)

    try:
        from src.core.thunderbird_extension_client import ThunderbirdExtensionClient

        client = ThunderbirdExtensionClient()
        results = {'client_created': True}

        print("✅ WebExtension client created")

        # Start WebSocket server
        print("🚀 Starting WebSocket server...")
        client.start_websocket_server()
        print("✅ WebSocket server started on port 8765")
        results['server_started'] = True

        # Test connection
        print("🔌 Testing connection...")
        async def test_connection():
            try:
                connected = await client.connect()
                return connected
            except Exception as e:
                print(f"❌ Connection failed: {str(e)}")
                return False

        connection_result = asyncio.run(test_connection())

        if connection_result:
            print("✅ WebSocket connection established")
            results['connected'] = True

            # Test availability check
            print("🔍 Testing availability check...")
            async def test_availability():
                try:
                    available = await client.check_availability()
                    return available
                except Exception as e:
                    print(f"⚠️  Availability check failed: {str(e)}")
                    return False

            availability_result = asyncio.run(test_availability())

            if availability_result:
                print("✅ WebExtension is available")
                results['available'] = True
            else:
                print("❌ WebExtension is not available")
                results['available'] = False
                results['issues'].append("WebExtension not available")

            # Cleanup
            async def cleanup():
                await client.disconnect()

            asyncio.run(cleanup())

        else:
            print("❌ WebSocket connection failed")
            results['connected'] = False
            results['issues'].append("WebSocket connection failed")

        return results

    except Exception as e:
        print(f"❌ WebSocket test failed: {e}")
        return {'error': str(e), 'issues': [str(e)]}

def test_email_sender() -> Dict[str, Any]:
    """Test email sender functionality"""
    print("\n📧 Testing Email Sender")
    print("=" * 50)

    try:
        from src.core.thunderbird_extension_client import ThunderbirdExtensionClient
        from src.core.email_sender import ThunderbirdSender

        # Create client and start server
        client = ThunderbirdExtensionClient()
        client.start_websocket_server()

        # Create sender
        sender = ThunderbirdSender(extension_client=client)
        results = {'sender_created': True}

        print("✅ ThunderbirdSender created")

        # Test connection
        print("🔌 Testing sender connection...")
        connection_result = sender.test_connection()

        if connection_result:
            print("✅ Sender connection test passed")
            results['connection_test'] = True
        else:
            print("❌ Sender connection test failed")
            results['connection_test'] = False
            results['issues'].append("Sender connection test failed")

        # Test with mock email data
        print("📝 Testing mock email send...")
        async def test_mock_send():
            try:
                # This will fail in test environment but tests the code path
                result = await sender.send_email(
                    to_emails=['test@example.com'],
                    subject='Test Email - Diagnostic',
                    body='<p>This is a diagnostic test email</p>'
                )

                return result
            except Exception as e:
                print(f"⚠️  Mock send failed (expected): {str(e)}")
                return False

        mock_result = asyncio.run(test_mock_send())

        if mock_result:
            print("✅ Mock email send succeeded")
            results['mock_send'] = True
        else:
            print("⚠️  Mock email send failed (expected in test environment)")
            results['mock_send'] = False

        return results

    except Exception as e:
        print(f"❌ Email sender test failed: {e}")
        return {'error': str(e), 'issues': [str(e)]}

def main():
    """Run all diagnostic tests"""
    print("🔧 Thunderbird WebExtension Diagnostic Tool")
    print("=" * 60)

    logger = setup_logging()
    logger.info("Starting Thunderbird WebExtension diagnostic")

    all_results = {}

    # Test 1: Installation
    install_results = test_webextension_installation()
    all_results['installation'] = install_results

    # Test 2: WebSocket Connection
    websocket_results = test_websocket_connection()
    all_results['websocket'] = websocket_results

    # Test 3: Email Sender
    sender_results = test_email_sender()
    all_results['email_sender'] = sender_results

    # Summary
    print("\n" + "=" * 60)
    print("📋 DIAGNOSTIC SUMMARY")
    print("=" * 60)

    total_issues = 0
    for test_name, results in all_results.items():
        issues = results.get('issues', [])
        total_issues += len(issues)

        print(f"\n{test_name.upper()} TEST:")
        if issues:
            print(f"⚠️  Issues found: {len(issues)}")
            for issue in issues:
                print(f"   - {issue}")
        else:
            print("✅ No issues found")

    print(f"\n📊 OVERALL RESULTS: {total_issues} total issues found")

    if total_issues == 0:
        print("🎉 All tests passed! WebExtension integration should work.")
        print("\nTo use the extension:")
        print("1. Install the generated .xpi file to Thunderbird")
        print("2. Restart Thunderbird")
        print("3. Select 'thunderbird' as email client in the application")
        print("4. Start monitoring to send emails via WebExtension")
    else:
        print("⚠️  Issues found. Please resolve them before using WebExtension integration.")
        print("\nCommon solutions:")
        print("1. Install the WebExtension to Thunderbird")
        print("2. Restart Thunderbird to load the extension")
        print("3. Check firewall settings for port 8765")
        print("4. Ensure Thunderbird 78+ is installed")

    # Show XPI file location if available
    if install_results.get('xpi_path'):
        print(f"\n📦 Extension file: {install_results['xpi_path']}")

    logger.info(f"Diagnostic completed with {total_issues} issues")

    return total_issues == 0

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)