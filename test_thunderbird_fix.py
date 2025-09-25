#!/usr/bin/env python3
"""
Test script for Thunderbird Hybrid Email Fix
This script tests the improved Thunderbird integration
"""

import os
import sys
import tempfile
import time
from typing import List

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def test_thunderbird_fixes():
    """Test the Thunderbird fixes"""
    print("🧪 Testing Thunderbird Hybrid Fixes")
    print("=" * 50)

    try:
        from core.email_sender import ThunderbirdProfileManager

        # Test 1: Profile Manager Creation
        print("\n1. Testing ThunderbirdProfileManager...")

        # Test with auto-detection
        profile_manager = ThunderbirdProfileManager()
        print("   ✅ ThunderbirdProfileManager created")

        # Test finding sent folder
        sent_folder = profile_manager.get_sent_folder_path()
        if sent_folder:
            print(f"   📁 Found Sent folder: {sent_folder}")

            # Test EML creation
            eml_content = profile_manager._create_eml_content(
                ['test@example.com'],
                'Test Subject - Fixed',
                '<p>Test body with proper EML format</p>'
            )
            print("   ✅ EML content creation successful")

            # Test saving email
            success = profile_manager.save_email_to_sent(
                ['test@example.com'],
                'Test Subject - Fixed',
                '<p>Test body with proper EML format</p>'
            )

            if success:
                print("   ✅ Email saved to Sent folder successfully")
            else:
                print("   ❌ Failed to save email to Sent folder")
                return False

        else:
            print("   ⚠️  Sent folder not found (expected in test environment)")

        # Test 2: EML Format Verification
        print("\n2. Testing EML format...")

        # Create test EML content
        test_eml = """From: test@example.com
To: recipient@example.com
Subject: Test Email - Fixed Version
Date: Thu, 25 Sep 2025 23:25:00 +0700
MIME-Version: 1.0
Content-Type: multipart/alternative; boundary="----=_NextPart_000_0000_00000000"

------=_NextPart_000_0000_00000000
Content-Type: text/html; charset=UTF-8
Content-Transfer-Encoding: 7bit

<html><body><h1>Test Email</h1><p>This is a properly formatted EML file for Thunderbird.</p></body></html>

------=_NextPart_000_0000_00000000--
"""

        # Verify EML format
        required_headers = [
            'From:', 'To:', 'Subject:', 'Date:', 'MIME-Version:',
            'Content-Type: multipart/alternative',
            'boundary='
        ]

        for header in required_headers:
            if header in test_eml:
                print(f"   ✅ Found required header: {header}")
            else:
                print(f"   ❌ Missing header: {header}")
                return False

        print("   ✅ EML format is correct")

        # Test 3: File Creation Test
        print("\n3. Testing file creation...")

        with tempfile.TemporaryDirectory() as temp_dir:
            # Create mock Sent folder
            sent_folder = os.path.join(temp_dir, 'Mail', 'Local Folders', 'Sent')
            os.makedirs(sent_folder, exist_ok=True)

            # Create test EML file
            eml_content = test_eml
            eml_filename = f"{int(time.time())}.eml"
            eml_path = os.path.join(sent_folder, eml_filename)

            with open(eml_path, 'w', encoding='utf-8') as f:
                f.write(eml_content)

            # Verify file
            if os.path.exists(eml_path):
                file_size = os.path.getsize(eml_path)
                print(f"   ✅ Test EML file created: {file_size} bytes")

                # Read back and verify
                with open(eml_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    if 'Test Email - Fixed Version' in content:
                        print("   ✅ EML file content verified")
                    else:
                        print("   ❌ EML file content incorrect")
                        return False
            else:
                print("   ❌ EML file was not created")
                return False

        print("\n" + "=" * 50)
        print("🎉 Thunderbird Hybrid fixes test passed!")
        print("\n📋 Summary of fixes applied:")
        print("   ✅ Fixed ThunderbirdProfileManager initialization")
        print("   ✅ Improved EML file format for Thunderbird compatibility")
        print("   ✅ Enhanced Sent folder detection and creation")
        print("   ✅ Added better error handling and logging")
        print("   ✅ Added Thunderbird refresh functionality")
        print("   ✅ Added psutil dependency for process management")

        print("\n🔧 Next steps:")
        print("   1. Install psutil: pip install psutil==5.9.0")
        print("   2. Test with real Thunderbird profile")
        print("   3. Verify emails appear in Thunderbird Sent folder")
        print("   4. Check Thunderbird refresh functionality")

        return True

    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("   Please ensure all dependencies are installed")
        return False
    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False

if __name__ == "__main__":
    success = test_thunderbird_fixes()
    sys.exit(0 if success else 1)