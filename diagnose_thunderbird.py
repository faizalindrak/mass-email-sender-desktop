#!/usr/bin/env python3
"""
Thunderbird Integration Diagnostic Script
This script helps diagnose issues with Thunderbird email saving
"""

import os
import sys
import time
import tempfile
from typing import List

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def diagnose_thunderbird():
    """Diagnose Thunderbird integration issues"""
    print("🔍 Thunderbird Integration Diagnostic")
    print("=" * 50)

    try:
        from core.email_sender import ThunderbirdProfileManager

        # Test 1: Profile Detection
        print("\n1. Testing Thunderbird Profile Detection...")
        profile_manager = ThunderbirdProfileManager()
        print(f"   Profile path: {profile_manager.profile_path}")

        if profile_manager.profile_path:
            print("   ✅ Profile found")
        else:
            print("   ❌ No profile found")
            return False

        # Test 2: Sent Folder Detection
        print("\n2. Testing Sent Folder Detection...")
        sent_folder = profile_manager.get_sent_folder_path()
        print(f"   Sent folder: {sent_folder}")

        if sent_folder:
            print("   ✅ Sent folder found")

            # List existing files
            try:
                files = os.listdir(sent_folder)
                eml_files = [f for f in files if f.endswith('.eml')]
                print(f"   📁 Existing EML files: {len(eml_files)}")

                for eml_file in eml_files[:5]:  # Show first 5
                    file_path = os.path.join(sent_folder, eml_file)
                    file_size = os.path.getsize(file_path)
                    print(f"      - {eml_file} ({file_size} bytes)")
            except Exception as e:
                print(f"   ❌ Cannot list Sent folder: {e}")
        else:
            print("   ❌ Sent folder not found")
            return False

        # Test 3: EML File Creation
        print("\n3. Testing EML File Creation...")
        test_emails = ['test@example.com']
        test_subject = 'Diagnostic Test Email'
        test_body = '<p>This is a diagnostic test email to verify Thunderbird integration.</p>'

        # Create test EML content
        eml_content = profile_manager._create_eml_content(test_emails, test_subject, test_body)
        print("   ✅ EML content created")

        # Save test EML file
        timestamp = int(time.time())
        test_eml_path = os.path.join(sent_folder, f"diagnostic_{timestamp}.eml")

        with open(test_eml_path, 'w', encoding='utf-8') as f:
            f.write(eml_content)

        if os.path.exists(test_eml_path):
            file_size = os.path.getsize(test_eml_path)
            print(f"   ✅ Test EML file created: {file_size} bytes")
            print(f"   📄 File location: {test_eml_path}")

            # Show EML content
            print("   📋 EML Content Preview:")
            lines = eml_content.split('\n')[:10]  # First 10 lines
            for i, line in enumerate(lines, 1):
                print(f"      {i:2d}: {line}")
            if len(eml_content.split('\n')) > 10:
                print(f"      ... ({len(eml_content.split('\n')) - 10} more lines)")
        else:
            print("   ❌ Failed to create test EML file")
            return False

        # Test 4: Thunderbird Process Check
        print("\n4. Checking Thunderbird Process...")
        try:
            import psutil
            thunderbird_processes = []

            for proc in psutil.process_iter(['pid', 'name']):
                if 'thunderbird' in proc.info['name'].lower():
                    thunderbird_processes.append(proc.info)

            if thunderbird_processes:
                print(f"   ✅ Found {len(thunderbird_processes)} Thunderbird process(es):")
                for proc in thunderbird_processes:
                    print(f"      PID {proc['pid']}: {proc['name']}")
            else:
                print("   ⚠️  Thunderbird is not currently running")
                print("   💡 Suggestion: Start Thunderbird to see new emails")
        except ImportError:
            print("   ⚠️  Cannot check processes (psutil not available)")

        # Test 5: MSF File Check
        print("\n5. Checking MSF Index Files...")
        msf_file = sent_folder + ".msf"
        if os.path.exists(msf_file):
            msf_size = os.path.getsize(msf_file)
            msf_mtime = os.path.getmtime(msf_file)
            print(f"   ✅ MSF file exists: {msf_size} bytes, modified: {time.ctime(msf_mtime)}")

            # Try to update MSF file timestamp
            os.utime(msf_file, None)
            print("   ✅ Updated MSF file timestamp")
        else:
            print("   ❌ MSF file not found")
            print("   💡 Suggestion: Thunderbird may need to be started to create MSF files")

        # Test 6: Permissions Check
        print("\n6. Checking File Permissions...")
        try:
            # Test write permission
            test_file = os.path.join(sent_folder, f"permission_test_{timestamp}.tmp")
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            print("   ✅ Write permission OK")
        except Exception as e:
            print(f"   ❌ Write permission failed: {e}")
            return False

        print("\n" + "=" * 50)
        print("🎉 Diagnostic completed successfully!")
        print("\n📋 Summary:")
        print(f"   Profile: {profile_manager.profile_path}")
        print(f"   Sent Folder: {sent_folder}")
        print(f"   Test EML: {test_eml_path}")
        print(f"   Thunderbird Running: {'Yes' if thunderbird_processes else 'No'}")

        print("\n🔧 Next Steps:")
        print("   1. Open Thunderbird and check the Sent folder")
        print("   2. Look for the diagnostic email with subject 'Diagnostic Test Email'")
        print("   3. If you don't see it, try restarting Thunderbird")
        print("   4. Check the logs for any error messages")

        return True

    except Exception as e:
        print(f"❌ Diagnostic failed: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = diagnose_thunderbird()
    sys.exit(0 if success else 1)