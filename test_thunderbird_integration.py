#!/usr/bin/env python3
"""
Test script for Thunderbird integration fixes.
This script validates that the native messaging setup works correctly.
"""

import os
import sys
import json
import time
import uuid
import subprocess
import tempfile
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

try:
    from core.email_sender import ThunderbirdExtensionSender
except ImportError as e:
    print(f"Could not import ThunderbirdExtensionSender: {e}")
    print("Make sure you're running this from the project root directory")
    sys.exit(1)

def test_queue_setup():
    """Test that the queue directories are properly set up."""
    print("[*] Testing queue setup...")
    
    queue_dir = os.path.join(os.getcwd(), "tb_queue")
    jobs_dir = os.path.join(queue_dir, "jobs")
    results_dir = os.path.join(queue_dir, "results")
    
    if not os.path.exists(queue_dir):
        print(f"[FAIL] Queue directory missing: {queue_dir}")
        return False
    
    if not os.path.exists(jobs_dir):
        print(f"[FAIL] Jobs directory missing: {jobs_dir}")
        return False
    
    if not os.path.exists(results_dir):
        print(f"[FAIL] Results directory missing: {results_dir}")
        return False
    
    print(f"[PASS] Queue directories exist at: {queue_dir}")
    return True

def test_native_host_files():
    """Test that native host files are properly configured."""
    print("[*] Testing native host files...")
    
    extension_dir = os.path.join(os.getcwd(), "thunderbird_extension")
    
    # Check for required files
    required_files = [
        "native_host.py",
        "native_host_windows.py", 
        "native_host_launcher.bat",
        "background.js",
        "manifest.json",
        "native_manifest_windows.json"
    ]
    
    for filename in required_files:
        filepath = os.path.join(extension_dir, filename)
        if not os.path.exists(filepath):
            print(f"[FAIL] Missing file: {filepath}")
            return False
    
    print("[PASS] All required native host files exist")
    return True

def test_registry_setup():
    """Test that Windows registry keys are properly set."""
    print("[*] Testing Windows registry setup...")
    
    try:
        import winreg
        
        registry_paths = [
            r"SOFTWARE\Mozilla\NativeMessagingHosts\com.emailautomation.tbhost",
            r"SOFTWARE\Thunderbird\NativeMessagingHosts\com.emailautomation.tbhost"
        ]
        
        for reg_path in registry_paths:
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path)
                manifest_path = winreg.QueryValueEx(key, None)[0]
                winreg.CloseKey(key)
                
                if os.path.exists(manifest_path):
                    print(f"[PASS] Registry key exists: {reg_path} -> {manifest_path}")
                else:
                    print(f"[WARN] Registry key exists but manifest file missing: {manifest_path}")
            except FileNotFoundError:
                print(f"[FAIL] Registry key missing: {reg_path}")
                return False
    except ImportError:
        print("[WARN] Cannot test registry on non-Windows platform")
        return True
    
    return True

def test_thunderbird_sender():
    """Test the ThunderbirdExtensionSender class."""
    print("[*] Testing ThunderbirdExtensionSender...")
    
    try:
        sender = ThunderbirdExtensionSender(
            queue_dir=os.path.join(os.getcwd(), "tb_queue"),
            timeout_seconds=5.0  # Short timeout for testing
        )
        print(f"[PASS] ThunderbirdExtensionSender created successfully")
        print(f"   Queue dir: {sender.queue_dir}")
        print(f"   Jobs dir: {sender.jobs_dir}")
        print(f"   Results dir: {sender.results_dir}")
        return True
    except Exception as e:
        print(f"[FAIL] Failed to create ThunderbirdExtensionSender: {e}")
        return False

def test_job_creation():
    """Test creating a job file without sending."""
    print("[*] Testing job file creation...")
    
    try:
        sender = ThunderbirdExtensionSender(
            queue_dir=os.path.join(os.getcwd(), "tb_queue"),
            timeout_seconds=5.0
        )
        
        # Create a test job file manually (don't send)
        job_id = str(uuid.uuid4())
        job = {
            "id": job_id,
            "type": "sendEmail",
            "payload": {
                "to": ["test@example.com"],
                "cc": [],
                "bcc": [],
                "subject": "Test Email",
                "bodyHtml": "<p>This is a test email</p>",
                "attachments": []
            },
            "meta": {
                "createdAt": int(time.time()),
                "client": "email_automation_test"
            }
        }
        
        job_path = os.path.join(sender.jobs_dir, f"{job_id}.json")
        
        with open(job_path, "w", encoding="utf-8") as f:
            json.dump(job, f, ensure_ascii=False, indent=2)
        
        print(f"[PASS] Test job file created: {job_path}")
        
        # Clean up
        time.sleep(0.5)
        if os.path.exists(job_path):
            os.remove(job_path)
            print("[PASS] Test job file cleaned up")
        
        return True
    except Exception as e:
        print(f"[FAIL] Failed to create test job: {e}")
        return False

def test_native_host_launch():
    """Test launching the native host manually."""
    print("[*] Testing native host launch...")
    
    try:
        launcher_path = os.path.join(os.getcwd(), "thunderbird_extension", "native_host_launcher.bat")
        
        if not os.path.exists(launcher_path):
            print(f"[FAIL] Launcher not found: {launcher_path}")
            return False
        
        print(f"[PASS] Native host launcher exists: {launcher_path}")
        print("[INFO] To test manually, run the launcher and check for native_host.log")
        
        return True
    except Exception as e:
        print(f"[FAIL] Error checking native host launcher: {e}")
        return False

def run_all_tests():
    """Run all tests and report results."""
    print("=" * 60)
    print("THUNDERBIRD INTEGRATION TEST SUITE")
    print("=" * 60)
    
    tests = [
        ("Queue Setup", test_queue_setup),
        ("Native Host Files", test_native_host_files),
        ("Windows Registry", test_registry_setup),
        ("ThunderbirdExtensionSender", test_thunderbird_sender),
        ("Job File Creation", test_job_creation),
        ("Native Host Launch", test_native_host_launch),
    ]
    
    results = {}
    
    for test_name, test_func in tests:
        print()
        try:
            results[test_name] = test_func()
        except Exception as e:
            print(f"[FAIL] {test_name} failed with exception: {e}")
            results[test_name] = False
    
    print()
    print("=" * 60)
    print("TEST RESULTS SUMMARY")
    print("=" * 60)
    
    passed = 0
    failed = 0
    
    for test_name, result in results.items():
        status = "[PASS]" if result else "[FAIL]"
        print(f"{test_name:<25} {status}")
        if result:
            passed += 1
        else:
            failed += 1
    
    print()
    print(f"Tests passed: {passed}/{len(tests)}")
    
    if failed == 0:
        print("ALL TESTS PASSED!")
        print()
        print("Next steps:")
        print("1. Open Thunderbird")
        print("2. Go to Add-ons and Themes -> Tools (gear) -> Debug Add-ons")
        print("3. Load Temporary Add-on -> Select: thunderbird_extension/manifest.json")
        print("4. Keep Thunderbird running")
        print("5. In the email app, set email_client to 'thunderbird'")
        print("6. Test sending an email through the app")
        return True
    else:
        print(f"WARNING: {failed} test(s) failed - please fix these issues before proceeding")
        return False

if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)