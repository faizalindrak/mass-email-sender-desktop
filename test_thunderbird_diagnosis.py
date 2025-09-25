#!/usr/bin/env python3
"""
Diagnostic script to test Thunderbird integration and identify connection issues.
"""

import os
import sys
import json
import time
import subprocess
import tempfile
import shutil
from pathlib import Path

def test_queue_directory():
    """Test if the queue directory is accessible and writable."""
    print("=== Testing Queue Directory ===")
    
    # Try different possible queue directories
    possible_dirs = [
        "tb_queue",  # Project local
        os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "EmailAutomation", "tb_queue"),  # Windows
        os.path.join(os.path.expanduser("~"), ".local", "share", "email_automation", "tb_queue"),  # Linux
        os.path.join(os.path.expanduser("~"), "Library", "Application Support", "EmailAutomation", "tb_queue"),  # macOS
    ]
    
    for queue_dir in possible_dirs:
        print(f"\nTesting queue directory: {queue_dir}")
        try:
            # Create directories
            jobs_dir = os.path.join(queue_dir, "jobs")
            results_dir = os.path.join(queue_dir, "results")
            os.makedirs(jobs_dir, exist_ok=True)
            os.makedirs(results_dir, exist_ok=True)
            
            # Test write permission
            test_file = os.path.join(queue_dir, "test_write.tmp")
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            
            # Test lock file creation
            lock_file = os.path.join(queue_dir, "native_host.lock")
            with open(lock_file, 'w') as f:
                f.write(str(os.getpid()))
            os.remove(lock_file)
            
            print(f"OK: Queue directory {queue_dir} is accessible and writable")
            return queue_dir
            
        except Exception as e:
            print(f"ERROR: Queue directory {queue_dir} failed: {e}")
    
    print("\n✗ No suitable queue directory found")
    return None

def test_native_host_launcher():
    """Test if the native host launcher script exists and is executable."""
    print("\n=== Testing Native Host Launcher ===")
    
    launcher_path = "thunderbird_extension/native_host_launcher.bat"
    if not os.path.exists(launcher_path):
        print(f"ERROR: Launcher script not found: {launcher_path}")
        return False
    
    print(f"OK: Launcher script found: {launcher_path}")
    
    # Check launcher content
    try:
        with open(launcher_path, 'r') as f:
            content = f.read()
        print("Launcher script content:")
        print(content)
        
        # Check if Python path is valid
        if "python.exe" in content:
            python_path = content.split("python.exe")[0] + "python.exe"
            python_path = python_path.strip().strip('"')
            if os.path.exists(python_path):
                print(f"OK: Python executable found: {python_path}")
            else:
                print(f"ERROR: Python executable not found: {python_path}")
                return False
        
        return True
        
    except Exception as e:
        print(f"✗ Error reading launcher script: {e}")
        return False

def test_native_host_process():
    """Test if the native host process can be started."""
    print("\n=== Testing Native Host Process ===")
    
    launcher_path = "thunderbird_extension/native_host_launcher.bat"
    if not os.path.exists(launcher_path):
        print("ERROR: Launcher script not found")
        return False
    
    try:
        # Try to start the native host process
        process = subprocess.Popen(
            [launcher_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Give it a moment to start
        time.sleep(2)
        
        # Check if process is still running
        if process.poll() is None:
            print("OK: Native host process started successfully")
            process.terminate()
            return True
        else:
            stdout, stderr = process.communicate()
            print(f"ERROR: Native host process exited with code {process.returncode}")
            if stdout:
                print(f"stdout: {stdout}")
            if stderr:
                print(f"stderr: {stderr}")
            return False
            
    except Exception as e:
        print(f"✗ Error starting native host process: {e}")
        return False

def test_thunderbird_extension():
    """Test if Thunderbird extension files are properly configured."""
    print("\n=== Testing Thunderbird Extension ===")
    
    # Check manifest.json
    manifest_path = "thunderbird_extension/manifest.json"
    if not os.path.exists(manifest_path):
        print(f"ERROR: Extension manifest not found: {manifest_path}")
        return False
    
    try:
        with open(manifest_path, 'r') as f:
            manifest = json.load(f)
        
        print("OK: Extension manifest found")
        print(f"  Extension ID: {manifest.get('applications', {}).get('gecko', {}).get('id', 'Not found')}")
        print(f"  Version: {manifest.get('version', 'Not found')}")
        print(f"  Permissions: {manifest.get('permissions', [])}")
        
    except Exception as e:
        print(f"ERROR: Error reading extension manifest: {e}")
        return False
    
    # Check native manifest
    native_manifest_path = "thunderbird_extension/native_manifest_windows.json"
    if not os.path.exists(native_manifest_path):
        print(f"ERROR: Native manifest not found: {native_manifest_path}")
        return False
    
    try:
        with open(native_manifest_path, 'r') as f:
            native_manifest = json.load(f)
        
        print("OK: Native manifest found")
        print(f"  Host name: {native_manifest.get('name', 'Not found')}")
        print(f"  Path: {native_manifest.get('path', 'Not found')}")
        print(f"  Type: {native_manifest.get('type', 'Not found')}")
        print(f"  Allowed extensions: {native_manifest.get('allowed_extensions', [])}")
        
        # Check if the launcher path in the manifest is correct
        launcher_path = native_manifest.get('path', '')
        if os.path.exists(launcher_path):
            print(f"OK: Launcher path in manifest is valid: {launcher_path}")
        else:
            print(f"ERROR: Launcher path in manifest is invalid: {launcher_path}")
            return False
        
    except Exception as e:
        print(f"ERROR: Error reading native manifest: {e}")
        return False
    
    return True

def test_job_queue_workflow():
    """Test the complete job queue workflow."""
    print("\n=== Testing Job Queue Workflow ===")
    
    queue_dir = test_queue_directory()
    if not queue_dir:
        print("ERROR: Cannot test job queue without valid queue directory")
        return False
    
    jobs_dir = os.path.join(queue_dir, "jobs")
    results_dir = os.path.join(queue_dir, "results")
    
    # Create a test job
    test_job = {
        "id": "test-job-123",
        "type": "sendEmail",
        "payload": {
            "to": ["test@example.com"],
            "subject": "Test Email",
            "bodyHtml": "<p>This is a test email</p>",
            "attachments": []
        }
    }
    
    job_file = os.path.join(jobs_dir, "test-job-123.json")
    try:
        with open(job_file, 'w') as f:
            json.dump(test_job, f, indent=2)
        print(f"OK: Test job created: {job_file}")
    except Exception as e:
        print(f"ERROR: Failed to create test job: {e}")
        return False
    
    # Wait a bit and check if job is processed
    print("Waiting for job processing (5 seconds)...")
    time.sleep(5)
    
    # Check if job file still exists
    if os.path.exists(job_file):
        print("WARNING: Job file still exists, may not have been processed")
    else:
        print("OK: Job file was processed (moved or deleted)")
    
    # Check for result file
    result_file = os.path.join(results_dir, "test-job-123.json")
    if os.path.exists(result_file):
        try:
            with open(result_file, 'r') as f:
                result = json.load(f)
            print(f"OK: Result file found: {result}")
        except Exception as e:
            print(f"ERROR: Error reading result file: {e}")
    else:
        print("WARNING: No result file found (expected if native host is not running)")
    
    # Clean up
    try:
        if os.path.exists(job_file):
            os.remove(job_file)
        if os.path.exists(result_file):
            os.remove(result_file)
    except:
        pass
    
    return True

def main():
    """Run all diagnostic tests."""
    print("Thunderbird Integration Diagnostic Tool")
    print("=" * 50)
    
    results = {
        "queue_directory": test_queue_directory(),
        "native_host_launcher": test_native_host_launcher(),
        "native_host_process": test_native_host_process(),
        "thunderbird_extension": test_thunderbird_extension(),
        "job_queue_workflow": test_job_queue_workflow()
    }
    
    print("\n" + "=" * 50)
    print("DIAGNOSTIC SUMMARY")
    print("=" * 50)
    
    for test_name, result in results.items():
        status = "PASS" if result else "FAIL"
        print(f"{test_name.replace('_', ' ').title()}: {status}")
    
    print("\nRECOMMENDATIONS:")
    print("=" * 50)
    
    if not results["queue_directory"]:
        print("- Fix queue directory permissions or create a valid queue directory")
    
    if not results["native_host_launcher"]:
        print("- Check the native host launcher script and Python executable path")
    
    if not results["native_host_process"]:
        print("- Check if the native host process can be started manually")
    
    if not results["thunderbird_extension"]:
        print("- Verify Thunderbird extension configuration and native manifest registration")
    
    if not results["job_queue_workflow"]:
        print("- Ensure the native host process is running and monitoring the queue directory")
    
    print("\nFor detailed Thunderbird extension installation instructions, see:")
    print("- thunderbird_extension/README.md")

if __name__ == "__main__":
    main()