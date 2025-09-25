#!/usr/bin/env python3
"""
Simple diagnostic script to test Thunderbird integration.
"""

import os
import sys
import json
import time
import subprocess

def test_queue_directory():
    """Test if the queue directory is accessible and writable."""
    print("=== Testing Queue Directory ===")
    
    queue_dir = "tb_queue"
    print(f"Testing queue directory: {queue_dir}")
    
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
        
        print(f"OK: Queue directory {queue_dir} is accessible and writable")
        return queue_dir
        
    except Exception as e:
        print(f"ERROR: Queue directory {queue_dir} failed: {e}")
        return None

def test_permissions():
    """Test file permissions in detail."""
    print("\n=== Testing File Permissions ===")
    
    # Check current user
    try:
        import getpass
        username = getpass.getuser()
        print(f"Current user: {username}")
    except:
        print("Could not get current username")
    
    # Check if we can access the tb_queue directory
    queue_dir = "tb_queue"
    if os.path.exists(queue_dir):
        print(f"Queue directory exists: {queue_dir}")
        
        # Check directory permissions
        try:
            # Try to list files
            files = os.listdir(queue_dir)
            print(f"Files in queue directory: {files}")
            
            # Check if we can create a file
            test_file = os.path.join(queue_dir, "permission_test.txt")
            with open(test_file, 'w') as f:
                f.write("permission test")
            os.remove(test_file)
            print("OK: Can create and delete files in queue directory")
            
        except Exception as e:
            print(f"ERROR: Cannot access queue directory properly: {e}")
    else:
        print(f"Queue directory does not exist: {queue_dir}")

def main():
    """Run diagnostic tests."""
    print("Thunderbird Integration Diagnostic Tool (Simple)")
    print("=" * 50)
    
    test_queue_directory()
    test_permissions()
    
    print("\n" + "=" * 50)
    print("MANUAL CHECKS NEEDED:")
    print("=" * 50)
    print("1. Check if Thunderbird is running")
    print("2. Check if the Email Automation extension is installed in Thunderbird")
    print("3. Check if the native manifest is registered in Thunderbird")
    print("4. Check Thunderbird's error console for extension errors")

if __name__ == "__main__":
    main()