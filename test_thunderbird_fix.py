#!/usr/bin/env python3
"""
Test script to verify Thunderbird integration fixes.
"""

import os
import sys
import json
import time
import uuid
from src.core.email_sender import ThunderbirdExtensionSender

def test_thunderbird_sender():
    """Test the ThunderbirdExtensionSender with a simple email."""
    print("=== Testing Thunderbird Email Sender ===")
    
    try:
        # Create sender
        sender = ThunderbirdExtensionSender()
        print("ThunderbirdExtensionSender created successfully")
        
        # Test email data
        test_email = {
            "to": ["test@example.com"],
            "subject": "Test Email from Thunderbird Extension",
            "body": "<h1>Test Email</h1><p>This is a test email sent via Thunderbird extension.</p>",
            "attachment_path": None  # No attachment for this test
        }
        
        print("Sending test email...")
        print(f"To: {test_email['to']}")
        print(f"Subject: {test_email['subject']}")
        print(f"Body length: {len(test_email['body'])} characters")
        
        # Send email
        result = sender.send_email(
            to_emails=test_email["to"],
            subject=test_email["subject"],
            body=test_email["body"]
        )
        
        print(f"Email send result: {result}")
        
        if result:
            print("SUCCESS: Email sent successfully!")
        else:
            print("FAILED: Email could not be sent")
            
        return result
        
    except Exception as e:
        print(f"ERROR: Exception during test: {e}")
        import traceback
        traceback.print_exc()
        return False

def check_queue_status():
    """Check the status of the queue directory."""
    print("\n=== Checking Queue Status ===")
    
    queue_dir = "tb_queue"
    jobs_dir = os.path.join(queue_dir, "jobs")
    results_dir = os.path.join(queue_dir, "results")
    
    print(f"Queue directory: {queue_dir}")
    print(f"Jobs directory: {jobs_dir}")
    print(f"Results directory: {results_dir}")
    
    # Check jobs
    if os.path.exists(jobs_dir):
        jobs = os.listdir(jobs_dir)
        print(f"Jobs in queue: {len(jobs)}")
        for job in jobs:
            print(f"  - {job}")
    else:
        print("Jobs directory does not exist")
    
    # Check results
    if os.path.exists(results_dir):
        results = os.listdir(results_dir)
        print(f"Results in queue: {len(results)}")
        for result in results:
            result_path = os.path.join(results_dir, result)
            try:
                with open(result_path, 'r') as f:
                    result_data = json.load(f)
                print(f"  - {result}: {result_data.get('success', 'unknown')}")
                if 'error' in result_data:
                    print(f"    Error: {result_data['error']}")
            except Exception as e:
                print(f"  - {result}: Could not read result file ({e})")
    else:
        print("Results directory does not exist")

def check_native_host_log():
    """Check the native host log for recent activity."""
    print("\n=== Checking Native Host Log ===")
    
    log_path = "tb_queue/native_host.log"
    if os.path.exists(log_path):
        try:
            with open(log_path, 'r') as f:
                lines = f.readlines()
            
            # Show last 20 lines
            recent_lines = lines[-20:] if len(lines) > 20 else lines
            print(f"Last {len(recent_lines)} lines from native host log:")
            
            for line in recent_lines:
                print(f"  {line.strip()}")
                
        except Exception as e:
            print(f"Could not read native host log: {e}")
    else:
        print("Native host log does not exist")

def main():
    """Run the test."""
    print("Thunderbird Integration Fix Verification")
    print("=" * 50)
    
    # Check queue status first
    check_queue_status()
    
    # Check native host log
    check_native_host_log()
    
    # Test email sending
    print("\n" + "=" * 50)
    result = test_thunderbird_sender()
    
    # Check queue status again
    print("\n" + "=" * 50)
    check_queue_status()
    
    # Final status
    print("\n" + "=" * 50)
    print("TEST SUMMARY")
    print("=" * 50)
    
    if result:
        print("STATUS: SUCCESS - Thunderbird integration appears to be working")
    else:
        print("STATUS: FAILED - Thunderbird integration needs further investigation")
    
    print("\nNext steps:")
    print("1. Check Thunderbird's error console for extension errors")
    print("2. Verify the Email Automation extension is installed and enabled")
    print("3. Check Thunderbird's extension settings for native messaging permissions")
    print("4. Review the native host log for detailed error information")

if __name__ == "__main__":
    main()