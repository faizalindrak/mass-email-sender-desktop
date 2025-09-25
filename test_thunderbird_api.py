#!/usr/bin/env python3
"""
Test script to verify Thunderbird API fixes.
"""

import os
import sys
import json
import time
import uuid
from src.core.email_sender import ThunderbirdExtensionSender

def test_thunderbird_with_attachment():
    """Test the ThunderbirdExtensionSender with an attachment."""
    print("=== Testing Thunderbird Email Sender with Attachment ===")
    
    try:
        # Create sender
        sender = ThunderbirdExtensionSender()
        print("ThunderbirdExtensionSender created successfully")
        
        # Test email data with attachment
        test_email = {
            "to": ["test@example.com"],
            "subject": "Test Email with Attachment",
            "body": "<h1>Test Email</h1><p>This is a test email sent via Thunderbird extension with an attachment.</p>",
            "attachment_path": "test-folder/ORDERSHEET TT003 TTI2.pdf"  # Use the actual attachment file
        }
        
        # Check if attachment file exists
        if os.path.exists(test_email["attachment_path"]):
            print(f"Attachment file found: {test_email['attachment_path']}")
        else:
            print(f"Attachment file not found: {test_email['attachment_path']}")
            # Create a simple test file
            os.makedirs("test-folder", exist_ok=True)
            with open(test_email["attachment_path"], "w") as f:
                f.write("This is a test attachment file.")
            print(f"Created test attachment file: {test_email['attachment_path']}")
        
        print("Sending test email with attachment...")
        print(f"To: {test_email['to']}")
        print(f"Subject: {test_email['subject']}")
        print(f"Body length: {len(test_email['body'])} characters")
        print(f"Attachment: {test_email['attachment_path']}")
        
        # Send email
        result = sender.send_email(
            to_emails=test_email["to"],
            subject=test_email["subject"],
            body=test_email["body"],
            attachment_path=test_email["attachment_path"]
        )
        
        print(f"Email send result: {result}")
        
        if result:
            print("SUCCESS: Email with attachment sent successfully!")
        else:
            print("FAILED: Email with attachment could not be sent")
            
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
    print("Thunderbird API Fix Verification")
    print("=" * 50)
    
    # Check queue status first
    check_queue_status()
    
    # Check native host log
    check_native_host_log()
    
    # Test email sending with attachment
    print("\n" + "=" * 50)
    result = test_thunderbird_with_attachment()
    
    # Check queue status again
    print("\n" + "=" * 50)
    check_queue_status()
    
    # Final status
    print("\n" + "=" * 50)
    print("TEST SUMMARY")
    print("=" * 50)
    
    if result:
        print("STATUS: SUCCESS - Thunderbird API fixes appear to be working")
    else:
        print("STATUS: FAILED - Thunderbird API needs further investigation")
    
    print("\nNext steps:")
    print("1. Check Thunderbird's error console for extension errors")
    print("2. Verify the Email Automation extension is installed and enabled")
    print("3. Check Thunderbird's extension settings for native messaging permissions")
    print("4. Review the native host log for detailed error information")

if __name__ == "__main__":
    main()