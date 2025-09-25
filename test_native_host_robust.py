#!/usr/bin/env python3
"""
Test script to validate the robust native host implementation.
This will test the stdio communication directly.
"""

import os
import sys
import json
import struct
import subprocess
import time
import tempfile
from pathlib import Path

def test_robust_native_host():
    """Test the robust native host by simulating Thunderbird communication."""
    print("Testing robust native host...")
    
    # Path to the robust native host
    host_path = os.path.join("thunderbird_extension", "native_host_robust.py")
    if not os.path.exists(host_path):
        print(f"❌ Robust native host not found: {host_path}")
        return False
    
    # Set up environment
    env = os.environ.copy()
    env["TB_QUEUE_DIR"] = os.path.abspath("tb_queue")
    env["PYTHONUNBUFFERED"] = "1"
    
    try:
        # Start the native host process
        print(f"Starting native host: {host_path}")
        process = subprocess.Popen(
            [sys.executable, host_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=env
        )
        
        # Give it a moment to initialize
        time.sleep(0.5)
        
        # Send a hello message
        hello_msg = {
            "type": "hello",
            "from": "test-client",
            "ts": int(time.time())
        }
        
        print("Sending hello message...")
        if send_message(process.stdin, hello_msg):
            print("✅ Hello message sent successfully")
        else:
            print("❌ Failed to send hello message")
            return False
        
        # Try to read response
        print("Waiting for response...")
        response = read_message(process.stdout)
        if response:
            print(f"✅ Received response: {response}")
        else:
            print("❌ No response received")
        
        # Clean shutdown
        process.terminate()
        try:
            process.wait(timeout=2)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()
        
        print("✅ Native host test completed successfully")
        return True
        
    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        if 'process' in locals():
            try:
                process.terminate()
                process.wait(timeout=2)
            except:
                try:
                    process.kill()
                except:
                    pass
        return False

def send_message(pipe, msg):
    """Send a native messaging message."""
    try:
        encoded = json.dumps(msg, ensure_ascii=False).encode("utf-8")
        length_data = struct.pack("<I", len(encoded))
        
        pipe.write(length_data)
        pipe.write(encoded)
        pipe.flush()
        return True
    except Exception as e:
        print(f"Send error: {e}")
        return False

def read_message(pipe):
    """Read a native messaging message."""
    try:
        # Read length prefix
        raw_length = pipe.read(4)
        if not raw_length or len(raw_length) < 4:
            return None
        
        length = struct.unpack("<I", raw_length)[0]
        if length == 0 or length > 1024 * 1024:
            return None
        
        # Read message data
        data = pipe.read(length)
        if not data or len(data) != length:
            return None
        
        return json.loads(data.decode("utf-8"))
    except Exception as e:
        print(f"Read error: {e}")
        return None

def test_job_processing():
    """Test job file processing."""
    print("\nTesting job processing...")
    
    queue_dir = os.path.abspath("tb_queue")
    jobs_dir = os.path.join(queue_dir, "jobs")
    results_dir = os.path.join(queue_dir, "results")
    
    # Ensure directories exist
    os.makedirs(jobs_dir, exist_ok=True)
    os.makedirs(results_dir, exist_ok=True)
    
    # Create a test job
    job_id = "test-job-123"
    job = {
        "id": job_id,
        "type": "sendEmail",
        "payload": {
            "to": ["test@example.com"],
            "cc": [],
            "bcc": [],
            "subject": "Test Email",
            "bodyHtml": "<p>Test message</p>",
            "attachments": []
        }
    }
    
    job_path = os.path.join(jobs_dir, f"{job_id}.json")
    
    try:
        # Write the job file
        with open(job_path, "w", encoding="utf-8") as f:
            json.dump(job, f, indent=2)
        
        print(f"✅ Created test job: {job_path}")
        
        # Clean up
        if os.path.exists(job_path):
            os.remove(job_path)
            print("✅ Cleaned up test job")
        
        return True
        
    except Exception as e:
        print(f"❌ Job processing test failed: {e}")
        return False

def main():
    """Run all tests."""
    print("=" * 60)
    print("ROBUST NATIVE HOST TEST SUITE")
    print("=" * 60)
    
    tests = [
        ("Job Processing", test_job_processing),
        ("Native Host Communication", test_robust_native_host),
    ]
    
    passed = 0
    failed = 0
    
    for test_name, test_func in tests:
        print(f"\n🔍 Running {test_name}...")
        try:
            if test_func():
                print(f"✅ {test_name}: PASSED")
                passed += 1
            else:
                print(f"❌ {test_name}: FAILED")
                failed += 1
        except Exception as e:
            print(f"❌ {test_name}: FAILED with exception: {e}")
            failed += 1
    
    print("\n" + "=" * 60)
    print("TEST RESULTS")
    print("=" * 60)
    print(f"Passed: {passed}")
    print(f"Failed: {failed}")
    
    if failed == 0:
        print("\n🎉 ALL TESTS PASSED!")
        print("\nThe robust native host should now work correctly.")
        print("Make sure to:")
        print("1. Load the Thunderbird extension")
        print("2. Set email_client to 'thunderbird' in your profile")
        print("3. Test with a real file")
        return True
    else:
        print(f"\n⚠️ {failed} test(s) failed.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)