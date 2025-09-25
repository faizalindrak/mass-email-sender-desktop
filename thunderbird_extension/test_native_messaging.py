#!/usr/bin/env python3
"""
Simple test script to verify native messaging protocol.
Run this script manually to test stdin/stdout communication.
"""
import sys
import json
import struct
import os

# Windows binary mode setup
if os.name == "nt":
    try:
        import msvcrt
        if hasattr(sys.stdin, 'fileno') and sys.stdin.fileno() >= 0:
            msvcrt.setmode(sys.stdin.fileno(), os.O_BINARY)
        if hasattr(sys.stdout, 'fileno') and sys.stdout.fileno() >= 0:
            msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)
    except (OSError, AttributeError, ImportError):
        pass

def log(msg):
    """Log to stderr so it doesn't interfere with native messaging."""
    sys.stderr.write(f"[TEST] {msg}\n")
    sys.stderr.flush()

def read_message():
    """Read a native messaging message from stdin."""
    try:
        # Read 4-byte length prefix
        raw_length = sys.stdin.buffer.read(4)
        if not raw_length or len(raw_length) < 4:
            return None
        
        length = struct.unpack('<I', raw_length)[0]
        log(f"Message length: {length}")
        
        if length == 0:
            return None
        if length > 1024 * 1024:  # 1MB limit
            log(f"Message too large: {length}")
            return None
        
        # Read message data
        data = sys.stdin.buffer.read(length)
        if not data or len(data) != length:
            log(f"Incomplete read: expected {length}, got {len(data) if data else 0}")
            return None
        
        # Parse JSON
        message = json.loads(data.decode('utf-8'))
        log(f"Received message: {message}")
        return message
    except Exception as e:
        log(f"Read error: {e}")
        return None

def send_message(msg):
    """Send a native messaging message to stdout."""
    try:
        encoded = json.dumps(msg, ensure_ascii=False).encode('utf-8')
        length = len(encoded)
        log(f"Sending message length: {length}")
        log(f"Sending message: {msg}")
        
        # Write length prefix
        sys.stdout.buffer.write(struct.pack('<I', length))
        # Write message data
        sys.stdout.buffer.write(encoded)
        # Flush
        sys.stdout.buffer.flush()
        
        log("Message sent successfully")
        return True
    except Exception as e:
        log(f"Send error: {e}")
        return False

def main():
    log("Native messaging test starting...")
    
    # Send a hello message
    send_message({
        "type": "hello",
        "message": "Test native messaging host",
        "timestamp": 12345
    })
    
    # Listen for incoming messages
    message_count = 0
    while message_count < 5:  # Only process a few messages for testing
        msg = read_message()
        if msg is None:
            log("No message received, continuing...")
            continue
        
        message_count += 1
        log(f"Processing message {message_count}: {msg}")
        
        # Send a response
        response = {
            "type": "response",
            "original_message": msg,
            "message_count": message_count,
            "status": "ok"
        }
        
        if not send_message(response):
            log("Failed to send response")
            break
    
    log("Test complete")

if __name__ == "__main__":
    main()