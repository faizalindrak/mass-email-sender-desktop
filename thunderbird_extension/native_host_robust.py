#!/usr/bin/env python3
"""
Robust native messaging host for Windows with proper stdio handling.
This version bypasses problematic Windows stdio issues by using low-level I/O.
"""

import sys
import os
import json
import struct
import threading
import time
import traceback
import subprocess
from typing import Dict, Any, Optional

# Ensure we're in the correct directory
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

# Import the original NativeHost as base
try:
    from native_host import Logger, default_queue_dir
except ImportError:
    # Fallback logger if import fails
    class Logger:
        def __init__(self, queue_dir: str):
            self._log_path = os.path.join(queue_dir, "native_host.log")
        
        def log(self, *args):
            msg = "[TB-NativeHost-Robust] " + " ".join(str(a) for a in args)
            try:
                with open(self._log_path, "a", encoding="utf-8") as f:
                    f.write(time.strftime("%Y-%m-%d %H:%M:%S") + " " + msg + "\n")
            except:
                pass
    
    def default_queue_dir():
        return os.path.abspath("tb_queue")

HOST_NAME = "com.emailautomation.tbhost"
DEFAULT_POLL_INTERVAL = 0.5
DEFAULT_TIMEOUT_SECONDS = 120.0

class RobustNativeHost:
    """Robust native messaging host with better Windows stdio handling."""
    
    def __init__(self, queue_dir: Optional[str] = None, poll_interval: float = DEFAULT_POLL_INTERVAL):
        self.queue_dir = os.path.abspath(queue_dir or os.getenv("TB_QUEUE_DIR") or default_queue_dir())
        self.jobs_dir = os.path.join(self.queue_dir, "jobs")
        self.processing_dir = os.path.join(self.queue_dir, "processing")
        self.results_dir = os.path.join(self.queue_dir, "results")
        
        # Ensure directories exist
        for d in (self.jobs_dir, self.processing_dir, self.results_dir):
            os.makedirs(d, exist_ok=True)

        self.log = Logger(self.queue_dir).log
        self._stop = threading.Event()
        self._stdin_thread = None
        self._incoming_lock = threading.Lock()
        self._incoming_queue = []
        self.poll_interval = float(poll_interval)
        self._connected = False
        self._last_send_error = None
        self._inflight = {}
        
        # Setup robust stdio handling
        self._setup_robust_stdio()
    
    def _setup_robust_stdio(self):
        """Setup robust stdio handling for Windows native messaging."""
        try:
            self.log("Setting up robust stdio for native messaging...")
            
            # For Windows, ensure we have proper binary streams
            if os.name == "nt":
                import msvcrt
                
                # Get the raw file descriptors
                stdin_fd = sys.stdin.fileno()
                stdout_fd = sys.stdout.fileno()
                
                self.log(f"Original stdin fd: {stdin_fd}, stdout fd: {stdout_fd}")
                
                # Set binary mode on the file descriptors
                try:
                    msvcrt.setmode(stdin_fd, os.O_BINARY)
                    msvcrt.setmode(stdout_fd, os.O_BINARY)
                    self.log("Successfully set binary mode on stdio")
                except Exception as e:
                    self.log(f"Failed to set binary mode: {e}")
                
                # Create our own binary streams using the file descriptors
                try:
                    self._stdin_raw = os.fdopen(stdin_fd, 'rb', buffering=0)
                    self._stdout_raw = os.fdopen(stdout_fd, 'wb', buffering=0)
                    self.log("Created raw binary streams")
                except Exception as e:
                    self.log(f"Failed to create raw streams: {e}")
                    # Fallback to system streams
                    self._stdin_raw = sys.stdin.buffer if hasattr(sys.stdin, 'buffer') else sys.stdin
                    self._stdout_raw = sys.stdout.buffer if hasattr(sys.stdout, 'buffer') else sys.stdout
            else:
                # Non-Windows platforms
                self._stdin_raw = sys.stdin.buffer if hasattr(sys.stdin, 'buffer') else sys.stdin
                self._stdout_raw = sys.stdout.buffer if hasattr(sys.stdout, 'buffer') else sys.stdout
                
        except Exception as e:
            self.log(f"Error setting up stdio: {e}")
            # Final fallback
            self._stdin_raw = sys.stdin
            self._stdout_raw = sys.stdout
    
    def read_message(self) -> Optional[dict]:
        """Read a native messaging message with robust error handling."""
        try:
            # Read the 4-byte length prefix
            raw_length = self._stdin_raw.read(4)
            if not raw_length or len(raw_length) < 4:
                return None
            
            message_length = struct.unpack("<I", raw_length)[0]
            if message_length == 0 or message_length > 1024 * 1024:  # Max 1MB
                self.log(f"Invalid message length: {message_length}")
                return None
            
            # Read the message data
            data = self._stdin_raw.read(message_length)
            if not data or len(data) != message_length:
                self.log(f"Incomplete message read: expected {message_length}, got {len(data) if data else 0}")
                return None
            
            # Parse JSON
            try:
                message = json.loads(data.decode("utf-8"))
                if not isinstance(message, dict):
                    self.log("Message is not a JSON object")
                    return None
                return message
            except json.JSONDecodeError as je:
                self.log(f"JSON decode error: {repr(je)}")
                return None
                
        except Exception as e:
            self.log("read_message error:", repr(e))
            return None
    
    def send_message(self, msg: dict) -> bool:
        """Send a native messaging message with robust error handling."""
        try:
            encoded = json.dumps(msg, ensure_ascii=False).encode("utf-8")
            length_data = struct.pack("<I", len(encoded))
            
            self.log(f"Sending message (length: {len(encoded)})")
            
            # Write directly to the raw stdout stream
            self._stdout_raw.write(length_data)
            self._stdout_raw.write(encoded)
            
            # Ensure data is flushed
            if hasattr(self._stdout_raw, 'flush'):
                self._stdout_raw.flush()
            
            self.log("Message sent successfully")
            self._last_send_error = None
            return True
            
        except Exception as e:
            err = repr(e)
            self._last_send_error = err
            self.log("send_message error:", err)
            self._connected = False
            return False
    
    def stdin_reader(self):
        """Continuously read messages from Thunderbird extension."""
        consecutive_failures = 0
        while not self._stop.is_set():
            try:
                msg = self.read_message()
                if msg is None:
                    consecutive_failures += 1
                    if consecutive_failures > 100:
                        self.log("Too many consecutive read failures, backing off...")
                        time.sleep(1.0)
                        consecutive_failures = 0
                    else:
                        time.sleep(0.05)
                    continue
                
                consecutive_failures = 0
                with self._incoming_lock:
                    self._incoming_queue.append(msg)
                    
            except Exception as e:
                self.log("stdin_reader exception:", repr(e))
                consecutive_failures += 1
                time.sleep(0.1)
    
    def handle_incoming(self, msg: dict):
        """Handle incoming messages from the extension."""
        try:
            # Mark as connected upon receiving any message from the extension
            self._connected = True
            
            # Handle response messages
            if "id" in msg and ("success" in msg or "error" in msg):
                job_id = str(msg.get("id"))
                success = bool(msg.get("success", False))
                error = msg.get("error")
                processing_path = self._inflight.pop(job_id, None)

                # Write result file
                result_path = os.path.join(self.results_dir, f"{job_id}.json")
                result = {"id": job_id, "success": success}
                if error:
                    result["error"] = str(error)
                
                try:
                    with open(result_path, "w", encoding="utf-8") as f:
                        json.dump(result, f, ensure_ascii=False, indent=2)
                    self.log("Wrote result:", result_path)
                except Exception as e:
                    self.log("Failed to write result file:", repr(e))

                # Clean processing file
                try:
                    if processing_path and os.path.exists(processing_path):
                        os.remove(processing_path)
                except Exception:
                    pass
                return

            # Handle hello/ping messages
            if msg.get("type") == "hello":
                self._connected = True
                self.send_message({"type": "hello_ack", "ts": int(time.time())})
                return

            # Log other messages
            self.log("Unhandled incoming message:", msg)
        except Exception:
            self.log("handle_incoming exception:", traceback.format_exc())
    
    def scan_and_forward_jobs(self):
        """Scan for jobs and forward them to the extension."""
        try:
            # Only forward jobs if connected
            if not self._connected:
                return
            
            files = []
            try:
                files = [f for f in os.listdir(self.jobs_dir) if f.lower().endswith(".json")]
            except Exception:
                pass

            for fname in files:
                job_path = os.path.join(self.jobs_dir, fname)
                try:
                    with open(job_path, "r", encoding="utf-8") as f:
                        job = json.load(f)
                except Exception as e:
                    self.log("Failed to read job file:", job_path, repr(e))
                    continue

                job_id = str(job.get("id") or "")
                if not job_id:
                    self.log("Skipping job without id:", job_path)
                    continue

                # Check if result already exists
                result_path = os.path.join(self.results_dir, f"{job_id}.json")
                if os.path.exists(result_path):
                    try:
                        os.remove(job_path)
                    except Exception:
                        pass
                    continue

                # Skip if already in flight
                if job_id in self._inflight:
                    continue

                # Try to move to processing
                processing_path = os.path.join(self.processing_dir, f"{job_id}.json")
                moved = False
                
                try:
                    if os.path.exists(processing_path):
                        os.remove(processing_path)
                except Exception:
                    pass
                
                # Attempt to move with retries
                for attempt in range(5):
                    try:
                        os.replace(job_path, processing_path)
                        moved = True
                        break
                    except (PermissionError, OSError) as pe:
                        if attempt < 4:
                            time.sleep(0.1 * (attempt + 1))
                            continue
                        self.log(f"Failed to move job after {attempt + 1} attempts: {repr(pe)}")
                        break

                if not moved:
                    # Process in place
                    processing_path = job_path
                    self.log("Processing job in-place")

                # Prepare message
                job_type = job.get("type") or "sendEmail"
                msg = dict(job)
                msg["type"] = job_type

                # Send to extension
                ok = self.send_message(msg)
                if not ok:
                    # Write failure result
                    err = self._last_send_error or "native_host_send_error"
                    try:
                        with open(result_path, "w", encoding="utf-8") as f:
                            json.dump({"id": job_id, "success": False, "error": f"Send failed: {err}"}, f)
                        self.log("Wrote failure result:", result_path)
                    except Exception as we:
                        self.log("Failed to write failure result:", repr(we))
                    
                    # Cleanup
                    try:
                        if processing_path and os.path.exists(processing_path):
                            os.remove(processing_path)
                    except Exception:
                        pass
                    continue

                self._inflight[job_id] = processing_path
                self.log("Forwarded job to extension:", job_id)
                
        except Exception:
            self.log("scan_and_forward_jobs exception:", traceback.format_exc())
    
    def run(self):
        """Main execution loop."""
        self.log("Robust Native Host starting. Queue:", self.queue_dir)
        
        # Start stdin reader thread
        self._stdin_thread = threading.Thread(target=self.stdin_reader, daemon=True)
        self._stdin_thread.start()
        
        try:
            while not self._stop.is_set():
                # Handle incoming messages
                while True:
                    msg = None
                    with self._incoming_lock:
                        if self._incoming_queue:
                            msg = self._incoming_queue.pop(0)
                    if msg is None:
                        break
                    self.handle_incoming(msg)

                # Scan and forward jobs
                self.scan_and_forward_jobs()

                time.sleep(self.poll_interval)
                
        except KeyboardInterrupt:
            pass
        except Exception:
            self.log("Main loop exception:", traceback.format_exc())
        finally:
            self._stop.set()
            self.log("Robust Native Host stopping.")

def main():
    """Main entry point."""
    try:
        queue_dir = None
        if len(sys.argv) > 1:
            queue_dir = sys.argv[1]
        else:
            queue_dir = os.getenv("TB_QUEUE_DIR")
        
        host = RobustNativeHost(queue_dir=queue_dir)
        host.run()
        
    except KeyboardInterrupt:
        sys.stderr.write("Native host interrupted by user\n")
    except Exception as e:
        sys.stderr.write(f"Native host fatal error: {repr(e)}\n")
        sys.stderr.write(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()