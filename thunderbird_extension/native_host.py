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
import atexit
import signal
from typing import Dict, Any, Optional

# Ensure we're in the correct directory
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

HOST_NAME = "com.emailautomation.tbhost"
DEFAULT_POLL_INTERVAL = 0.5
DEFAULT_TIMEOUT_SECONDS = 120.0

def default_queue_dir() -> str:
    try:
        if os.name == "nt":
            appdata = os.getenv("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
            return os.path.join(appdata, "EmailAutomation", "tb_queue")
        else:
            if sys.platform == "darwin":
                base = os.path.join(os.path.expanduser("~"), "Library", "Application Support", "EmailAutomation")
            else:
                base = os.path.join(os.path.expanduser("~"), ".local", "share", "email_automation")
            return os.path.join(base, "tb_queue")
    except Exception:
        return os.path.abspath(os.path.join("tb_queue"))

class Logger:
    def __init__(self, queue_dir: str):
        self._lock = threading.Lock()
        self._log_path = os.path.join(queue_dir, "native_host.log")
        try:
            os.makedirs(queue_dir, exist_ok=True)
        except Exception:
            pass

    def log(self, *args):
        msg = "[TB-NativeHost-Robust] " + " ".join(str(a) for a in args)
        try:
            with self._lock:
                with open(self._log_path, "a", encoding="utf-8") as f:
                    f.write(time.strftime("%Y-%m-%d %H:%M:%S") + " " + msg + "\n")
        except Exception:
            # Fallback to stderr
            try:
                sys.stderr.write(msg + "\n")
                sys.stderr.flush()
            except Exception:
                pass

class NativeHost:
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
        self._lock_file = None
        
        # Try to acquire process lock to prevent multiple instances
        if not self._acquire_process_lock():
            self.log("Another native host instance is already running, exiting...")
            sys.exit(0)
        else:
            self.log("Process lock acquired successfully")
        
        # Register cleanup on exit
        atexit.register(self._cleanup)
        try:
            signal.signal(signal.SIGTERM, self._signal_handler)
            signal.signal(signal.SIGINT, self._signal_handler)
        except ValueError:
            # Signals not available in all contexts
            pass
        
        # Setup robust stdio handling
        self._setup_robust_stdio()
    
    def _acquire_process_lock(self) -> bool:
        """Acquire a file lock to ensure only one instance runs."""
        try:
            lock_path = os.path.join(self.queue_dir, "native_host.lock")
            self.log(f"Attempting to acquire lock at: {lock_path}")
            
            # Check if we can write to the queue directory
            try:
                test_file = os.path.join(self.queue_dir, "test_write.tmp")
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                self.log("Queue directory is writable")
            except Exception as e:
                self.log(f"ERROR: Cannot write to queue directory: {e}")
                return False
            
            # Clean up any stale lock file first
            if os.path.exists(lock_path):
                try:
                    # Try to read PID from lock file
                    with open(lock_path, 'r') as f:
                        old_pid = int(f.read().strip())
                    
                    self.log(f"Found existing lock file with PID: {old_pid}")
                    
                    # Check if process is still running
                    try:
                        if os.name == "nt":
                            # Windows: check if process exists
                            import subprocess
                            result = subprocess.run(['tasklist', '/FI', f'PID eq {old_pid}'],
                                                 capture_output=True, text=True, timeout=5)
                            if str(old_pid) not in result.stdout:
                                # Process not running, remove stale lock
                                os.remove(lock_path)
                                self.log(f"Removed stale lock file (PID {old_pid} not running)")
                            else:
                                self.log(f"Process {old_pid} is still running, cannot start")
                                return False
                        else:
                            # Unix: send signal 0 to check if process exists
                            os.kill(old_pid, 0)
                            self.log(f"Process {old_pid} is still running, cannot start")
                            return False
                    except (ProcessLookupError, OSError, subprocess.TimeoutExpired):
                        # Process doesn't exist, remove stale lock
                        try:
                            os.remove(lock_path)
                            self.log(f"Removed stale lock file (PID {old_pid} not found)")
                        except:
                            pass
                except (ValueError, OSError, FileNotFoundError) as e:
                    # Invalid lock file, remove it
                    self.log(f"Invalid lock file found, removing: {e}")
                    try:
                        os.remove(lock_path)
                        self.log("Removed invalid lock file")
                    except:
                        pass
            
            # Try to create new lock file
            try:
                self._lock_file = open(lock_path, 'w')
                self.log(f"Created lock file successfully")
            except Exception as e:
                self.log(f"ERROR: Failed to create lock file: {e}")
                return False
            
            if os.name == "nt":
                # Windows file locking
                try:
                    import msvcrt
                    msvcrt.locking(self._lock_file.fileno(), msvcrt.LK_NBLCK, 1)
                    self._lock_file.write(str(os.getpid()))
                    self._lock_file.flush()
                    self.log(f"Acquired process lock with file locking (PID: {os.getpid()})")
                    return True
                except (OSError, ImportError) as e:
                    self.log(f"File locking failed, using fallback: {e}")
                    # Fallback: just write PID and hope for the best
                    try:
                        self._lock_file.write(str(os.getpid()))
                        self._lock_file.flush()
                        self.log(f"Acquired process lock without file locking (PID: {os.getpid()})")
                        return True
                    except Exception as e2:
                        self.log(f"Fallback lock method also failed: {e2}")
                        return False
            else:
                # Unix file locking
                try:
                    import fcntl
                    fcntl.flock(self._lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    self._lock_file.write(str(os.getpid()))
                    self._lock_file.flush()
                    self.log(f"Acquired process lock (PID: {os.getpid()})")
                    return True
                except OSError as e:
                    self.log(f"Unix file locking failed: {e}")
                    return False
                    
        except Exception as e:
            self.log(f"Failed to acquire process lock: {e}")
            import traceback
            self.log(f"Lock acquisition traceback: {traceback.format_exc()}")
            return False
    
    def _release_process_lock(self):
        """Release the process lock."""
        try:
            if self._lock_file:
                self._lock_file.close()
                self._lock_file = None
            
            lock_path = os.path.join(self.queue_dir, "native_host.lock")
            if os.path.exists(lock_path):
                os.remove(lock_path)
                self.log("Released process lock")
        except Exception as e:
            self.log(f"Error releasing lock: {e}")
    
    def _cleanup(self):
        """Cleanup function called on exit."""
        self._stop.set()
        self._release_process_lock()
    
    def _signal_handler(self, signum, frame):
        """Handle signals for graceful shutdown."""
        self.log(f"Received signal {signum}, shutting down...")
        self._cleanup()
        sys.exit(0)
    
    def _setup_robust_stdio(self):
        """Setup robust stdio handling for Windows native messaging."""
        try:
            self.log("Setting up robust stdio for native messaging...")
            
            # For Windows, ensure we have proper binary streams
            if os.name == "nt":
                try:
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
                    
                    # Don't create new file descriptor streams - use direct os.read/os.write instead
                    # This avoids conflicts with Python's existing stdio streams
                    self._stdin_fd = stdin_fd
                    self._stdout_fd = stdout_fd
                    self._stdin_raw = None  # We'll use os.read directly
                    self._stdout_raw = None  # We'll use os.write directly
                    self.log("Will use direct os.read/os.write for stdio")
                except ImportError:
                    self.log("msvcrt not available, using fallback")
                    self._stdin_fd = sys.stdin.fileno()
                    self._stdout_fd = sys.stdout.fileno()
                    self._stdin_raw = None
                    self._stdout_raw = None
            else:
                # Non-Windows platforms - use direct file descriptors too
                self._stdin_fd = sys.stdin.fileno()
                self._stdout_fd = sys.stdout.fileno()
                self._stdin_raw = None
                self._stdout_raw = None
                
        except Exception as e:
            self.log(f"Error setting up stdio: {e}")
            # Final fallback
            self._stdin_fd = 0
            self._stdout_fd = 1
            self._stdin_raw = None
            self._stdout_raw = None
    
    def read_message(self) -> Optional[dict]:
        """Read a native messaging message using direct os.read."""
        try:
            # Read the 4-byte length prefix using os.read
            raw_length = os.read(self._stdin_fd, 4)
            if not raw_length or len(raw_length) < 4:
                return None
            
            message_length = struct.unpack("<I", raw_length)[0]
            if message_length == 0 or message_length > 1024 * 1024:  # Max 1MB
                self.log(f"Invalid message length: {message_length}")
                return None
            
            # Read the message data using os.read
            data = b""
            bytes_to_read = message_length
            while bytes_to_read > 0:
                chunk = os.read(self._stdin_fd, bytes_to_read)
                if not chunk:
                    break
                data += chunk
                bytes_to_read -= len(chunk)
            
            if len(data) != message_length:
                self.log(f"Incomplete message read: expected {message_length}, got {len(data)}")
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
                
        except (OSError, IOError) as e:
            # Handle specific os.read errors
            if e.errno == 9:  # Bad file descriptor
                self.log("stdin file descriptor closed")
                return None
            self.log("read_message os.read error:", repr(e))
            return None
        except Exception as e:
            self.log("read_message error:", repr(e))
            return None
    
    def send_message(self, msg: dict) -> bool:
        """Send a native messaging message using direct os.write."""
        try:
            encoded = json.dumps(msg, ensure_ascii=False).encode("utf-8")
            length_data = struct.pack("<I", len(encoded))
            
            self.log(f"Sending message (length: {len(encoded)})")
            
            # Use os.write directly - this has been working reliably
            os.write(self._stdout_fd, length_data)
            os.write(self._stdout_fd, encoded)
            
            self.log("Message sent via os.write")
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
            
            # Handle file data request from extension
            if msg.get("type") == "getFileData":
                file_path = msg.get("filePath", "")
                request_id = msg.get("id", "")
                
                if os.path.exists(file_path):
                    try:
                        import base64
                        # Read file as base64
                        with open(file_path, "rb") as f:
                            file_data = base64.b64encode(f.read()).decode('utf-8')
                        
                        # Send file data back to extension
                        self.send_message({
                            "type": "fileDataResponse",
                            "id": request_id,
                            "success": True,
                            "data": file_data,
                            "name": os.path.basename(file_path),
                            "size": os.path.getsize(file_path)
                        })
                        self.log(f"Sent file data for {file_path}")
                    except Exception as e:
                        self.log(f"Error reading file {file_path}: {e}")
                        self.send_message({
                            "type": "fileDataResponse",
                            "id": request_id,
                            "success": False,
                            "error": str(e)
                        })
                else:
                    self.log(f"File not found: {file_path}")
                    self.send_message({
                        "type": "fileDataResponse",
                        "id": request_id,
                        "success": False,
                        "error": "File not found"
                    })
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
                    # Process in place if we can't move
                    if os.path.exists(job_path):
                        processing_path = job_path
                        self.log("Processing job in-place")
                    else:
                        self.log("Job file disappeared, skipping")
                        continue

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
                        if processing_path != job_path and processing_path and os.path.exists(processing_path):
                            os.remove(processing_path)
                        elif processing_path == job_path and os.path.exists(job_path):
                            os.remove(job_path)
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
            self._cleanup()
            self.log("Robust Native Host stopping.")

def main():
    """Main entry point."""
    try:
        queue_dir = None
        if len(sys.argv) > 1:
            queue_dir = sys.argv[1]
        else:
            queue_dir = os.getenv("TB_QUEUE_DIR")
        
        host = NativeHost(queue_dir=queue_dir)
        host.run()
        
    except KeyboardInterrupt:
        sys.stderr.write("Native host interrupted by user\n")
    except Exception as e:
        sys.stderr.write(f"Native host fatal error: {repr(e)}\n")
        sys.stderr.write(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()