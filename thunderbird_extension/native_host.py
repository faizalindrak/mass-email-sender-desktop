#!/usr/bin/env python3
# Native Messaging host for Thunderbird MailExtension bridge.
# - Watches a filesystem queue for email jobs
# - Forwards jobs to the MailExtension via native messaging (stdio, length-prefixed)
# - Writes result files back to the queue on completion

import sys
import os
import json
import struct
import threading
import time
import traceback
from typing import Dict, Any, Optional

# Windows: ensure binary mode on stdio
if os.name == "nt":
    try:
        import msvcrt  # type: ignore
        msvcrt.setmode(sys.stdin.fileno(), os.O_BINARY)
        msvcrt.setmode(sys.stdout.fileno(), os.O_BINARY)
    except Exception:
        pass

HOST_NAME = "com.emailautomation.tbhost"
DEFAULT_POLL_INTERVAL = 0.5  # seconds
DEFAULT_TIMEOUT_SECONDS = 120.0  # for reference; enforced in app-side

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
        msg = "[TB-NativeHost] " + " ".join(str(a) for a in args)
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
    def __init__(self, queue_dir: Optional[str] = None, poll_interval: float = DEFAULT_POLL_INTERVAL):
        self.queue_dir = os.path.abspath(queue_dir or os.getenv("TB_QUEUE_DIR") or default_queue_dir())
        self.jobs_dir = os.path.join(self.queue_dir, "jobs")
        self.processing_dir = os.path.join(self.queue_dir, "processing")
        self.results_dir = os.path.join(self.queue_dir, "results")
        for d in (self.jobs_dir, self.processing_dir, self.results_dir):
            os.makedirs(d, exist_ok=True)

        self.log = Logger(self.queue_dir).log
        self._stop = threading.Event()
        self._stdin_thread = None
        self._incoming_lock = threading.Lock()
        self._incoming_queue = []  # type: list[dict]
        self.poll_interval = float(poll_interval)
        # Connection state with the Thunderbird extension (set True after first incoming message)
        self._connected = False
        # Last native messaging send error (for diagnostics/result writing)
        self._last_send_error = None
  
        # Track in-flight jobs: id -> processing_path
        self._inflight = {}  # type: Dict[str, str]

    # Native messaging helpers
    def read_message(self) -> Optional[dict]:
        try:
            raw_length = sys.stdin.buffer.read(4)
            if not raw_length or len(raw_length) < 4:
                return None
            message_length = struct.unpack("<I", raw_length)[0]
            if message_length == 0:
                return None
            data = sys.stdin.buffer.read(message_length)
            if not data:
                return None
            return json.loads(data.decode("utf-8"))
        except Exception as e:
            self.log("read_message error:", repr(e))
            return None

    def send_message(self, msg: dict) -> bool:
        """Send a message to the extension via native messaging. Returns True on success, False on failure.
        Stores last error into self._last_send_error for diagnostics."""
        try:
            encoded = json.dumps(msg, ensure_ascii=False).encode("utf-8")
            sys.stdout.buffer.write(struct.pack("<I", len(encoded)))
            sys.stdout.buffer.write(encoded)
            sys.stdout.buffer.flush()
            self._last_send_error = None
            return True
        except Exception as e:
            err = repr(e)
            self._last_send_error = err
            self.log("send_message error:", err)
            # mark disconnected so we stop forwarding until a fresh message arrives
            try:
                self._connected = False
            except Exception:
                pass
            return False

    def stdin_reader(self):
        # Continuously read messages from Thunderbird extension
        while not self._stop.is_set():
            msg = self.read_message()
            if msg is None:
                # End of stream or error
                time.sleep(0.05)
                continue
            with self._incoming_lock:
                self._incoming_queue.append(msg)

    def handle_incoming(self, msg: dict):
        try:
            # Mark as connected upon receiving any message from the extension
            self._connected = True
            # Expected responses: {"id": "...", "success": bool, "error": "...?"}
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

                # Optionally clean processing file
                try:
                    if processing_path and os.path.exists(processing_path):
                        os.remove(processing_path)
                except Exception:
                    pass
                return

            # Ping/pong handling
            if msg.get("type") == "hello":
                # Reply a simple ack and mark connected
                self._connected = True
                self.send_message({"type": "hello_ack", "ts": int(time.time())})
                return

            # Other messages can be logged
            self.log("Unhandled incoming message:", msg)
        except Exception:
            self.log("handle_incoming exception:", traceback.format_exc())

    def scan_and_forward_jobs(self):
        try:
            # Do not forward jobs until extension has connected via native messaging.
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

                # If result already exists, skip and remove stale job
                result_path = os.path.join(self.results_dir, f"{job_id}.json")
                if os.path.exists(result_path):
                    try:
                        os.remove(job_path)
                    except Exception:
                        pass
                    continue

                # If already inflight, skip
                if job_id in self._inflight:
                    continue

                # Move to processing (with retries to avoid transient file locks on Windows)
                processing_path = os.path.join(self.processing_dir, f"{job_id}.json")
                # Overwrite if exists (stale crash)
                try:
                    if os.path.exists(processing_path):
                        os.remove(processing_path)
                except Exception:
                    pass

                moved = False
                last_err = None
                for attempt in range(10):
                    try:
                        os.replace(job_path, processing_path)
                        moved = True
                        break
                    except PermissionError as pe:
                        last_err = pe
                        self.log(f"Move retry {attempt+1}/10 for {job_path} -> {processing_path} due to PermissionError; waiting 100ms")
                        time.sleep(0.1)
                    except Exception as e:
                        last_err = e
                        break

                if not moved:
                    # If we cannot move due to transient locks, process in-place and clean up after result
                    self.log("Failed to move job to processing, processing in-place:", repr(last_err))
                    processing_path = job_path

                # Normalize message for extension
                # Ensure type is "sendEmail"
                job_type = job.get("type") or "sendEmail"
                msg = dict(job)
                msg["type"] = job_type

                # Send to extension (handle send errors immediately with a failure result)
                ok = self.send_message(msg)
                if not ok:
                    err = getattr(self, "_last_send_error", "native_host_send_error")
                    result_path = os.path.join(self.results_dir, f"{job_id}.json")
                    try:
                        with open(result_path, "w", encoding="utf-8") as f:
                            json.dump({"id": job_id, "success": False, "error": err}, f, ensure_ascii=False, indent=2)
                        self.log("Wrote immediate failure result (send error):", result_path)
                    except Exception as we:
                        self.log("Failed to write immediate failure result:", repr(we))
                    # Clean processing file (or job file if processed in-place)
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
        self.log("Host starting. Queue:", self.queue_dir)
        # Start stdin reader thread
        self._stdin_thread = threading.Thread(target=self.stdin_reader, daemon=True)
        self._stdin_thread.start()

        # Defer any outbound messages until the extension connects to this host
 
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

                # Scan and forward any new jobs
                self.scan_and_forward_jobs()

                time.sleep(self.poll_interval)
        except KeyboardInterrupt:
            pass
        except Exception:
            self.log("Main loop exception:", traceback.format_exc())
        finally:
            self._stop.set()
            self.log("Host stopping.")

if __name__ == "__main__":
    qdir = os.getenv("TB_QUEUE_DIR") or None
    host = NativeHost(queue_dir=qdir)
    host.run()