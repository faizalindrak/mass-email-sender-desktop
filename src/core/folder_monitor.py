from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import re
import os
import shutil
import logging
from typing import Callable, Optional

class FileHandler(FileSystemEventHandler):
    def __init__(self, callback: Callable, key_pattern: str, file_extensions: list = None):
        self.callback = callback
        self.key_pattern = re.compile(key_pattern)
        self.file_extensions = file_extensions or ['.pdf', '.xlsx', '.docx', '.txt']
        self.logger = logging.getLogger(__name__)

    def on_created(self, event):
        if not event.is_directory:
            self.process_file(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            self.process_file(event.dest_path)

    def process_file(self, file_path: str):
        """Process new file and extract key"""
        try:
            filename = os.path.basename(file_path)
            file_ext = os.path.splitext(filename)[1].lower()

            # Check if file extension is allowed
            if self.file_extensions and file_ext not in self.file_extensions:
                self.logger.debug(f"Skipping file {filename} - extension {file_ext} not in allowed list")
                return

            # Extract key using regex pattern
            match = self.key_pattern.search(filename)
            if match:
                key = match.group(1) if match.groups() else match.group(0)
                self.logger.info(f"File detected: {filename}, extracted key: {key}")
                self.callback(file_path, key)
            else:
                self.logger.warning(f"No key found in filename: {filename}")

        except Exception as e:
            self.logger.error(f"Error processing file {file_path}: {str(e)}")

class FolderMonitor:
    def __init__(self):
        self.observer = None
        self.is_monitoring = False
        self.current_handler = None
        self.monitor_path = None
        self.logger = logging.getLogger(__name__)

    def start_monitoring(self, folder_path: str, callback: Callable, key_pattern: str, file_extensions: list = None):
        """Start monitoring folder for new files"""
        try:
            if self.is_monitoring:
                self.stop_monitoring()

            if not os.path.exists(folder_path):
                raise FileNotFoundError(f"Monitor folder does not exist: {folder_path}")

            # Auto-create sent folder if it doesn't exist
            sent_folder = os.path.join(folder_path, "sent")
            if not os.path.exists(sent_folder):
                os.makedirs(sent_folder, exist_ok=True)
                self.logger.info(f"Created sent folder: {sent_folder}")

            self.current_handler = FileHandler(callback, key_pattern, file_extensions)
            # Ensure a fresh Observer instance (watchdog threads cannot be restarted once stopped)
            if self.observer and self.observer.is_alive():
                self.stop_monitoring()
            self.observer = Observer()
            self.observer.schedule(self.current_handler, folder_path, recursive=False)
            self.observer.start()
            self.is_monitoring = True
            self.monitor_path = folder_path

            self.logger.info(f"Started monitoring folder: {folder_path}")
            self.logger.info(f"Using pattern: {key_pattern}")
            self.logger.info(f"Allowed extensions: {file_extensions}")

            return True

        except Exception as e:
            self.logger.error(f"Failed to start monitoring: {str(e)}")
            return False

    def stop_monitoring(self):
        """Stop folder monitoring"""
        try:
            if self.observer and self.observer.is_alive():
                self.observer.stop()
                self.observer.join(timeout=5)
            # Reset observer to allow future restarts
            self.observer = None

            self.is_monitoring = False
            self.current_handler = None
            self.monitor_path = None

            self.logger.info("Stopped folder monitoring")
            return True

        except Exception as e:
            self.logger.error(f"Error stopping monitoring: {str(e)}")
            return False

    def process_existing_files(self, folder_path: str, callback: Callable, key_pattern: str, file_extensions: list = None):
        """Process existing files in folder (useful for initial scan)"""
        try:
            if not os.path.exists(folder_path):
                return []

            processed_files = []
            pattern = re.compile(key_pattern)
            extensions = file_extensions or ['.pdf', '.xlsx', '.docx', '.txt']

            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)

                if os.path.isfile(file_path):
                    file_ext = os.path.splitext(filename)[1].lower()

                    if extensions and file_ext in extensions:
                        match = pattern.search(filename)
                        if match:
                            key = match.group(1) if match.groups() else match.group(0)
                            callback(file_path, key)
                            processed_files.append((file_path, key))

            self.logger.info(f"Processed {len(processed_files)} existing files")
            return processed_files

        except Exception as e:
            self.logger.error(f"Error processing existing files: {str(e)}")
            return []

    def move_file_to_sent(self, file_path: str, sent_folder: str) -> Optional[str]:
        """Move processed file to sent folder"""
        try:
            if not os.path.exists(sent_folder):
                os.makedirs(sent_folder, exist_ok=True)

            filename = os.path.basename(file_path)
            dest_path = os.path.join(sent_folder, filename)

            # Handle duplicate filenames
            counter = 1
            base_name, ext = os.path.splitext(filename)
            while os.path.exists(dest_path):
                new_filename = f"{base_name}_{counter}{ext}"
                dest_path = os.path.join(sent_folder, new_filename)
                counter += 1

            shutil.move(file_path, dest_path)
            self.logger.info(f"Moved file from {file_path} to {dest_path}")
            return dest_path

        except Exception as e:
            self.logger.error(f"Failed to move file {file_path} to sent folder: {str(e)}")
            return None

    def get_status(self) -> dict:
        """Get current monitoring status"""
        return {
            'is_monitoring': self.is_monitoring,
            'monitor_path': self.monitor_path,
            'observer_alive': self.observer.is_alive() if self.observer else False
        }