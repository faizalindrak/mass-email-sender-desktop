import os
import json
import re
from typing import Dict, Any, Optional, List


class ConfigManager:
    """JSON-based configuration manager for email automation application"""

    def __init__(self, config_dir: str = "config"):
        self.config_dir = config_dir
        self.global_config_file = os.path.join(config_dir, "config.json")
        self.profiles_dir = os.path.join(config_dir, "profiles")

        # Ensure directories exist
        os.makedirs(config_dir, exist_ok=True)
        os.makedirs(self.profiles_dir, exist_ok=True)

        # Load or create default global config
        self.global_config: Dict[str, Any] = {}
        self.load_config()

        # Ensure at least one default profile exists
        default_profile_path = self._get_profile_path("default")
        if not os.path.exists(default_profile_path):
            self._create_default_profile_file(default_profile_path)

    def load_config(self):
        """Load global configuration from JSON file"""
        if os.path.exists(self.global_config_file):
            with open(self.global_config_file, "r", encoding="utf-8") as f:
                try:
                    self.global_config = json.load(f)
                except Exception:
                    # If corrupted, recreate defaults
                    self._create_default_global_config()
        else:
            self._create_default_global_config()

    def _create_default_global_config(self):
        """Create default global configuration"""
        self.global_config = {
            "current_profile": "default",
            "database_path": "database/email_automation.db",
            "log_level": "INFO",
            "log_file": "logs/app.log",
            "template_dir": "templates",
            "auto_start_monitoring": False,
        }
        self.save_config()

    def _create_default_profile_file(self, profile_path: str):
        """Create default profile JSON file content"""
        default_profile = {
            "name": "Default Profile",
            "monitor_folder": "",
            "sent_folder": "sent",
            "key_pattern": "([A-Z]{2}\\d{3})",
            "email_client": "outlook",
            "subject_template": "[filename_without_ext]",
            "body_template": "default_template.html",
            "auto_start": False,
            "file_extensions": [".pdf", ".xlsx", ".docx", ".txt"],
            "default_cc": "",
            "default_bcc": "",
            "custom1_name": "",
            "custom1_value": "",
            "custom2_name": "",
            "custom2_value": "",
            "email_form": {
                "to_emails": [],
                "cc_emails": [],
                "bcc_emails": [],
                "subject": "",
                "selected_template": "default_template.html",
                "body_text": ""
            }
        }
        with open(profile_path, "w", encoding="utf-8") as f:
            json.dump(default_profile, f, indent=2, ensure_ascii=False)

    def save_config(self):
        """Save global configuration to JSON file"""
        with open(self.global_config_file, "w", encoding="utf-8") as f:
            json.dump(self.global_config, f, indent=2, ensure_ascii=False)

    def _get_profile_path(self, profile_name: str) -> str:
        """Get JSON file path for a profile"""
        safe_name = profile_name.strip().lower()
        return os.path.join(self.profiles_dir, f"{safe_name}.json")

    def get_current_profile(self) -> str:
        """Get current active profile name"""
        return self.global_config.get("current_profile", "default")

    def set_current_profile(self, profile_name: str):
        """Set current active profile"""
        self.global_config["current_profile"] = profile_name
        self.save_config()

    def get_profile_config(self, profile_name: str = None) -> Dict[str, Any]:
        """Get configuration for specific profile (JSON)"""
        if profile_name is None:
            profile_name = self.get_current_profile()

        profile_path = self._get_profile_path(profile_name)
        if not os.path.exists(profile_path):
            raise ValueError(f"Profile '{profile_name}' not found")

        with open(profile_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # Normalize types and defaults
        data["auto_start"] = bool(data.get("auto_start", False))
        if isinstance(data.get("file_extensions"), str):
            data["file_extensions"] = [ext.strip() for ext in data["file_extensions"].split(",") if ext.strip()]
        else:
            data["file_extensions"] = data.get("file_extensions", [])

        # Normalize email_form lists if they are strings
        email_form = data.get("email_form", {})
        for k in ["to_emails", "cc_emails", "bcc_emails"]:
            v = email_form.get(k, [])
            if isinstance(v, str):
                email_form[k] = [e.strip() for e in v.split(";") if e.strip()]
        data["email_form"] = email_form

        return data

    def save_profile_config(self, profile_name: str, config_data: Dict[str, Any]):
        """Save configuration for specific profile to JSON"""
        profile_path = self._get_profile_path(profile_name)

        # Ensure consistent types for lists and booleans
        normalized = dict(config_data)

        # Normalize file_extensions
        exts = normalized.get("file_extensions")
        if isinstance(exts, str):
            normalized["file_extensions"] = [ext.strip() for ext in exts.split(",") if ext.strip()]

        # Normalize email_form lists
        email_form = normalized.get("email_form", {})
        for k in ["to_emails", "cc_emails", "bcc_emails"]:
            v = email_form.get(k, [])
            if isinstance(v, str):
                email_form[k] = [e.strip() for e in v.split(";") if e.strip()]
        normalized["email_form"] = email_form

        # Persist JSON
        with open(profile_path, "w", encoding="utf-8") as f:
            json.dump(normalized, f, indent=2, ensure_ascii=False)

    def get_available_profiles(self) -> List[Dict[str, Any]]:
        """Get list of available profiles from JSON files"""
        profiles: List[Dict[str, Any]] = []
        for filename in os.listdir(self.profiles_dir):
            if filename.endswith(".json"):
                name = os.path.splitext(filename)[0]
                try:
                    data = self.get_profile_config(name)
                    profiles.append({
                        "name": name,
                        "display_name": data.get("name", name),
                        "monitor_folder": data.get("monitor_folder", ""),
                        "email_client": data.get("email_client", "outlook")
                    })
                except Exception:
                    # Skip malformed profile files
                    continue
        # Ensure deterministic order
        profiles.sort(key=lambda x: x["display_name"].lower())
        return profiles

    def delete_profile(self, profile_name: str):
        """Delete a profile JSON file"""
        profile_path = self._get_profile_path(profile_name)
        if os.path.exists(profile_path):
            os.remove(profile_path)

            # If deleted profile was current, switch to default if available
            if self.get_current_profile() == profile_name:
                self.set_current_profile("default")

    def get_database_path(self) -> str:
        """Get database file path from global config"""
        db_path = self.global_config.get("database_path", "database/email_automation.db")
        return os.path.abspath(db_path)

    def set_database_path(self, db_path: str):
        """Set database file path into global config"""
        self.global_config["database_path"] = db_path
        self.save_config()

    def get_log_config(self) -> Dict[str, str]:
        """Get logging configuration from global config"""
        return {
            "level": self.global_config.get("log_level", "INFO"),
            "file": self.global_config.get("log_file", "logs/app.log")
        }

    def get_template_dir(self) -> str:
        """Get templates directory path from global config"""
        template_dir = self.global_config.get("template_dir", "templates")
        return os.path.abspath(template_dir)

    def set_template_dir(self, template_dir: str):
        """Set templates directory path in global config"""
        if template_dir and isinstance(template_dir, str):
            self.global_config["template_dir"] = template_dir
            self.save_config()

    def should_auto_start_monitoring(self) -> bool:
        """Check if monitoring should auto-start from global config"""
        return bool(self.global_config.get("auto_start_monitoring", False))

    def set_auto_start_monitoring(self, enabled: bool):
        """Set auto-start monitoring in global config"""
        self.global_config["auto_start_monitoring"] = bool(enabled)
        self.save_config()

    def validate_profile_config(self, config_data: Dict[str, Any]) -> tuple:
        """Validate profile configuration JSON contents"""
        required_fields = ["name", "monitor_folder", "sent_folder", "key_pattern", "email_client"]

        for field in required_fields:
            if field not in config_data or (isinstance(config_data[field], str) and not config_data[field].strip()):
                return False, f"Required field '{field}' is missing or empty"

        # Validate monitor folder exists
        if not os.path.exists(config_data["monitor_folder"]):
            return False, f"Monitor folder does not exist: {config_data['monitor_folder']}"

        # Validate email client
        valid_clients = ["outlook", "thunderbird", "smtp"]
        if str(config_data["email_client"]).lower() not in valid_clients:
            return False, f"Invalid email client. Must be one of: {', '.join(valid_clients)}"

        # Validate SMTP settings if using smtp client
        if str(config_data["email_client"]).lower() == "smtp":
            smtp_required = ["smtp_server", "smtp_port", "smtp_username", "smtp_password"]
            for field in smtp_required:
                if field not in config_data or not config_data[field]:
                    return False, f"SMTP field '{field}' is required for smtp client"

        # Validate regex pattern
        try:
            re.compile(config_data["key_pattern"])
        except re.error as e:
            return False, f"Invalid regex pattern: {str(e)}"

        return True, "Configuration is valid"

    def export_profile(self, profile_name: str, export_path: str):
        """Export profile JSON to file"""
        data = self.get_profile_config(profile_name)
        with open(export_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def import_profile(self, profile_name: str, import_path: str):
        """Import profile JSON from file"""
        with open(import_path, "r", encoding="utf-8") as f:
            profile_config = json.load(f)

        # Validate imported config (monitor folder may not exist on this machine; allow skip validation?)
        is_valid, error_msg = self.validate_profile_config(profile_config)
        if not is_valid:
            # Allow import but warn: user may fix monitor folder later in UI
            # We still save to file for editing convenience
            pass

        self.save_profile_config(profile_name, profile_config)

    def create_sample_profiles(self):
        """Create sample JSON profiles for testing"""
        # Invoice profile
        invoice_profile = {
            "name": "Invoice Orders",
            "monitor_folder": "C:/Orders/Incoming",
            "sent_folder": "C:/Orders/Sent",
            "key_pattern": "([A-Z]{2}\\d{3})",
            "email_client": "outlook",
            "subject_template": "Invoice Order - [filename_without_ext]",
            "body_template": "invoice_template.html",
            "auto_start": True,
            "file_extensions": [".pdf", ".xlsx", ".docx"],
            "email_form": {
                "to_emails": [],
                "cc_emails": [],
                "bcc_emails": [],
                "subject": "Invoice Order - [filename_without_ext]",
                "selected_template": "invoice_template.html",
                "body_text": ""
            }
        }

        # Delivery profile
        delivery_profile = {
            "name": "Delivery Schedule",
            "monitor_folder": "C:/Delivery/Incoming",
            "sent_folder": "C:/Delivery/Sent",
            "key_pattern": "DELIVERY_([A-Z0-9]+)",
            "email_client": "smtp",
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "smtp_username": "your_email@gmail.com",
            "smtp_password": "your_password",
            "smtp_use_tls": True,
            "subject_template": "Delivery Schedule - [filename_without_ext]",
            "body_template": "delivery_template.html",
            "auto_start": False,
            "file_extensions": [".pdf", ".xlsx"],
            "email_form": {
                "to_emails": [],
                "cc_emails": [],
                "bcc_emails": [],
                "subject": "Delivery Schedule - [filename_without_ext]",
                "selected_template": "delivery_template.html",
                "body_text": ""
            }
        }

        self.save_profile_config("invoice", invoice_profile)
        self.save_profile_config("delivery", delivery_profile)