import configparser
import os
import json
from typing import Dict, Any, Optional

class ConfigManager:
    """Configuration manager for email automation application"""

    def __init__(self, config_dir: str = "config"):
        self.config_dir = config_dir
        self.config_file = os.path.join(config_dir, "default.ini")
        self.profiles_dir = os.path.join(config_dir, "profiles")
        self.config = configparser.ConfigParser()

        # Ensure directories exist
        os.makedirs(config_dir, exist_ok=True)
        os.makedirs(self.profiles_dir, exist_ok=True)

        # Load or create default configuration
        self.load_config()

    def load_config(self):
        """Load configuration from file"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding='utf-8')
        else:
            self.create_default_config()

    def create_default_config(self):
        """Create default configuration file"""
        self.config['DEFAULT'] = {
            'current_profile': 'default',
            'database_path': 'database/email_automation.db',
            'log_level': 'INFO',
            'log_file': 'logs/app.log',
            'template_dir': 'templates',
            'auto_start_monitoring': 'false'
        }

        # Default profile
        self.config['profile_default'] = {
            'name': 'Default Profile',
            'monitor_folder': '',
            'sent_folder': 'sent',
            'key_pattern': r'([A-Z]{2}\d{3})',
            'email_client': 'outlook',
            'subject_template': 'Document - [filename_without_ext]',
            'body_template': 'default_template.html',
            'auto_start': 'false',
            'file_extensions': '.pdf,.xlsx,.docx,.txt'
        }

        self.save_config()

    def save_config(self):
        """Save configuration to file"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def get_current_profile(self) -> str:
        """Get current active profile name"""
        return self.config.get('DEFAULT', 'current_profile', fallback='default')

    def set_current_profile(self, profile_name: str):
        """Set current active profile"""
        self.config.set('DEFAULT', 'current_profile', profile_name)
        self.save_config()

    def get_profile_config(self, profile_name: str = None) -> Dict[str, Any]:
        """Get configuration for specific profile"""
        if profile_name is None:
            profile_name = self.get_current_profile()

        section_name = f'profile_{profile_name}'

        if not self.config.has_section(section_name):
            raise ValueError(f"Profile '{profile_name}' not found")

        profile_config = dict(self.config[section_name])

        # Convert string values to appropriate types
        profile_config['auto_start'] = self.config.getboolean(section_name, 'auto_start', fallback=False)
        profile_config['file_extensions'] = [ext.strip() for ext in profile_config.get('file_extensions', '').split(',') if ext.strip()]

        # Add SMTP settings if available
        smtp_keys = ['smtp_server', 'smtp_port', 'smtp_username', 'smtp_password', 'smtp_use_tls']
        for key in smtp_keys:
            if self.config.has_option(section_name, key):
                if key == 'smtp_port':
                    profile_config[key] = self.config.getint(section_name, key)
                elif key == 'smtp_use_tls':
                    profile_config[key] = self.config.getboolean(section_name, key, fallback=True)
                else:
                    profile_config[key] = self.config.get(section_name, key)

        return profile_config

    def save_profile_config(self, profile_name: str, config_data: Dict[str, Any]):
        """Save configuration for specific profile"""
        section_name = f'profile_{profile_name}'

        if not self.config.has_section(section_name):
            self.config.add_section(section_name)

        for key, value in config_data.items():
            if isinstance(value, list):
                self.config.set(section_name, key, ','.join(value))
            elif isinstance(value, bool):
                self.config.set(section_name, key, str(value).lower())
            else:
                self.config.set(section_name, key, str(value))

        self.save_config()

    def get_available_profiles(self) -> list:
        """Get list of available profiles"""
        profiles = []
        for section in self.config.sections():
            if section.startswith('profile_'):
                profile_name = section[8:]  # Remove 'profile_' prefix
                profile_data = self.get_profile_config(profile_name)
                profiles.append({
                    'name': profile_name,
                    'display_name': profile_data.get('name', profile_name),
                    'monitor_folder': profile_data.get('monitor_folder', ''),
                    'email_client': profile_data.get('email_client', 'outlook')
                })
        return profiles

    def delete_profile(self, profile_name: str):
        """Delete a profile"""
        section_name = f'profile_{profile_name}'
        if self.config.has_section(section_name):
            self.config.remove_section(section_name)
            self.save_config()

            # If deleted profile was current, switch to default
            if self.get_current_profile() == profile_name:
                self.set_current_profile('default')

    def get_database_path(self) -> str:
        """Get database file path"""
        db_path = self.config.get('DEFAULT', 'database_path', fallback='database/email_automation.db')
        return os.path.abspath(db_path)

    def get_log_config(self) -> Dict[str, str]:
        """Get logging configuration"""
        return {
            'level': self.config.get('DEFAULT', 'log_level', fallback='INFO'),
            'file': self.config.get('DEFAULT', 'log_file', fallback='logs/app.log')
        }

    def get_template_dir(self) -> str:
        """Get templates directory path"""
        template_dir = self.config.get('DEFAULT', 'template_dir', fallback='templates')
        return os.path.abspath(template_dir)

    def should_auto_start_monitoring(self) -> bool:
        """Check if monitoring should auto-start"""
        return self.config.getboolean('DEFAULT', 'auto_start_monitoring', fallback=False)

    def set_auto_start_monitoring(self, enabled: bool):
        """Set auto-start monitoring setting"""
        self.config.set('DEFAULT', 'auto_start_monitoring', str(enabled).lower())
        self.save_config()

    def validate_profile_config(self, config_data: Dict[str, Any]) -> tuple[bool, str]:
        """Validate profile configuration"""
        required_fields = ['name', 'monitor_folder', 'sent_folder', 'key_pattern', 'email_client']

        for field in required_fields:
            if field not in config_data or not config_data[field]:
                return False, f"Required field '{field}' is missing or empty"

        # Validate monitor folder exists
        if not os.path.exists(config_data['monitor_folder']):
            return False, f"Monitor folder does not exist: {config_data['monitor_folder']}"

        # Validate email client
        valid_clients = ['outlook', 'thunderbird', 'smtp']
        if config_data['email_client'].lower() not in valid_clients:
            return False, f"Invalid email client. Must be one of: {', '.join(valid_clients)}"

        # Validate SMTP settings if using thunderbird/smtp
        if config_data['email_client'].lower() in ['thunderbird', 'smtp']:
            smtp_required = ['smtp_server', 'smtp_port', 'smtp_username', 'smtp_password']
            for field in smtp_required:
                if field not in config_data or not config_data[field]:
                    return False, f"SMTP field '{field}' is required for thunderbird/smtp client"

        # Validate regex pattern
        try:
            import re
            re.compile(config_data['key_pattern'])
        except re.error as e:
            return False, f"Invalid regex pattern: {str(e)}"

        return True, "Configuration is valid"

    def export_profile(self, profile_name: str, export_path: str):
        """Export profile to file"""
        profile_config = self.get_profile_config(profile_name)
        with open(export_path, 'w', encoding='utf-8') as f:
            json.dump(profile_config, f, indent=2, ensure_ascii=False)

    def import_profile(self, profile_name: str, import_path: str):
        """Import profile from file"""
        with open(import_path, 'r', encoding='utf-8') as f:
            profile_config = json.load(f)

        # Validate imported config
        is_valid, error_msg = self.validate_profile_config(profile_config)
        if not is_valid:
            raise ValueError(f"Invalid profile configuration: {error_msg}")

        self.save_profile_config(profile_name, profile_config)

    def create_sample_profiles(self):
        """Create sample profiles for testing"""
        # Invoice profile
        invoice_profile = {
            'name': 'Invoice Orders',
            'monitor_folder': 'C:/Orders/Incoming',
            'sent_folder': 'C:/Orders/Sent',
            'key_pattern': r'([A-Z]{2}\d{3})',
            'email_client': 'outlook',
            'subject_template': 'Invoice Order - [filename_without_ext]',
            'body_template': 'invoice_template.html',
            'auto_start': True,
            'file_extensions': ['.pdf', '.xlsx', '.docx']
        }

        # Delivery profile
        delivery_profile = {
            'name': 'Delivery Schedule',
            'monitor_folder': 'C:/Delivery/Incoming',
            'sent_folder': 'C:/Delivery/Sent',
            'key_pattern': r'DELIVERY_([A-Z0-9]+)',
            'email_client': 'smtp',
            'smtp_server': 'smtp.gmail.com',
            'smtp_port': 587,
            'smtp_username': 'your_email@gmail.com',
            'smtp_password': 'your_password',
            'smtp_use_tls': True,
            'subject_template': 'Delivery Schedule - [filename_without_ext]',
            'body_template': 'delivery_template.html',
            'auto_start': False,
            'file_extensions': ['.pdf', '.xlsx']
        }

        self.save_profile_config('invoice', invoice_profile)
        self.save_profile_config('delivery', delivery_profile)