import os
import sys
import json
import shutil
import platform
import logging
from pathlib import Path
from typing import Optional

class ThunderbirdNativeConfig:
    """Utility class for configuring Thunderbird native messaging host"""

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.system = platform.system()
        
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )

    def get_thunderbird_native_messaging_dir(self) -> Path:
        """Get the Thunderbird native messaging directory for the current platform"""
        try:
            if self.system == "Windows":
                # Windows: %APPDATA%\Thunderbird\NativeMessagingHosts
                appdata = os.environ.get('APPDATA', '')
                if not appdata:
                    raise Exception("APPDATA environment variable not found")
                return Path(appdata) / "Thunderbird" / "NativeMessagingHosts"

            elif self.system == "Darwin":  # macOS
                # macOS: ~/Library/Application Support/Thunderbird/NativeMessagingHosts
                home = Path.home()
                return home / "Library" / "Application Support" / "Thunderbird" / "NativeMessagingHosts"

            elif self.system == "Linux":
                # Linux: ~/.thunderbird/native-messaging-hosts/
                home = Path.home()
                return home / ".thunderbird" / "native-messaging-hosts"

            else:
                raise Exception(f"Unsupported platform: {self.system}")

        except Exception as e:
            self.logger.error(f"Error getting Thunderbird native messaging directory: {str(e)}")
            raise

    def get_python_executable(self) -> str:
        """Get the Python executable path"""
        try:
            # Use the current Python interpreter
            return sys.executable
        except Exception as e:
            self.logger.error(f"Error getting Python executable: {str(e)}")
            raise

    def get_native_host_script_path(self) -> Path:
        """Get the path to the native host script"""
        try:
            # Get the directory where this script is located
            current_dir = Path(__file__).parent.parent
            return current_dir / "core" / "thunderbird_native_client.py"
        except Exception as e:
            self.logger.error(f"Error getting native host script path: {str(e)}")
            raise

    def get_thunderbird_profile_path(self) -> Optional[str]:
        """Get default Thunderbird profile path by reading profiles.ini"""
        if self.system == "Windows":
            thunderbird_dir = os.path.expanduser(r"~\AppData\Roaming\Thunderbird")
            profiles_dir = os.path.join(thunderbird_dir, "Profiles")
            profiles_ini = os.path.join(thunderbird_dir, "profiles.ini")
        elif self.system == "Darwin":  # macOS
            thunderbird_dir = os.path.expanduser("~/Library/Thunderbird")
            profiles_dir = os.path.join(thunderbird_dir, "Profiles")
            profiles_ini = os.path.join(thunderbird_dir, "profiles.ini")
        else:  # Linux
            thunderbird_dir = os.path.expanduser("~/.thunderbird")
            profiles_dir = thunderbird_dir
            profiles_ini = os.path.join(thunderbird_dir, "profiles.ini")

        # First, try to read profiles.ini to find the default profile
        default_profile_path = self._get_default_profile_from_ini(profiles_ini, profiles_dir)
        if default_profile_path and os.path.exists(default_profile_path):
            self.logger.info(f"Found default profile from profiles.ini: {default_profile_path}")
            return default_profile_path

        # Fallback: look for profiles ending with .default or .default-esr
        if os.path.exists(profiles_dir):
            # Priority order: .default-esr, .default, .default-release
            priority_suffixes = ['.default-esr', '.default', '.default-release']

            for suffix in priority_suffixes:
                for item in os.listdir(profiles_dir):
                    profile_dir = os.path.join(profiles_dir, item)
                    if os.path.isdir(profile_dir) and item.endswith(suffix):
                        self.logger.info(f"Found profile with suffix {suffix}: {profile_dir}")
                        return profile_dir

        self.logger.warning("No suitable Thunderbird profile found")
        return None

    def _get_default_profile_from_ini(self, profiles_ini_path: str, profiles_dir: str) -> Optional[str]:
        """Parse profiles.ini to find the default profile"""
        try:
            if not os.path.exists(profiles_ini_path):
                return None

            import configparser
            config = configparser.ConfigParser()
            config.read(profiles_ini_path, encoding='utf-8')

            # First, check for the Install section's Default setting (last used profile)
            if config.has_section('InstallD78BF5DD33499EC2'):  # Thunderbird's install section
                default_path = config.get('InstallD78BF5DD33499EC2', 'Default', fallback=None)
                if default_path:
                    # Remove 'Profiles/' prefix if present
                    if default_path.startswith('Profiles/'):
                        default_path = default_path[9:]
                    profile_path = os.path.join(profiles_dir, default_path)
                    if os.path.exists(profile_path):
                        return profile_path

            # Fallback: look for profiles marked as Default=1
            for section in config.sections():
                if section.startswith('Profile'):
                    is_default = config.getboolean(section, 'Default', fallback=False)
                    if is_default:
                        profile_path_rel = config.get(section, 'Path', fallback='')
                        if profile_path_rel.startswith('Profiles/'):
                            profile_path_rel = profile_path_rel[9:]
                        profile_path = os.path.join(profiles_dir, profile_path_rel)
                        if os.path.exists(profile_path):
                            return profile_path

        except Exception as e:
            self.logger.warning(f"Error reading profiles.ini: {str(e)}")

        return None

    def create_native_host_config(self, output_path: Path) -> bool:
        """Create the native host configuration file"""
        try:
            python_executable = self.get_python_executable()
            script_path = self.get_native_host_script_path()
            
            config = {
                "name": "email_automation_native_host",
                "description": "Email Automation Native Messaging Host",
                "path": python_executable,
                "type": "stdio",
                "allowed_extensions": ["email-automation@thunderbird"]
            }
            
            # For Windows, we need to use the script path as an argument
            if self.system == "Windows":
                config["arguments"] = [str(script_path)]
            
            # Write the configuration file
            with open(output_path, 'w') as f:
                json.dump(config, f, indent=2)
            
            self.logger.info(f"Created native host configuration at: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error creating native host configuration: {str(e)}")
            return False

    def install_native_host(self) -> bool:
        """Install the native messaging host"""
        try:
            # Get the Thunderbird native messaging directory
            thunderbird_dir = self.get_thunderbird_native_messaging_dir()
            
            # Create the directory if it doesn't exist
            thunderbird_dir.mkdir(parents=True, exist_ok=True)
            
            # Path to the configuration file
            config_path = thunderbird_dir / "email_automation_native_host.json"
            
            # Create the configuration file
            if not self.create_native_host_config(config_path):
                return False
            
            self.logger.info(f"Native messaging host installed successfully at: {config_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error installing native messaging host: {str(e)}")
            return False

    def check_installation(self) -> bool:
        """Check if the native messaging host is properly installed"""
        try:
            # Get the Thunderbird native messaging directory
            thunderbird_dir = self.get_thunderbird_native_messaging_dir()
            
            # Check if the configuration file exists
            config_path = thunderbird_dir / "email_automation_native_host.json"
            if not config_path.exists():
                self.logger.warning("Native messaging host configuration file not found")
                return False
            
            # Check if the Python executable exists
            python_executable = self.get_python_executable()
            if not os.path.exists(python_executable):
                self.logger.warning(f"Python executable not found: {python_executable}")
                return False
            
            # Check if the native host script exists
            script_path = self.get_native_host_script_path()
            if not script_path.exists():
                self.logger.warning(f"Native host script not found: {script_path}")
                return False
            
            # Load and validate the configuration
            with open(config_path, 'r') as f:
                config = json.load(f)
            
            required_keys = ["name", "description", "path", "type", "allowed_extensions"]
            for key in required_keys:
                if key not in config:
                    self.logger.warning(f"Missing required key in configuration: {key}")
                    return False
            
            self.logger.info("Native messaging host installation is valid")
            return True
            
        except Exception as e:
            self.logger.error(f"Error checking native messaging host installation: {str(e)}")
            return False

    def uninstall_native_host(self) -> bool:
        """Uninstall the native messaging host"""
        try:
            # Get the Thunderbird native messaging directory
            thunderbird_dir = self.get_thunderbird_native_messaging_dir()
            
            # Path to the configuration file
            config_path = thunderbird_dir / "email_automation_native_host.json"
            
            # Remove the configuration file if it exists
            if config_path.exists():
                config_path.unlink()
                self.logger.info(f"Removed native messaging host configuration: {config_path}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error uninstalling native messaging host: {str(e)}")
            return False

    def print_status(self):
        """Print the current status of the native messaging host"""
        print("Thunderbird Native Messaging Host Status:")
        print("=" * 50)
        
        try:
            # Get the Thunderbird native messaging directory
            thunderbird_dir = self.get_thunderbird_native_messaging_dir()
            print(f"Thunderbird Native Messaging Directory: {thunderbird_dir}")
            
            # Check if the configuration file exists
            config_path = thunderbird_dir / "email_automation_native_host.json"
            if config_path.exists():
                print(f"Configuration File: {config_path} [OK]")

                # Load and display configuration
                with open(config_path, 'r') as f:
                    config = json.load(f)

                print(f"  Name: {config.get('name', 'N/A')}")
                print(f"  Description: {config.get('description', 'N/A')}")
                print(f"  Path: {config.get('path', 'N/A')}")
                print(f"  Type: {config.get('type', 'N/A')}")
                print(f"  Allowed Extensions: {config.get('allowed_extensions', 'N/A')}")

                if self.system == "Windows" and "arguments" in config:
                    print(f"  Arguments: {config.get('arguments', 'N/A')}")
            else:
                print(f"Configuration File: {config_path} [MISSING]")

            # Check if the Python executable exists
            python_executable = self.get_python_executable()
            if os.path.exists(python_executable):
                print(f"Python Executable: {python_executable} [OK]")
            else:
                print(f"Python Executable: {python_executable} [MISSING]")

            # Check if the native host script exists
            script_path = self.get_native_host_script_path()
            if script_path.exists():
                print(f"Native Host Script: {script_path} [OK]")
            else:
                print(f"Native Host Script: {script_path} [MISSING]")

            # Overall status
            if self.check_installation():
                print("\nStatus: INSTALLED [OK]")
            else:
                print("\nStatus: NOT INSTALLED [ERROR]")
                
        except Exception as e:
            print(f"\nError checking status: {str(e)}")
            print("Status: ERROR ✗")

def main():
    """Main function for command-line usage"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Configure Thunderbird Native Messaging Host")
    parser.add_argument("action", choices=["install", "uninstall", "check", "status"],
                       help="Action to perform")
    
    args = parser.parse_args()
    
    config = ThunderbirdNativeConfig()
    
    if args.action == "install":
        if config.install_native_host():
            print("Native messaging host installed successfully!")
        else:
            print("Failed to install native messaging host.")
            sys.exit(1)
    
    elif args.action == "uninstall":
        if config.uninstall_native_host():
            print("Native messaging host uninstalled successfully!")
        else:
            print("Failed to uninstall native messaging host.")
            sys.exit(1)
    
    elif args.action == "check":
        if config.check_installation():
            print("Native messaging host is properly installed.")
        else:
            print("Native messaging host is not properly installed.")
            sys.exit(1)
    
    elif args.action == "status":
        config.print_status()

if __name__ == "__main__":
    main()