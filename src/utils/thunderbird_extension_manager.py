import os
import json
import zipfile
import tempfile
import shutil
import logging
import platform
import subprocess
import time
from typing import Optional, Dict, Any
import webbrowser

class ThunderbirdExtensionManager:
    """Manages Thunderbird WebExtension installation and configuration"""

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.extension_path = os.path.join(os.path.dirname(__file__), '..', '..', 'thunderbird_extension')
        self.manifest_path = os.path.join(self.extension_path, 'manifest.json')

    def get_extension_info(self) -> Dict[str, Any]:
        """Get extension information from manifest"""
        try:
            with open(self.manifest_path, 'r', encoding='utf-8') as f:
                manifest = json.load(f)
                return {
                    'name': manifest.get('name', 'Email Automation Extension'),
                    'version': manifest.get('version', '1.0.0'),
                    'description': manifest.get('description', ''),
                    'id': manifest.get('applications', {}).get('gecko', {}).get('id', 'email-automation@thunderbird')
                }
        except Exception as e:
            self.logger.error(f"Failed to read extension manifest: {str(e)}")
            return {}

    def build_extension(self, output_path: Optional[str] = None) -> Optional[str]:
        """Build extension as XPI file"""
        try:
            if not os.path.exists(self.extension_path):
                self.logger.error(f"Extension path does not exist: {self.extension_path}")
                return None

            # Create output filename if not provided
            if not output_path:
                extension_info = self.get_extension_info()
                extension_name = extension_info.get('name', 'email_automation').lower().replace(' ', '_')
                extension_version = extension_info.get('version', '1.0.0')
                output_path = f"{extension_name}-{extension_version}.xpi"

            # Create temporary directory for building
            with tempfile.TemporaryDirectory() as temp_dir:
                # Copy extension files to temp directory
                for item in os.listdir(self.extension_path):
                    src_path = os.path.join(self.extension_path, item)
                    dst_path = os.path.join(temp_dir, item)
                    if os.path.isfile(src_path):
                        shutil.copy2(src_path, dst_path)
                    elif os.path.isdir(src_path):
                        shutil.copytree(src_path, dst_path)

                # Create XPI file (ZIP archive)
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, temp_dir)
                            zipf.write(file_path, arcname)

                self.logger.info(f"Extension built successfully: {output_path}")
                return output_path

        except Exception as e:
            self.logger.error(f"Failed to build extension: {str(e)}")
            return None

    def get_thunderbird_profile_path(self) -> Optional[str]:
        """Get default Thunderbird profile path"""
        system = platform.system()

        if system == "Windows":
            base_path = os.path.expanduser(r"~\AppData\Roaming\Thunderbird\Profiles")
        elif system == "Darwin":  # macOS
            base_path = os.path.expanduser("~/Library/Thunderbird/Profiles")
        else:  # Linux
            base_path = os.path.expanduser("~/.thunderbird")

        if os.path.exists(base_path):
            # Find the default profile (usually ends with .default)
            for item in os.listdir(base_path):
                profile_dir = os.path.join(base_path, item)
                if os.path.isdir(profile_dir) and item.endswith('.default'):
                    return profile_dir

        return None

    def get_extensions_path(self, profile_path: Optional[str] = None) -> Optional[str]:
        """Get extensions directory path"""
        if not profile_path:
            profile_path = self.get_thunderbird_profile_path()

        if not profile_path:
            return None

        extensions_path = os.path.join(profile_path, 'extensions')
        os.makedirs(extensions_path, exist_ok=True)
        return extensions_path

    def install_extension(self, profile_path: Optional[str] = None) -> bool:
        """Install extension to Thunderbird profile"""
        try:
            # Build extension first
            xpi_path = self.build_extension()
            if not xpi_path:
                return False

            # Get extensions directory
            extensions_path = self.get_extensions_path(profile_path)
            if not extensions_path:
                self.logger.error("Could not find Thunderbird profile")
                return False

            # Copy XPI file to extensions directory
            extension_info = self.get_extension_info()
            extension_id = extension_info.get('id', 'email-automation@thunderbird')
            target_path = os.path.join(extensions_path, f"{extension_id}.xpi")

            shutil.copy2(xpi_path, target_path)

            self.logger.info(f"Extension installed to: {target_path}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to install extension: {str(e)}")
            return False

    def uninstall_extension(self, profile_path: Optional[str] = None) -> bool:
        """Uninstall extension from Thunderbird profile"""
        try:
            extensions_path = self.get_extensions_path(profile_path)
            if not extensions_path:
                return False

            extension_info = self.get_extension_info()
            extension_id = extension_info.get('id', 'email-automation@thunderbird')
            extension_file = os.path.join(extensions_path, f"{extension_id}.xpi")

            if os.path.exists(extension_file):
                os.remove(extension_file)
                self.logger.info(f"Extension uninstalled: {extension_file}")
                return True
            else:
                self.logger.warning(f"Extension not found: {extension_file}")
                return False

        except Exception as e:
            self.logger.error(f"Failed to uninstall extension: {str(e)}")
            return False

    def is_installed(self, profile_path: Optional[str] = None) -> bool:
        """Check if extension is installed"""
        try:
            extensions_path = self.get_extensions_path(profile_path)
            if not extensions_path:
                return False

            extension_info = self.get_extension_info()
            extension_id = extension_info.get('id', 'email-automation@thunderbird')
            extension_file = os.path.join(extensions_path, f"{extension_id}.xpi")

            return os.path.exists(extension_file)

        except Exception as e:
            self.logger.error(f"Failed to check installation status: {str(e)}")
            return False

    def open_thunderbird_extensions_page(self):
        """Open Thunderbird add-ons page"""
        try:
            # Try to open Thunderbird with about:addons page
            system = platform.system()

            if system == "Windows":
                thunderbird_path = r"C:\Program Files\Mozilla Thunderbird\thunderbird.exe"
                if not os.path.exists(thunderbird_path):
                    thunderbird_path = r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe"

                if os.path.exists(thunderbird_path):
                    subprocess.Popen([thunderbird_path, "-new-tab", "about:addons"])
                else:
                    webbrowser.open("https://addons.thunderbird.net/")

            elif system == "Darwin":  # macOS
                subprocess.Popen(["open", "-a", "Thunderbird", "--args", "-new-tab", "about:addons"])

            elif system == "Linux":
                subprocess.Popen(["thunderbird", "-new-tab", "about:addons"])

        except Exception as e:
            self.logger.error(f"Failed to open Thunderbird extensions page: {str(e)}")
            # Fallback to web
            webbrowser.open("https://addons.thunderbird.net/")

    def restart_thunderbird(self):
        """Restart Thunderbird"""
        try:
            system = platform.system()

            if system == "Windows":
                # Find Thunderbird process and restart it
                subprocess.call(["taskkill", "/f", "/im", "thunderbird.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                time.sleep(2)
                thunderbird_path = r"C:\Program Files\Mozilla Thunderbird\thunderbird.exe"
                if not os.path.exists(thunderbird_path):
                    thunderbird_path = r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe"
                if os.path.exists(thunderbird_path):
                    subprocess.Popen([thunderbird_path])

            elif system == "Darwin":  # macOS
                subprocess.call(["pkill", "Thunderbird"])
                time.sleep(2)
                subprocess.Popen(["open", "-a", "Thunderbird"])

            elif system == "Linux":
                subprocess.call(["pkill", "thunderbird"])
                time.sleep(2)
                subprocess.Popen(["thunderbird"])

            self.logger.info("Thunderbird restart initiated")

        except Exception as e:
            self.logger.error(f"Failed to restart Thunderbird: {str(e)}")

    def get_installation_instructions(self) -> str:
        """Get installation instructions for the extension"""
        return """
Thunderbird WebExtension Installation Instructions:

1. Build the Extension:
   - The extension will be built as an XPI file automatically

2. Install to Thunderbird:
   - The extension will be copied to your Thunderbird profile's extensions folder
   - Restart Thunderbird to load the extension

3. Manual Installation (if automatic fails):
   - Open Thunderbird
   - Go to Tools > Add-ons
   - Click the gear icon and select "Install Add-on From File"
   - Select the .xpi file from the thunderbird_extension folder
   - Restart Thunderbird

4. Verify Installation:
   - After restart, check Tools > Add-ons
   - Look for "Email Automation Thunderbird Extension"
   - Ensure it's enabled

5. Troubleshooting:
   - If extension doesn't appear, check the Browser Console (Ctrl+Shift+J)
   - Look for any error messages related to the extension
   - Try restarting Thunderbird again
   - Check that you're using Thunderbird 78 or later

Note: This extension requires Thunderbird 78+ with WebExtension support.
        """.strip()

    def validate_environment(self) -> Dict[str, Any]:
        """Validate environment for extension installation"""
        validation = {
            'thunderbird_installed': False,
            'profile_found': False,
            'extension_built': False,
            'permissions_ok': False,
            'issues': []
        }

        try:
            # Check if Thunderbird is installed
            system = platform.system()
            thunderbird_paths = []

            if system == "Windows":
                thunderbird_paths = [
                    r"C:\Program Files\Mozilla Thunderbird\thunderbird.exe",
                    r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe"
                ]
            elif system == "Darwin":
                thunderbird_paths = ["/Applications/Thunderbird.app/Contents/MacOS/thunderbird"]
            elif system == "Linux":
                thunderbird_paths = ["/usr/bin/thunderbird", "/usr/local/bin/thunderbird"]

            for path in thunderbird_paths:
                if os.path.exists(path):
                    validation['thunderbird_installed'] = True
                    break

            if not validation['thunderbird_installed']:
                validation['issues'].append("Thunderbird not found in standard locations")

            # Check profile
            profile_path = self.get_thunderbird_profile_path()
            if profile_path:
                validation['profile_found'] = True
            else:
                validation['issues'].append("No Thunderbird profile found")

            # Check extension build
            xpi_path = self.build_extension()
            if xpi_path and os.path.exists(xpi_path):
                validation['extension_built'] = True
            else:
                validation['issues'].append("Failed to build extension")

            # Check permissions
            if profile_path:
                extensions_path = self.get_extensions_path(profile_path)
                if extensions_path:
                    # Test write permission
                    test_file = os.path.join(extensions_path, 'test_write.tmp')
                    try:
                        with open(test_file, 'w') as f:
                            f.write('test')
                        os.remove(test_file)
                        validation['permissions_ok'] = True
                    except:
                        validation['issues'].append("No write permission to extensions folder")

        except Exception as e:
            validation['issues'].append(f"Validation error: {str(e)}")

        return validation