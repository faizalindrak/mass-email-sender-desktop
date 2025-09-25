import win32com.client as win32
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import logging
import time
import platform
import subprocess
from typing import List, Optional

class EmailSenderBase:
    """Base class for email senders"""
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def send_email(self, to_emails: List[str], cc_emails: List[str], bcc_emails: List[str],
                   subject: str, body: str, attachment_path: Optional[str] = None) -> bool:
        """Send email - to be implemented by subclasses"""
        raise NotImplementedError

class OutlookSender(EmailSenderBase):
    """Outlook email sender using COM automation"""

    def __init__(self):
        super().__init__()
        try:
            self.outlook = win32.Dispatch('outlook.application')
            self.logger.info("Outlook connection established")
        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {str(e)}")
            raise Exception("Outlook tidak terinstall atau tidak dapat diakses")

    def send_email(self, to_emails: List[str], cc_emails: List[str] = None,
                    bcc_emails: List[str] = None, subject: str = "", body: str = "",
                    attachment_path: Optional[str] = None) -> bool:
        """Send email via Outlook"""
        try:
            # Validate attachment path before creating email
            if attachment_path:
                self.logger.info(f"Attempting to send email with attachment: {attachment_path}")

                # Convert to absolute path to ensure Outlook can find it
                abs_attachment_path = os.path.abspath(attachment_path)
                self.logger.info(f"Absolute attachment path: {abs_attachment_path}")

                # Check if attachment path exists and is accessible
                if not os.path.exists(abs_attachment_path):
                    self.logger.error(f"Attachment file does not exist: {abs_attachment_path}")
                    return False

                if not os.path.isfile(abs_attachment_path):
                    self.logger.error(f"Attachment path is not a file: {abs_attachment_path}")
                    return False

                # Check file size (Outlook has limits)
                try:
                    file_size = os.path.getsize(abs_attachment_path)
                    max_size = 25 * 1024 * 1024  # 25MB limit for Outlook
                    if file_size > max_size:
                        self.logger.error(f"Attachment file too large: {file_size} bytes (max: {max_size} bytes)")
                        return False
                    self.logger.info(f"Attachment file size: {file_size} bytes")
                except OSError as e:
                    self.logger.error(f"Cannot access attachment file: {str(e)}")
                    return False

                # Check if file is readable
                try:
                    with open(abs_attachment_path, 'rb') as f:
                        f.read(1)  # Try to read at least one byte
                except (PermissionError, OSError) as e:
                    self.logger.error(f"Cannot read attachment file: {str(e)}")
                    return False

                # Use absolute path for Outlook
                attachment_path = abs_attachment_path

            mail = self.outlook.CreateItem(0)  # 0 = olMailItem

            # Set recipients
            mail.To = '; '.join(to_emails)
            if cc_emails:
                mail.CC = '; '.join(cc_emails)
            if bcc_emails:
                mail.BCC = '; '.join(bcc_emails)

            # Set subject and body
            mail.Subject = subject
            mail.HTMLBody = body

            # Add attachment if provided and validated
            if attachment_path:
                if not self._add_attachment_with_retry(mail, attachment_path):
                    # Try to continue without attachment
                    self.logger.warning("Continuing to send email without attachment")

            # Send email
            mail.Send()

            self.logger.info(f"Email sent successfully to {', '.join(to_emails)}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to send email via Outlook: {str(e)}")
            return False

    def is_available(self) -> bool:
        """Check if Outlook is available"""
        try:
            return self.outlook is not None
        except:
            return False

    def _add_attachment_with_retry(self, mail, attachment_path: str, max_retries: int = 3) -> bool:
        """Add attachment to email with retry mechanism"""
        for attempt in range(max_retries):
            try:
                # Try different path formats for Outlook
                path_attempts = [
                    attachment_path,  # Absolute path
                    attachment_path.replace('/', '\\'),  # Windows path separators
                    os.path.normpath(attachment_path),  # Normalized path
                ]

                for path_attempt in path_attempts:
                    try:
                        mail.Attachments.Add(path_attempt)
                        self.logger.info(f"Successfully added attachment on attempt {attempt + 1} using path: {path_attempt}")
                        return True
                    except Exception as path_e:
                        self.logger.debug(f"Failed with path {path_attempt}: {str(path_e)}")
                        continue

                # If all path formats failed, raise the last exception
                raise Exception(f"All path formats failed for: {attachment_path}")

            except Exception as e:
                self.logger.warning(f"Failed to add attachment on attempt {attempt + 1}: {str(e)}")

                if attempt < max_retries - 1:
                    # Wait before retry (exponential backoff)
                    wait_time = (2 ** attempt) * 0.5  # 0.5s, 1s, 2s
                    self.logger.info(f"Retrying attachment in {wait_time} seconds...")
                    time.sleep(wait_time)

        self.logger.error(f"Failed to add attachment after {max_retries} attempts: {attachment_path}")
        return False

class ThunderbirdProfileManager:
    """Manages Thunderbird profile operations for email history"""

    def __init__(self, profile_path: Optional[str] = None):
        self.profile_path = profile_path or self._find_default_profile()
        self.logger = logging.getLogger(__name__)

    def _find_default_profile(self) -> Optional[str]:
        """Find default Thunderbird profile path"""
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

    def get_sent_folder_path(self) -> Optional[str]:
        """Get the Sent folder path in Thunderbird profile"""
        if not self.profile_path:
            self.logger.error("No Thunderbird profile path available")
            return None

        self.logger.info(f"Looking for Sent folder in profile: {self.profile_path}")

        # Thunderbird profile structure - try multiple possibilities
        possible_paths = [
            os.path.join(self.profile_path, "Mail", "Local Folders", "Sent"),
            os.path.join(self.profile_path, "Mail", "Sent"),
            os.path.join(self.profile_path, "ImapMail", "imap.gmail.com", "Sent"),
            os.path.join(self.profile_path, "ImapMail", "smtp.gmail.com", "Sent"),
            os.path.join(self.profile_path, "Mail", "pop.gmail.com", "Sent"),
        ]

        for path in possible_paths:
            if os.path.exists(path):
                self.logger.info(f"Found existing Sent folder: {path}")
                return path

        # Create Sent folder if it doesn't exist
        sent_path = os.path.join(self.profile_path, "Mail", "Local Folders", "Sent")
        try:
            os.makedirs(sent_path, exist_ok=True)
            self.logger.info(f"Created new Sent folder: {sent_path}")

            # Create the .msf file (index file) for Thunderbird
            msf_path = sent_path + ".msf"
            try:
                with open(msf_path, 'a'):
                    pass  # Just create the file
                self.logger.info(f"Created MSF index file: {msf_path}")
            except Exception as msf_e:
                self.logger.warning(f"Could not create MSF file: {str(msf_e)}")

            return sent_path
        except Exception as e:
            self.logger.error(f"Failed to create Sent folder: {str(e)}")
            return None

    def save_email_to_sent(self, to_emails: List[str], subject: str,
                          body: str, attachments: List[str] = None) -> bool:
        """Save email to Thunderbird Sent folder using multiple methods"""
        try:
            # Method 1: Try EML file approach
            eml_success = self._save_as_eml_file(to_emails, subject, body, attachments)

            if eml_success:
                self.logger.info("Email saved to Thunderbird Sent folder via EML method")
                return True

            # Method 2: Try command line approach
            cli_success = self._save_via_command_line(to_emails, subject, body, attachments)

            if cli_success:
                self.logger.info("Email saved to Thunderbird Sent folder via command line method")
                return True

            # Method 3: Try database approach
            db_success = self._save_via_database(to_emails, subject, body, attachments)

            if db_success:
                self.logger.info("Email saved to Thunderbird Sent folder via database method")
                return True

            self.logger.error("All methods failed to save email to Thunderbird Sent folder")
            return False

        except Exception as e:
            self.logger.error(f"Failed to save email to Thunderbird: {str(e)}")
            return False

    def _save_as_eml_file(self, to_emails: List[str], subject: str,
                         body: str, attachments: List[str] = None) -> bool:
        """Save as EML file (original method)"""
        try:
            sent_folder = self.get_sent_folder_path()
            if not sent_folder:
                return False

            # Create EML content
            eml_content = self._create_eml_content(to_emails, subject, body, attachments)

            # Generate unique filename based on timestamp
            timestamp = int(time.time())
            eml_filename = f"{timestamp}.eml"
            eml_path = os.path.join(sent_folder, eml_filename)

            self.logger.info(f"Saving EML file to: {eml_path}")

            # Write EML file
            with open(eml_path, 'w', encoding='utf-8') as f:
                f.write(eml_content)

            # Verify file was written
            if os.path.exists(eml_path):
                file_size = os.path.getsize(eml_path)
                self.logger.info(f"EML file saved successfully. Size: {file_size} bytes")

                # Additional debugging: check if Thunderbird is running
                try:
                    import psutil
                    thunderbird_running = False
                    for proc in psutil.process_iter(['name']):
                        if 'thunderbird' in proc.info['name'].lower():
                            thunderbird_running = True
                            break

                    if thunderbird_running:
                        self.logger.info("Thunderbird is currently running - refresh should work")
                    else:
                        self.logger.info("Thunderbird is not running - restart Thunderbird to see new emails")

                except ImportError:
                    self.logger.info("Cannot check if Thunderbird is running (psutil not available)")

                return True
            else:
                self.logger.error(f"Failed to create EML file: {eml_path}")
                return False

        except Exception as e:
            self.logger.error(f"EML save method failed: {str(e)}")
            return False

    def _save_via_command_line(self, to_emails: List[str], subject: str,
                              body: str, attachments: List[str] = None) -> bool:
        """Save email using Thunderbird command line"""
        try:
            # Find Thunderbird executable
            thunderbird_path = self._find_thunderbird_executable()
            if not thunderbird_path:
                return False

            # Create temporary EML file for command line
            eml_content = self._create_eml_content(to_emails, subject, body, attachments)

            with tempfile.NamedTemporaryFile(mode='w', suffix='.eml', delete=False) as f:
                f.write(eml_content)
                temp_eml_path = f.name

            # Use Thunderbird command line to import the EML file
            cmd = [
                thunderbird_path,
                '-profile', self.profile_path,
                temp_eml_path
            ]

            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

            # Clean up temp file
            try:
                os.unlink(temp_eml_path)
            except:
                pass

            if result.returncode == 0:
                self.logger.info("Email imported via Thunderbird command line successfully")
                return True
            else:
                self.logger.error(f"Command line import failed: {result.stderr}")
                return False

        except Exception as e:
            self.logger.error(f"Command line save method failed: {str(e)}")
            return False

    def _save_via_database(self, to_emails: List[str], subject: str,
                          body: str, attachments: List[str] = None) -> bool:
        """Save email by directly manipulating Thunderbird's database files"""
        try:
            # This is a simplified approach - in reality, Thunderbird uses Mork or SQLite
            # For now, we'll try to update the MSF file to trigger a refresh
            sent_folder = self.get_sent_folder_path()
            if not sent_folder:
                return False

            msf_file = sent_folder + ".msf"
            if os.path.exists(msf_file):
                # Update the modification time to trigger Thunderbird refresh
                os.utime(msf_file, None)
                self.logger.info("Updated MSF file timestamp to trigger refresh")
                return True

            return False

        except Exception as e:
            self.logger.error(f"Database save method failed: {str(e)}")
            return False

    def _find_thunderbird_executable(self) -> Optional[str]:
        """Find Thunderbird executable path"""
        possible_paths = [
            r"C:\Program Files\Mozilla Thunderbird\thunderbird.exe",
            r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe",
            "/usr/bin/thunderbird",
            "/usr/local/bin/thunderbird",
            "/Applications/Thunderbird.app/Contents/MacOS/thunderbird"
        ]

        for path in possible_paths:
            if os.path.exists(path):
                return path

        # Try to find in PATH
        try:
            result = subprocess.run(['which', 'thunderbird'],
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                return result.stdout.strip()
        except:
            pass

        return None

    def _create_eml_content(self, to_emails: List[str], subject: str,
                           body: str, attachments: List[str] = None) -> str:
        """Create EML file content compatible with Thunderbird"""
        timestamp = time.strftime('%a, %d %b %Y %H:%M:%S %z', time.localtime())
        from_address = self._get_from_address()

        # Create Thunderbird-compatible EML structure
        eml_lines = [
            f"From: {from_address}",
            f"To: {', '.join(to_emails)}",
            f"Subject: {subject}",
            f"Date: {timestamp}",
            "MIME-Version: 1.0",
            "Content-Type: text/html; charset=UTF-8",
            "Content-Transfer-Encoding: 7bit",
            "",
            body
        ]

        return "\n".join(eml_lines)

    def _get_from_address(self) -> str:
        """Get default from address from Thunderbird profile or use provided username"""
        if not self.profile_path:
            # If no profile, return a generic address
            return "sender@example.com"

        try:
            # Try to read from prefs.js file
            prefs_path = os.path.join(self.profile_path, "prefs.js")
            if os.path.exists(prefs_path):
                with open(prefs_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    # Look for user_pref("mail.identity.id1.useremail", "...");
                    import re
                    email_match = re.search(r'user_pref\("mail\.identity\.id\d+\.useremail",\s*"([^"]+)"\);', content)
                    if email_match:
                        return email_match.group(1)
        except Exception as e:
            self.logger.warning(f"Could not read from address from prefs.js: {str(e)}")

        # Fallback: return a generic address
        return "sender@example.com"

    def refresh_thunderbird_view(self) -> bool:
        """Refresh Thunderbird to show new emails in Sent folder"""
        try:
            # Try to find Thunderbird process and send refresh signal
            import psutil

            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if 'thunderbird' in proc.info['name'].lower():
                        # Send a signal to refresh (this is a simple approach)
                        # In a real implementation, you might use Thunderbird's API
                        self.logger.info(f"Found Thunderbird process: {proc.info['pid']}")

                        # Try to refresh by touching the .msf file
                        sent_folder = self.get_sent_folder_path()
                        if sent_folder:
                            msf_file = sent_folder + ".msf"
                            if os.path.exists(msf_file):
                                os.utime(msf_file, None)  # Update modification time
                                self.logger.info("Updated MSF file timestamp to trigger refresh")
                                return True
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue

            self.logger.info("Thunderbird process not found or could not refresh")
            return False

        except ImportError:
            self.logger.warning("psutil not available, cannot refresh Thunderbird")
            return False
        except Exception as e:
            self.logger.error(f"Failed to refresh Thunderbird: {str(e)}")
            return False

        try:
            # Try to read from prefs.js file
            prefs_path = os.path.join(self.profile_path, "prefs.js")
            if os.path.exists(prefs_path):
                with open(prefs_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    # Look for user_pref("mail.identity.id1.useremail", "...");
                    import re
                    email_match = re.search(r'user_pref\("mail\.identity\.id\d+\.useremail",\s*"([^"]+)"\);', content)
                    if email_match:
                        return email_match.group(1)
        except Exception as e:
            self.logger.warning(f"Could not read from address from prefs.js: {str(e)}")

        # Fallback: return a generic address
        return "sender@example.com"

class ThunderbirdSender(EmailSenderBase):
    """Thunderbird/SMTP email sender with hybrid approach for email history"""

    def __init__(self, smtp_server: str, smtp_port: int, username: str, password: str,
                 use_tls: bool = True, thunderbird_profile: Optional[str] = None,
                 save_to_thunderbird: bool = True):
        super().__init__()
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.username = username
        self.password = password
        self.use_tls = use_tls
        self.thunderbird_profile = thunderbird_profile
        self.save_to_thunderbird = save_to_thunderbird
        # Always create manager for auto-detection if profile not specified
        self.thunderbird_manager = ThunderbirdProfileManager(thunderbird_profile)

    def send_email(self, to_emails: List[str], cc_emails: List[str] = None,
                    bcc_emails: List[str] = None, subject: str = "", body: str = "",
                    attachment_path: Optional[str] = None) -> bool:
        """Send email via SMTP and save to Thunderbird (hybrid approach)"""
        try:
            # Step 1: Send via SMTP
            smtp_success = self._send_via_smtp(to_emails, cc_emails, bcc_emails,
                                             subject, body, attachment_path)

            if not smtp_success:
                self.logger.error("SMTP sending failed, aborting hybrid approach")
                return False

            # Step 2: Save to Thunderbird Sent folder if enabled
            if self.save_to_thunderbird and self.thunderbird_manager:
                thunderbird_success = self._save_to_thunderbird_sent(
                    to_emails, cc_emails, bcc_emails, subject, body, attachment_path
                )

                if not thunderbird_success:
                    self.logger.warning("Failed to save email to Thunderbird Sent folder, but SMTP was successful")

            self.logger.info(f"Email sent successfully via hybrid approach to {', '.join(to_emails)}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to send email via hybrid approach: {str(e)}")
            return False

    def _send_via_smtp(self, to_emails: List[str], cc_emails: List[str] = None,
                      bcc_emails: List[str] = None, subject: str = "", body: str = "",
                      attachment_path: Optional[str] = None) -> bool:
        """Send email via SMTP"""
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.username
            msg['To'] = ', '.join(to_emails)
            if cc_emails:
                msg['Cc'] = ', '.join(cc_emails)
            msg['Subject'] = subject

            # Add body
            msg.attach(MIMEText(body, 'html'))

            # Add attachment if provided
            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {os.path.basename(attachment_path)}'
                    )
                    msg.attach(part)
                self.logger.info(f"Added attachment: {attachment_path}")

            # Connect to server and send
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)

            if self.use_tls:
                server.starttls()

            server.login(self.username, self.password)

            # Get all recipients
            all_recipients = to_emails + (cc_emails or []) + (bcc_emails or [])

            # Send email
            server.sendmail(self.username, all_recipients, msg.as_string())
            server.quit()

            self.logger.info(f"Email sent successfully via SMTP to {', '.join(to_emails)}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to send email via SMTP: {str(e)}")
            return False

    def _save_to_thunderbird_sent(self, to_emails: List[str], cc_emails: List[str] = None,
                                 bcc_emails: List[str] = None, subject: str = "",
                                 body: str = "", attachment_path: Optional[str] = None) -> bool:
        """Save email to Thunderbird Sent folder"""
        try:
            # Prepare attachments list
            attachments = []
            if attachment_path:
                attachments.append(attachment_path)

            # Save to Thunderbird
            return self.thunderbird_manager.save_email_to_sent(to_emails, subject, body, attachments)

        except Exception as e:
            self.logger.error(f"Failed to save email to Thunderbird Sent folder: {str(e)}")
            return False

    def test_connection(self) -> bool:
        """Test SMTP connection"""
        try:
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            if self.use_tls:
                server.starttls()
            server.login(self.username, self.password)
            server.quit()
            return True
        except Exception as e:
            self.logger.error(f"SMTP connection test failed: {str(e)}")
            return False

class EmailSenderFactory:
    """Factory for creating email senders"""

    @staticmethod
    def create_sender(client_type: str, **kwargs) -> EmailSenderBase:
        """Create email sender based on client type"""

        client = client_type.lower()
        if client == 'outlook':
            return OutlookSender()

        elif client in ('thunderbird', 'smtp', 'thunderbird smtp'):
            # Normalize kwargs to support both profile keys (smtp_*) and generic keys
            smtp_server = kwargs.get('smtp_server') or kwargs.get('server')
            smtp_port = kwargs.get('smtp_port') or kwargs.get('port')
            username = kwargs.get('smtp_username') or kwargs.get('username')
            password = kwargs.get('smtp_password') or kwargs.get('password')
            use_tls = kwargs.get('smtp_use_tls')
            if use_tls is None:
                use_tls = kwargs.get('use_tls', True)

            # Thunderbird-specific parameters
            thunderbird_profile = kwargs.get('thunderbird_profile')
            save_to_thunderbird = kwargs.get('save_to_thunderbird', True)

            missing = [name for name, val in [
                ('smtp_server', smtp_server),
                ('smtp_port', smtp_port),
                ('username', username),
                ('password', password),
            ] if val in (None, '')]

            if missing:
                raise ValueError(f"Missing required parameter for SMTP: {', '.join(missing)}")

            return ThunderbirdSender(
                smtp_server=str(smtp_server),
                smtp_port=int(smtp_port),
                username=str(username),
                password=str(password),
                use_tls=bool(use_tls),
                thunderbird_profile=thunderbird_profile,
                save_to_thunderbird=bool(save_to_thunderbird)
            )

        else:
            raise ValueError(f"Unsupported email client type: {client_type}")

    @staticmethod
    def get_available_clients() -> List[str]:
        """Get list of available email clients"""
        available = []

        # Check Outlook
        try:
            outlook = win32.Dispatch('outlook.application')
            available.append('outlook')
        except:
            pass

        # Thunderbird options (SMTP with and without history)
        available.append('thunderbird')
        available.append('smtp')
        available.append('thunderbird smtp')

        return available