import win32com.client as win32
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import logging
import time
import sys
import json
import uuid
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

class SMTPSender(EmailSenderBase):
    """SMTP email sender (generic, can be used with Thunderbird SMTP profile)"""

    def __init__(self, smtp_server: str, smtp_port: int, username: str, password: str, use_tls: bool = True):
        super().__init__()
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.username = username
        self.password = password
        self.use_tls = use_tls

    def send_email(self, to_emails: List[str], cc_emails: List[str] = None,
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

class ThunderbirdExtensionSender(EmailSenderBase):
    """Headless Thunderbird MailExtension sender using a filesystem job queue.

    Workflow:
    - Write a job JSON to {queue_dir}/jobs/{job_id}.json
    - Wait for result JSON from native host/extension at {queue_dir}/results/{job_id}.json
    Integration notes:
    - The Thunderbird MailExtension should connect to a registered native host.
    - The native host process watches the queue_dir and forwards jobs to the extension via native messaging.
    - After the extension sends the email, it posts a result back to the native host, which writes the result file.
    """

    def __init__(self, queue_dir: Optional[str] = None, timeout_seconds: float = 120.0, poll_interval: float = 0.5):
        super().__init__()
        self.timeout_seconds = float(timeout_seconds)
        self.poll_interval = float(poll_interval)
    
        # Resolution order for queue directory:
        # 1) Explicit argument queue_dir (profile 'tb_queue_dir')
        # 2) Environment variable TB_QUEUE_DIR
        # 3) OS default
        env_qdir = os.environ.get("TB_QUEUE_DIR", "").strip()
        if queue_dir and isinstance(queue_dir, str) and queue_dir.strip():
            base_queue_dir = queue_dir
        elif env_qdir:
            base_queue_dir = env_qdir
        else:
            base_queue_dir = self._default_queue_dir()
    
        # Normalize and create folders
        self.queue_dir = os.path.abspath(base_queue_dir)
        self.jobs_dir = os.path.join(self.queue_dir, "jobs")
        self.results_dir = os.path.join(self.queue_dir, "results")
    
        os.makedirs(self.jobs_dir, exist_ok=True)
        os.makedirs(self.results_dir, exist_ok=True)
    
        self.logger.info(f"ThunderbirdExtensionSender queue initialized at: {self.queue_dir}")
        self.logger.info(f"Jobs dir: {self.jobs_dir}")
        self.logger.info(f"Results dir: {self.results_dir}")

    def _default_queue_dir(self) -> str:
        """Compute a sensible default queue directory.
        
        Preference order:
        1) Project-local tb_queue (next to this repository), to keep app/host aligned without extra config
        2) OS-specific roaming path
        3) CWD tb_queue fallback
        """
        try:
            # Prefer project-local tb_queue based on repository root (src/core/ -> project root)
            project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
            local_q = os.path.join(project_root, "tb_queue")
            try:
                os.makedirs(os.path.join(local_q, "jobs"), exist_ok=True)
                os.makedirs(os.path.join(local_q, "results"), exist_ok=True)
                return local_q
            except Exception:
                # If cannot create local queue, fall back to OS path below
                pass
            
            # OS-specific roaming path
            if os.name == "nt":
                # Windows: %APPDATA%\\EmailAutomation\\tb_queue
                appdata = os.getenv("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
                return os.path.join(appdata, "EmailAutomation", "tb_queue")
            else:
                # Linux/macOS: ~/.local/share/email_automation/tb_queue (Linux) or ~/Library/Application Support/... (macOS)
                if sys.platform == "darwin":
                    base = os.path.join(os.path.expanduser("~"), "Library", "Application Support", "EmailAutomation")
                else:
                    base = os.path.join(os.path.expanduser("~"), ".local", "share", "email_automation")
                return os.path.join(base, "tb_queue")
        except Exception:
            # Fallback to project-local folder if anything goes wrong
            return os.path.abspath(os.path.join("tb_queue"))

    def send_email(
        self,
        to_emails: List[str],
        cc_emails: List[str] = None,
        bcc_emails: List[str] = None,
        subject: str = "",
        body: str = "",
        attachment_path: Optional[str] = None
    ) -> bool:
        """Queue an email send job and wait for the MailExtension result."""
        try:
            job_id = str(uuid.uuid4())

            # Normalize lists
            to_emails = to_emails or []
            cc_emails = cc_emails or []
            bcc_emails = bcc_emails or []

            # Validate/normalize attachment
            attachments = []
            if attachment_path:
                abs_attachment = os.path.abspath(attachment_path)
                if not os.path.exists(abs_attachment):
                    self.logger.error(f"Attachment file does not exist: {abs_attachment}")
                    return False
                if not os.path.isfile(abs_attachment):
                    self.logger.error(f"Attachment path is not a file: {abs_attachment}")
                    return False
                attachments.append({"path": abs_attachment})

            job = {
                "id": job_id,
                "type": "sendEmail",
                "payload": {
                    "to": to_emails,
                    "cc": cc_emails,
                    "bcc": bcc_emails,
                    "subject": subject,
                    "bodyHtml": body or "",
                    "attachments": attachments
                },
                "meta": {
                    "createdAt": int(time.time()),
                    "client": "email_automation_desktop"
                }
            }

            job_path = os.path.join(self.jobs_dir, f"{job_id}.json")
            result_path = os.path.join(self.results_dir, f"{job_id}.json")

            # Best-effort cleanup of stale result file if exists
            try:
                if os.path.exists(result_path):
                    os.remove(result_path)
            except Exception:
                pass

            # Write job file
            with open(job_path, "w", encoding="utf-8") as f:
                json.dump(job, f, ensure_ascii=False, indent=2)

            self.logger.info(f"Queued Thunderbird job: {job_path}")
            self.logger.info(f"Waiting for result: {result_path} (timeout {self.timeout_seconds}s, interval {self.poll_interval}s)")
    
            # Wait for result
            deadline = time.time() + self.timeout_seconds
            while time.time() < deadline:
                if os.path.exists(result_path):
                    try:
                        with open(result_path, "r", encoding="utf-8") as f:
                            result = json.load(f)
                        success = bool(result.get("success", False))
                        err = result.get("error")
                        if success:
                            self.logger.info("Thunderbird MailExtension reported success.")
                            return True
                        else:
                            self.logger.error(f"Thunderbird MailExtension reported failure: {err}")
                            return False
                    except Exception as e:
                        self.logger.warning(f"Result file read error, retrying: {e}")
                        # Small delay to avoid busy loop if partial write
                        time.sleep(self.poll_interval)
                        continue
    
                time.sleep(self.poll_interval)
    
            self.logger.error(f"Timed out waiting for Thunderbird MailExtension result. Queue dir: {self.queue_dir}")
            return False

        except Exception as e:
            self.logger.error(f"Failed to enqueue Thunderbird MailExtension job: {str(e)}")
            return False
class EmailSenderFactory:
    """Factory for creating email senders"""

    @staticmethod
    def create_sender(client_type: str, **kwargs) -> EmailSenderBase:
        """Create email sender based on client type"""

        client = client_type.lower()
        if client == 'outlook':
            return OutlookSender()
        
        elif client == 'thunderbird':
            # Use Thunderbird MailExtension via native messaging queue (headless, no SMTP required)
            queue_dir = kwargs.get('tb_queue_dir') or kwargs.get('queue_dir')
            return ThunderbirdExtensionSender(queue_dir=queue_dir)
        
        elif client == 'smtp':
            # Normalize kwargs to support both profile keys (smtp_*) and generic keys
            smtp_server = kwargs.get('smtp_server') or kwargs.get('server')
            smtp_port = kwargs.get('smtp_port') or kwargs.get('port')
            username = kwargs.get('smtp_username') or kwargs.get('username')
            password = kwargs.get('smtp_password') or kwargs.get('password')
            use_tls = kwargs.get('smtp_use_tls')
            if use_tls is None:
                use_tls = kwargs.get('use_tls', True)
        
            missing = [name for name, val in [
                ('smtp_server', smtp_server),
                ('smtp_port', smtp_port),
                ('username', username),
                ('password', password),
            ] if val in (None, '')]
        
            if missing:
                raise ValueError(f"Missing required parameter for SMTP: {', '.join(missing)}")
        
            return SMTPSender(
                smtp_server=str(smtp_server),
                smtp_port=int(smtp_port),
                username=str(username),
                password=str(password),
                use_tls=bool(use_tls)
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

        # Thunderbird label (SMTP) and SMTP are available
        available.append('thunderbird')
        available.append('smtp')

        return available