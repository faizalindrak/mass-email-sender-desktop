import win32com.client as win32
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import logging
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

            # Add attachment if provided
            if attachment_path and os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
                self.logger.info(f"Added attachment: {attachment_path}")

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

class ThunderbirdSender(EmailSenderBase):
    """Thunderbird/SMTP email sender"""

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

class EmailSenderFactory:
    """Factory for creating email senders"""

    @staticmethod
    def create_sender(client_type: str, **kwargs) -> EmailSenderBase:
        """Create email sender based on client type"""

        if client_type.lower() == 'outlook':
            return OutlookSender()

        elif client_type.lower() == 'thunderbird' or client_type.lower() == 'smtp':
            required_params = ['smtp_server', 'smtp_port', 'username', 'password']
            for param in required_params:
                if param not in kwargs:
                    raise ValueError(f"Missing required parameter for SMTP: {param}")

            return ThunderbirdSender(
                smtp_server=kwargs['smtp_server'],
                smtp_port=kwargs['smtp_port'],
                username=kwargs['username'],
                password=kwargs['password'],
                use_tls=kwargs.get('use_tls', True)
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

        # SMTP is always available
        available.append('smtp')

        return available