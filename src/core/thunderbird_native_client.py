import json
import logging
import os
import sys
from typing import Dict, List, Optional, Any
from nativemessaging import nativemessaging

class ThunderbirdNativeHost:
    """Native messaging host for Thunderbird WebExtension using nativemessaging-ng"""

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.request_id = 0

    def handle_message(self, message: Dict[str, Any]) -> Dict[str, Any]:
        """Handle incoming messages from Thunderbird extension"""
        try:
            message_type = message.get('type')
            request_id = message.get('requestId')

            self.logger.info(f"Received message: {message_type}")

            if message_type == 'sendEmail':
                return self.handle_send_email(message, request_id)
            elif message_type == 'checkAvailability':
                return self.handle_check_availability(request_id)
            elif message_type == 'getAccounts':
                return self.handle_get_accounts(request_id)
            elif message_type == 'ping':
                return {'type': 'pong', 'timestamp': message.get('timestamp'), 'requestId': request_id}
            else:
                self.logger.warning(f"Unknown message type: {message_type}")
                return {'type': 'error', 'error': f'Unknown message type: {message_type}', 'requestId': request_id}

        except Exception as e:
            self.logger.error(f"Error handling message: {str(e)}")
            return {'type': 'error', 'error': str(e), 'requestId': message.get('requestId')}

    def handle_send_email(self, message: Dict[str, Any], request_id: int) -> Dict[str, Any]:
        """Handle send email request"""
        try:
            email_data = message.get('emailData', {})
            self.logger.info(f"Processing send email request: {email_data}")

            # Extract email parameters
            to_emails = email_data.get('to', [])
            cc_emails = email_data.get('cc', [])
            bcc_emails = email_data.get('bcc', [])
            subject = email_data.get('subject', '')
            body = email_data.get('body', '')
            attachment_path = email_data.get('attachmentPath')

            # Send the email using Thunderbird compose API
            success = self._send_via_thunderbird_compose(
                to_emails, cc_emails, bcc_emails, subject, body, attachment_path
            )

            if success:
                return {
                    'type': 'emailSent',
                    'requestId': request_id,
                    'success': True,
                    'messageId': f"msg_{request_id}_{os.getpid()}"
                }
            else:
                return {
                    'type': 'emailSent',
                    'requestId': request_id,
                    'success': False,
                    'error': 'Failed to send email via Thunderbird'
                }

        except Exception as e:
            self.logger.error(f"Error in handle_send_email: {str(e)}")
            return {
                'type': 'emailSent',
                'requestId': request_id,
                'success': False,
                'error': str(e)
            }

    def _send_via_thunderbird_compose(self, to_emails: List[str], cc_emails: List[str],
                                      bcc_emails: List[str], subject: str, body: str,
                                      attachment_path: Optional[str] = None) -> bool:
        """Send email using Thunderbird compose API"""
        try:
            # This would be implemented using Thunderbird's compose API
            # For now, we'll simulate the process
            self.logger.info(f"Simulating email send to: {to_emails}")

            # In a real implementation, this would:
            # 1. Create a compose window
            # 2. Set recipients, subject, body
            # 3. Add attachment if provided
            # 4. Send the message

            return True

        except Exception as e:
            self.logger.error(f"Error sending via Thunderbird compose: {str(e)}")
            return False

    def handle_check_availability(self, request_id: int) -> Dict[str, Any]:
        """Handle availability check"""
        try:
            # Check if Thunderbird is available
            available = self._check_thunderbird_available()

            return {
                'type': 'availability',
                'requestId': request_id,
                'available': available
            }

        except Exception as e:
            self.logger.error(f"Error checking availability: {str(e)}")
            return {
                'type': 'availability',
                'requestId': request_id,
                'available': False
            }

    def _check_thunderbird_available(self) -> bool:
        """Check if Thunderbird is available"""
        try:
            # Check if Thunderbird process is running
            import psutil
            for proc in psutil.process_iter(['name']):
                if 'thunderbird' in proc.info['name'].lower():
                    return True
            return False

        except ImportError:
            # If psutil is not available, assume Thunderbird is available
            return True
        except Exception:
            return False

    def handle_get_accounts(self, request_id: int) -> Dict[str, Any]:
        """Handle get accounts request"""
        try:
            # Return mock accounts for now
            accounts = [
                {
                    'id': 'account1',
                    'name': 'Default Account',
                    'type': 'imap',
                    'identities': [
                        {
                            'id': 'id1',
                            'name': 'User Name',
                            'email': 'user@example.com'
                        }
                    ]
                }
            ]

            return {
                'type': 'accounts',
                'requestId': request_id,
                'accounts': accounts
            }

        except Exception as e:
            self.logger.error(f"Error getting accounts: {str(e)}")
            return {
                'type': 'accounts',
                'requestId': request_id,
                'accounts': []
            }

    def run(self):
        """Run the native messaging host"""
        self.logger.info("Thunderbird native messaging host started")

        try:
            while True:
                try:
                    # Get message from extension
                    message = nativemessaging.get_message()

                    # Handle the message
                    response = self.handle_message(message)

                    # Send response back
                    nativemessaging.send_message(response)

                except KeyboardInterrupt:
                    self.logger.info("Received keyboard interrupt, shutting down")
                    break
                except Exception as e:
                    self.logger.error(f"Error processing message: {str(e)}")
                    # Send error response
                    try:
                        nativemessaging.send_message({
                            'type': 'error',
                            'error': str(e)
                        })
                    except:
                        pass

        except Exception as e:
            self.logger.error(f"Fatal error in native messaging host: {str(e)}")

class ThunderbirdNativeClient:
    """Client for communicating with Thunderbird via native messaging using nativemessaging-ng"""

    def __init__(self, extension_id: str = "email-automation@thunderbird"):
        self.extension_id = extension_id
        self.logger = logging.getLogger(__name__)
        self.connected = False
        self.process = None

    async def connect(self) -> bool:
        """Connect to Thunderbird via native messaging"""
        try:
            # For the client side, we don't actually connect to a process
            # The connection is implicit through the native messaging protocol
            # We just check if the native messaging host is configured
            self.connected = True
            self.logger.info("Native messaging client initialized")
            return True

        except Exception as e:
            self.logger.error(f"Failed to initialize native messaging client: {str(e)}")
            return False

    async def disconnect(self):
        """Disconnect from native messaging"""
        self.connected = False
        if self.process:
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
            except:
                self.process.kill()
            self.process = None
        self.logger.info("Disconnected from Thunderbird native messaging")

    async def send_message(self, message: Dict[str, Any]) -> bool:
        """Send message to Thunderbird extension"""
        if not self.connected:
            self.logger.error("Not connected to native messaging")
            return False

        try:
            # For client-side, we would need to start the native host process
            # and communicate with it. This is complex and for now we'll simulate.
            self.logger.info(f"Simulating send message: {message}")
            return True

        except Exception as e:
            self.logger.error(f"Error sending message: {str(e)}")
            return False

    async def read_message(self) -> Optional[Dict[str, Any]]:
        """Read message from Thunderbird extension"""
        if not self.connected:
            return None

        try:
            # Simulate reading a response
            self.logger.info("Simulating read message")
            return None

        except Exception as e:
            self.logger.error(f"Error reading message: {str(e)}")
            return None

    async def send_email(self, email_data: Dict[str, Any]) -> Dict[str, Any]:
        """Send email via native messaging"""
        self.request_id = getattr(self, 'request_id', 0) + 1
        request_id = self.request_id

        # For now, simulate a successful response
        self.logger.info(f"Simulating email send: {email_data}")

        return {
            'type': 'emailSent',
            'requestId': request_id,
            'success': True,
            'messageId': f"msg_{request_id}_simulated"
        }

    async def check_availability(self) -> bool:
        """Check if native messaging is available"""
        # For now, assume it's available
        return True

    async def get_accounts(self) -> List[Dict[str, Any]]:
        """Get Thunderbird accounts"""
        # Return mock accounts
        return [
            {
                'id': 'account1',
                'name': 'Default Account',
                'type': 'imap',
                'identities': [
                    {
                        'id': 'id1',
                        'name': 'User Name',
                        'email': 'user@example.com'
                    }
                ]
            }
        ]

    def is_connected(self) -> bool:
        """Check if connected to native messaging"""
        return self.connected

def run_native_host():
    """Run the native messaging host (for standalone execution)"""
    logging.basicConfig(level=logging.INFO)
    host = ThunderbirdNativeHost()
    
    # Run the host
    host.run()

if __name__ == '__main__':
    run_native_host()