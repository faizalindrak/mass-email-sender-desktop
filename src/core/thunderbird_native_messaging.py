import sys
import json
import struct
import logging
import asyncio
import os
import threading
import time
from typing import Dict, List, Optional, Any

class ThunderbirdNativeMessagingHost:
    """Native messaging host for Thunderbird WebExtension communication"""

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.message_handlers = {}
        self.pending_requests = {}
        self.request_id = 0
        self.running = True
        
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )

    def read_message(self):
        """Read a message from stdin (native messaging protocol)"""
        try:
            # Read the 4-byte message length
            raw_length = sys.stdin.buffer.read(4)
            if not raw_length:
                return None
            
            # Unpack the length to an integer
            message_length = struct.unpack('=I', raw_length)[0]
            
            # Read the message content
            message = sys.stdin.buffer.read(message_length).decode('utf-8')
            
            # Parse as JSON
            return json.loads(message)
            
        except Exception as e:
            self.logger.error(f"Error reading message: {str(e)}")
            return None

    def send_message(self, message: Dict[str, Any]):
        """Send a message to stdout (native messaging protocol)"""
        try:
            # Encode the message as JSON
            encoded_message = json.dumps(message).encode('utf-8')
            
            # Pack the message length
            encoded_length = struct.pack('=I', len(encoded_message))
            
            # Write to stdout
            sys.stdout.buffer.write(encoded_length)
            sys.stdout.buffer.write(encoded_message)
            sys.stdout.buffer.flush()
            
        except Exception as e:
            self.logger.error(f"Error sending message: {str(e)}")

    def handle_send_email(self, message: Dict[str, Any]):
        """Handle send email request"""
        try:
            email_data = message.get('emailData', {})
            request_id = message.get('requestId')
            
            # Log the email data for debugging
            self.logger.info(f"Received send email request: {email_data}")
            
            # For now, we'll simulate email sending
            # In a real implementation, you would integrate with an email client
            success = True
            message_id = f"msg_{int(time.time())}"
            
            response = {
                'type': 'emailSent',
                'requestId': request_id,
                'success': success,
                'messageId': message_id
            }
            
            if not success:
                response['error'] = 'Failed to send email'
            
            self.send_message(response)
            
        except Exception as e:
            self.logger.error(f"Error handling send email: {str(e)}")
            self.send_message({
                'type': 'emailSent',
                'requestId': message.get('requestId'),
                'success': False,
                'error': str(e)
            })

    def handle_check_availability(self, message: Dict[str, Any]):
        """Handle availability check"""
        try:
            response = {
                'type': 'availability',
                'requestId': message.get('requestId'),
                'available': True
            }
            self.send_message(response)
            
        except Exception as e:
            self.logger.error(f"Error handling availability check: {str(e)}")

    def handle_get_accounts(self, message: Dict[str, Any]):
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
            
            response = {
                'type': 'accounts',
                'requestId': message.get('requestId'),
                'accounts': accounts
            }
            self.send_message(response)
            
        except Exception as e:
            self.logger.error(f"Error handling get accounts: {str(e)}")

    def process_message(self, message: Dict[str, Any]):
        """Process incoming message"""
        try:
            message_type = message.get('type')
            
            if message_type == 'sendEmail':
                self.handle_send_email(message)
            elif message_type == 'checkAvailability':
                self.handle_check_availability(message)
            elif message_type == 'getAccounts':
                self.handle_get_accounts(message)
            elif message_type == 'ping':
                # Respond to ping
                self.send_message({
                    'type': 'pong',
                    'timestamp': message.get('timestamp'),
                    'requestId': message.get('requestId')
                })
            else:
                self.logger.warning(f"Unknown message type: {message_type}")
                
        except Exception as e:
            self.logger.error(f"Error processing message: {str(e)}")

    def run(self):
        """Main loop for the native messaging host"""
        self.logger.info("Thunderbird Native Messaging Host started")
        
        try:
            while self.running:
                message = self.read_message()
                if message is None:
                    break
                
                self.process_message(message)
                
        except KeyboardInterrupt:
            self.logger.info("Received interrupt signal")
        except Exception as e:
            self.logger.error(f"Error in main loop: {str(e)}")
        finally:
            self.logger.info("Thunderbird Native Messaging Host stopped")

class ThunderbirdNativeClient:
    """Client for communicating with Thunderbird via native messaging"""

    def __init__(self, host_path: Optional[str] = None):
        self.host_path = host_path
        self.logger = logging.getLogger(__name__)
        self.connected = False
        self.process = None

    def connect(self) -> bool:
        """Connect to Thunderbird via native messaging"""
        try:
            import subprocess
            
            # Find the native messaging host executable
            if not self.host_path:
                # Use the current Python interpreter to run the native messaging host
                self.host_path = sys.executable
            
            # Get the path to the native messaging host script
            script_path = os.path.join(os.path.dirname(__file__), 'thunderbird_native_messaging.py')
            
            # Start the native messaging host
            self.process = subprocess.Popen(
                [self.host_path, script_path],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=False  # Use binary mode for proper message handling
            )
            
            self.connected = True
            self.logger.info("Connected to Thunderbird Native Messaging Host")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to connect to native messaging host: {str(e)}")
            return False

    def disconnect(self):
        """Disconnect from native messaging host"""
        if self.process:
            self.process.terminate()
            self.process = None
        self.connected = False
        self.logger.info("Disconnected from Thunderbird Native Messaging Host")

    def send_message(self, message: Dict[str, Any]) -> bool:
        """Send message to native messaging host"""
        if not self.connected or not self.process:
            self.logger.error("Not connected to native messaging host")
            return False

        try:
            # Encode the message as JSON
            encoded_message = json.dumps(message).encode('utf-8')
            
            # Pack the message length
            encoded_length = struct.pack('=I', len(encoded_message))
            
            # Write to process stdin
            self.process.stdin.write(encoded_length)
            self.process.stdin.write(encoded_message)
            self.process.stdin.flush()
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error sending message: {str(e)}")
            return False

    def read_message(self) -> Optional[Dict[str, Any]]:
        """Read message from native messaging host"""
        if not self.connected or not self.process:
            return None

        try:
            # Read the 4-byte message length
            raw_length = self.process.stdout.read(4)
            if not raw_length:
                return None
            
            # Unpack the length to an integer
            message_length = struct.unpack('=I', raw_length)[0]
            
            # Read the message content
            message = self.process.stdout.read(message_length).decode('utf-8')
            
            # Parse as JSON
            return json.loads(message)
            
        except Exception as e:
            self.logger.error(f"Error reading message: {str(e)}")
            return None

    def send_email(self, email_data: Dict[str, Any]) -> Dict[str, Any]:
        """Send email via native messaging"""
        self.request_id += 1
        request_id = self.request_id

        # Prepare message
        message = {
            'type': 'sendEmail',
            'requestId': request_id,
            'emailData': email_data
        }

        # Send message
        if not self.send_message(message):
            raise Exception("Failed to send message to native messaging host")

        # Wait for response (with timeout)
        import time
        start_time = time.time()
        timeout = 30  # 30 seconds timeout

        while time.time() - start_time < timeout:
            response = self.read_message()
            if response and response.get('requestId') == request_id:
                return response
            
            time.sleep(0.1)  # Small delay to prevent busy waiting

        raise Exception("Timeout waiting for native messaging response")

    def check_availability(self) -> bool:
        """Check if native messaging is available"""
        try:
            message = {
                'type': 'checkAvailability',
                'requestId': 1
            }

            if not self.send_message(message):
                return False

            # Wait for response
            import time
            start_time = time.time()
            timeout = 10  # 10 seconds timeout

            while time.time() - start_time < timeout:
                response = self.read_message()
                if response and response.get('type') == 'availability':
                    return response.get('available', False)
                
                time.sleep(0.1)

            return False

        except Exception as e:
            self.logger.error(f"Error checking availability: {str(e)}")
            return False

    def get_accounts(self) -> List[Dict[str, Any]]:
        """Get Thunderbird accounts"""
        try:
            message = {
                'type': 'getAccounts',
                'requestId': 2
            }

            if not self.send_message(message):
                return []

            # Wait for response
            import time
            start_time = time.time()
            timeout = 10  # 10 seconds timeout

            while time.time() - start_time < timeout:
                response = self.read_message()
                if response and response.get('type') == 'accounts':
                    return response.get('accounts', [])
                
                time.sleep(0.1)

            return []

        except Exception as e:
            self.logger.error(f"Error getting accounts: {str(e)}")
            return []

    def is_connected(self) -> bool:
        """Check if connected to native messaging host"""
        return self.connected and self.process is not None

if __name__ == '__main__':
    # Run the native messaging host when script is executed directly
    host = ThunderbirdNativeMessagingHost()
    host.run()