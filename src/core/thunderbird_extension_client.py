import asyncio
import json
import websockets
import threading
import time
import logging
import socket
from typing import Dict, List, Optional, Any
import os
import tempfile
import base64

class ThunderbirdExtensionClient:
    """Client for communicating with Thunderbird WebExtension"""

    def __init__(self, host: str = "localhost", port: int = 8765):
        self.host = host
        self.port = port
        self.websocket = None
        self.logger = logging.getLogger(__name__)
        self.connected = False
        self.message_handlers = {}
        self.pending_requests = {}
        self.request_id = 0
        self.server_thread = None
        self.server = None
        self.loop = None

    async def connect(self) -> bool:
        """Connect to Thunderbird WebExtension websocket server"""
        try:
            uri = f"ws://{self.host}:{self.port}"
            self.logger.info(f"Attempting to connect to Thunderbird WebExtension at {uri}")

            # Check if WebSocket server is running first
            try:
                self.websocket = await websockets.connect(uri, timeout=5.0)
                self.logger.info(f"Successfully connected to WebSocket at {uri}")
            except websockets.exceptions.ConnectionClosed as e:
                self.logger.error(f"WebSocket connection closed immediately: {str(e)}")
                return False
            except asyncio.TimeoutError:
                self.logger.error(f"WebSocket connection timeout - server may not be running on {uri}")
                return False
            except Exception as conn_e:
                self.logger.error(f"WebSocket connection failed: {str(conn_e)}")
                self.logger.info("This usually means:")
                self.logger.info("1. Thunderbird WebExtension is not installed")
                self.logger.info("2. WebSocket server is not running")
                self.logger.info("3. Port 8765 is blocked or in use")
                self.logger.info("4. Extension is not loaded in Thunderbird")
                return False

            # Start message handler
            asyncio.create_task(self._message_handler())

            self.connected = True
            self.logger.info(f"Connected to Thunderbird WebExtension at {uri}")
            self.logger.info("WebExtension connection established successfully")
            return True

        except Exception as e:
            self.logger.error(f"Failed to connect to Thunderbird WebExtension: {str(e)}")
            self.logger.error("DIAGNOSTIC: Thunderbird WebExtension not connected")
            self.logger.info("Please check:")
            self.logger.info("1. Is the Thunderbird WebExtension installed?")
            self.logger.info("2. Is Thunderbird running with the extension loaded?")
            self.logger.info("3. Is the WebSocket server running on port 8765?")
            self.logger.info("4. Are there any firewall blocking port 8765?")
            return False

    async def disconnect(self):
        """Disconnect from Thunderbird WebExtension"""
        if self.websocket:
            await self.websocket.close()
            self.connected = False
            self.logger.info("Disconnected from Thunderbird WebExtension")

    def is_port_available(self, port: int) -> bool:
        """Check if a port is available for binding"""
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                result = sock.bind(('', port))
                sock.close()
                return True
        except OSError:
            return False

    def find_available_port(self, start_port: int = 8765, max_attempts: int = 10) -> Optional[int]:
        """Find an available port starting from start_port"""
        for port in range(start_port, start_port + max_attempts):
            if self.is_port_available(port):
                return port
        return None

    def stop_websocket_server(self):
        """Stop the websocket server"""
        try:
            if self.loop and not self.loop.is_closed():
                # Schedule the server stop in the event loop
                if self.server:
                    asyncio.run_coroutine_threadsafe(self._stop_server_async(), self.loop)

                # Stop the event loop
                if self.loop.is_running():
                    self.loop.call_soon_threadsafe(self.loop.stop)

                # Wait a bit for cleanup
                time.sleep(0.1)

                # Close the loop
                if not self.loop.is_closed():
                    self.loop.close()

            # Clean up thread
            if self.server_thread and self.server_thread.is_alive():
                self.server_thread.join(timeout=2.0)

            self.logger.info("WebSocket server stopped successfully")
        except Exception as e:
            self.logger.error(f"Error stopping WebSocket server: {str(e)}")

    async def _stop_server_async(self):
        """Async method to stop the server"""
        if self.server:
            self.server.close()
            await self.server.wait_closed()
            self.server = None

    async def _message_handler(self):
        """Handle incoming messages from WebExtension"""
        try:
            async for message in self.websocket:
                try:
                    data = json.loads(message)
                    await self._process_message(data)
                except json.JSONDecodeError as e:
                    self.logger.error(f"Failed to decode message: {str(e)}")
                except Exception as e:
                    self.logger.error(f"Error processing message: {str(e)}")
        except Exception as e:
            self.logger.error(f"Message handler error: {str(e)}")
            self.connected = False

    async def _process_message(self, data: Dict[str, Any]):
        """Process incoming message"""
        message_type = data.get('type')
        request_id = data.get('requestId')

        if request_id and request_id in self.pending_requests:
            if message_type == 'error':
                self.pending_requests[request_id]['error'] = data.get('error', 'Unknown error')
            else:
                self.pending_requests[request_id]['result'] = data

            self.pending_requests[request_id]['event'].set()

        elif message_type == 'ping':
            # Respond to ping
            await self._send_message({
                'type': 'pong',
                'timestamp': data.get('timestamp')
            })

        elif message_type in ['sendSuccess', 'sendError']:
            # These are secondary notifications from onAfterSend, we can log them.
            # The primary response is handled via the requestId.
            self.logger.info(f"Received notification from extension: {message_type}")

        else:
            self.logger.warning(f"Unknown message type or no matching request_id: {message_type}")

    async def _send_message(self, message: Dict[str, Any]) -> bool:
        """Send message to WebExtension"""
        if not self.connected or not self.websocket:
            self.logger.error("Not connected to WebExtension")
            return False

        try:
            json_message = json.dumps(message)
            await self.websocket.send(json_message)
            return True
        except Exception as e:
            self.logger.error(f"Failed to send message: {str(e)}")
            return False

    async def _send_request_and_wait(self, message: Dict[str, Any], timeout: float = 30.0) -> Dict[str, Any]:
        """Helper to send a request and wait for a response."""
        if not self.is_connected():
            raise Exception("Not connected to WebExtension")

        self.request_id += 1
        request_id = self.request_id
        message['requestId'] = request_id

        # Create pending request
        import asyncio
        event = asyncio.Event()
        self.pending_requests[request_id] = {
            'event': event,
            'result': None,
            'error': None
        }

        # Send message
        success = await self._send_message(message)
        if not success:
            del self.pending_requests[request_id]
            raise Exception("Failed to send message to WebExtension")

        # Wait for response
        try:
            await asyncio.wait_for(event.wait(), timeout=timeout)
        except asyncio.TimeoutError:
            # Clean up pending request
            if request_id in self.pending_requests:
                del self.pending_requests[request_id]
            raise Exception(f"Timeout waiting for response to {message.get('type')}")

        # Get result
        result = self.pending_requests.pop(request_id)
        if result['error']:
            raise Exception(f"WebExtension error for {message.get('type')}: {result['error']}")

        return result.get('result', {})

    async def send_email(self, email_data: Dict[str, Any]) -> Dict[str, Any]:
        """Send email via Thunderbird WebExtension"""
        message = {
            'type': 'sendEmail',
            'emailData': email_data
        }
        return await self._send_request_and_wait(message, timeout=60.0)

    async def check_availability(self) -> bool:
        """Check if WebExtension is available and responsive."""
        try:
            message = {'type': 'checkAvailability'}
            result = await self._send_request_and_wait(message, timeout=10.0)
            return result.get('available', False)
        except Exception as e:
            self.logger.warning(f"Availability check failed: {e}")
            return False

    async def get_accounts(self) -> List[Dict[str, Any]]:
        """Get Thunderbird accounts"""
        try:
            message = {'type': 'getAccounts'}
            result = await self._send_request_and_wait(message)
            return result.get('accounts', [])
        except Exception as e:
            self.logger.error(f"Error getting accounts: {e}")
            return []

    def start_websocket_server(self):
        """Start websocket server for WebExtension communication"""
        try:
            # Check if port is available, if not find an alternative
            actual_port = self.port
            if not self.is_port_available(self.port):
                available_port = self.find_available_port(self.port, 10)
                if available_port:
                    actual_port = available_port
                    self.logger.info(f"Port {self.port} is in use, using alternative port {actual_port}")
                else:
                    raise Exception(f"No available ports found starting from {self.port}")

            # Update port if changed
            if actual_port != self.port:
                self.port = actual_port

            def run_server():
                async def handler(websocket, path):
                    self.websocket = websocket
                    self.connected = True

                    try:
                        async for message in websocket:
                            try:
                                data = json.loads(message)
                                await self._process_message(data)
                            except json.JSONDecodeError as e:
                                self.logger.error(f"Failed to decode message: {str(e)}")
                            except Exception as e:
                                self.logger.error(f"Error processing message: {str(e)}")
                    except Exception as e:
                        self.logger.error(f"Websocket handler error: {str(e)}")
                    finally:
                        self.connected = False

                # Create new event loop for this thread
                self.loop = asyncio.new_event_loop()
                asyncio.set_event_loop(self.loop)

                try:
                    # Create server
                    self.server = websockets.serve(handler, self.host, self.port)

                    # Start server
                    start_server = self.loop.run_until_complete(self.server)
                    self.logger.info(f"WebSocket server started on {self.host}:{self.port}")

                    # Run forever until stopped
                    self.loop.run_forever()

                except Exception as e:
                    self.logger.error(f"Error in WebSocket server: {str(e)}")
                finally:
                    # Cleanup
                    try:
                        if self.server:
                            self.server.close()
                            self.loop.run_until_complete(self.server.wait_closed())
                    except:
                        pass

                    try:
                        self.loop.close()
                    except:
                        pass

            # Start server thread
            self.server_thread = threading.Thread(target=run_server, daemon=True)
            self.server_thread.start()

            # Wait a bit for server to start
            time.sleep(0.1)

            self.logger.info(f"WebSocket server started on {self.host}:{self.port}")

        except Exception as e:
            self.logger.error(f"Failed to start WebSocket server: {str(e)}")
            raise

    def is_connected(self) -> bool:
        """Check if connected to WebExtension"""
        return self.connected and self.websocket is not None

    def __del__(self):
        """Destructor to ensure proper cleanup"""
        try:
            self.stop_websocket_server()
        except:
            pass

class ThunderbirdWebExtensionSender:
    """Thunderbird email sender using WebExtension"""

    def __init__(self, extension_client: ThunderbirdExtensionClient):
        self.extension_client = extension_client
        self.logger = logging.getLogger(__name__)

    async def send_email(self, to_emails: List[str], cc_emails: List[str] = None,
                        bcc_emails: List[str] = None, subject: str = "", body: str = "",
                        attachment_path: Optional[str] = None) -> bool:
        """Send email via Thunderbird WebExtension"""
        try:
            # Prepare email data
            email_data = {
                'to': to_emails,
                'cc': cc_emails or [],
                'bcc': bcc_emails or [],
                'subject': subject,
                'body': body,
                'attachmentPath': attachment_path,
                'attachmentName': os.path.basename(attachment_path) if attachment_path else None
            }

            # Send email
            result = await self.extension_client.send_email(email_data)

            if result.get('success'):
                self.logger.info(f"Email sent successfully via WebExtension: {result.get('messageId')}")
                return True
            else:
                self.logger.error(f"Failed to send email via WebExtension: {result.get('error')}")
                return False

        except Exception as e:
            self.logger.error(f"Error sending email via WebExtension: {str(e)}")
            return False

    async def check_availability(self) -> bool:
        """Check if WebExtension is available"""
        return await self.extension_client.check_availability()

    async def get_accounts(self) -> List[Dict[str, Any]]:
        """Get Thunderbird accounts"""
        return await self.extension_client.get_accounts()

    def is_available(self) -> bool:
        """Check if WebExtension is available (synchronous)"""
        return self.extension_client.is_connected()