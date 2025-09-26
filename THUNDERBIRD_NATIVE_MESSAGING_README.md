# Thunderbird Native Messaging Integration

This document describes the native messaging implementation for Thunderbird WebExtension integration, replacing the previous WebSocket-based approach.

## Overview

The native messaging approach provides a more reliable and secure communication channel between the Python application and Thunderbird WebExtension compared to WebSocket connections.

## Architecture

```
Python Application (Email Automation)
        |
        | Native Messaging Protocol
        |
Thunderbird WebExtension (background.js)
        |
        | Thunderbird Compose API
        |
Thunderbird Email Client
```

## Components

### 1. Native Messaging Host (`src/core/thunderbird_native_client.py`)

- **ThunderbirdNativeHost**: Server-side component that handles messages from Thunderbird
- **ThunderbirdNativeClient**: Client-side component that connects to the native messaging host
- Uses the `nativemessaging-ng` package for proper protocol handling

### 2. Thunderbird Extension (`thunderbird_extension/`)

- **background.js**: Updated to use `browser.runtime.connectNative()` instead of WebSocket
- **manifest.json**: Added `"nativeMessaging"` permission
- **email_automation_native_host.json**: Native messaging host configuration

### 3. Configuration Utility (`src/utils/thunderbird_native_config.py`)

- Installs/uninstalls the native messaging host configuration
- Checks installation status
- Cross-platform support (Windows, macOS, Linux)

### 4. Email Sender Integration (`src/core/email_sender.py`)

- Updated `ThunderbirdSender` to use native messaging client
- `EmailSenderFactory` creates native messaging clients instead of WebSocket clients

## Installation

### 1. Install Dependencies

```bash
pip install nativemessaging-ng
```

### 2. Configure Native Messaging Host

```bash
# Install the native messaging host
python -m src.utils.thunderbird_native_config install

# Check installation status
python -m src.utils.thunderbird_native_config status

# Uninstall if needed
python -m src.utils.thunderbird_native_config uninstall
```

### 3. Install Thunderbird Extension

1. Build the extension:
   ```bash
   # The extension files are in thunderbird_extension/
   # Build XPI file using Thunderbird's extension manager or web-ext
   ```

2. Install the XPI file in Thunderbird:
   - Tools → Add-ons
   - Click the gear icon → Install Add-on From File
   - Select the built XPI file

## Configuration Files

### Native Messaging Host Configuration

The configuration file is automatically created at:
- **Windows**: `%APPDATA%\Thunderbird\NativeMessagingHosts\email_automation_native_host.json`
- **macOS**: `~/Library/Application Support/Thunderbird/NativeMessagingHosts/email_automation_native_host.json`
- **Linux**: `~/.thunderbird/native-messaging-hosts/email_automation_native_host.json`

Content:
```json
{
  "name": "email_automation_native_host",
  "description": "Email Automation Native Messaging Host",
  "path": "python",
  "type": "stdio",
  "allowed_extensions": ["email-automation@thunderbird"]
}
```

## Testing

### Test Script

Run the test script to verify the native messaging setup:

```bash
python test_thunderbird_native_messaging.py
```

### Manual Testing

1. Start Thunderbird with the extension installed
2. Run the email automation application
3. Configure a profile to use Thunderbird as email client
4. Send a test email

## Message Protocol

### From Extension to Host

```javascript
{
  "type": "sendEmail",
  "requestId": "unique_id",
  "emailData": {
    "to": ["recipient@example.com"],
    "cc": [],
    "bcc": [],
    "subject": "Email Subject",
    "body": "Email body content",
    "attachmentPath": "/path/to/attachment.pdf",
    "attachmentName": "attachment.pdf"
  }
}
```

### From Host to Extension

```javascript
{
  "type": "emailSent",
  "requestId": "unique_id",
  "success": true,
  "messageId": "msg_1234567890"
}
```

## Troubleshooting

### Common Issues

1. **"Native messaging host not found"**
   - Run: `python -m src.utils.thunderbird_native_config install`
   - Check that Thunderbird is installed and the extension is loaded

2. **"Extension not allowed"**
   - Verify the extension ID in the configuration matches the manifest
   - Restart Thunderbird after installing the extension

3. **"Python executable not found"**
   - Ensure Python is in PATH
   - Check the configuration file points to the correct Python executable

4. **"Permission denied"**
   - Check file permissions on the configuration file
   - Ensure Thunderbird has access to the native messaging directory

### Debug Logging

Enable debug logging in the Python application:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

Check Thunderbird's Browser Console (Ctrl+Shift+J) for extension errors.

### Manual Verification

1. Check native messaging host configuration:
   ```bash
   python -m src.utils.thunderbird_native_config status
   ```

2. Test extension loading in Thunderbird:
   - Open Thunderbird
   - Go to Tools → Add-ons
   - Check that the extension is enabled and has no errors

3. Test native messaging connection:
   ```bash
   python test_thunderbird_native_messaging.py
   ```

## Migration from WebSocket

The native messaging implementation replaces the previous WebSocket-based approach:

### Changes Made

1. **Removed**: WebSocket server (`ThunderbirdExtensionClient`)
2. **Added**: Native messaging host (`ThunderbirdNativeClient`)
3. **Updated**: Thunderbird extension background script
4. **Updated**: Email sender factory and Thunderbird sender
5. **Added**: Configuration utility for native messaging setup

### Benefits

- **More Reliable**: No port conflicts or firewall issues
- **Secure**: Uses Thunderbird's built-in native messaging security
- **Standard**: Follows Mozilla's recommended approach for WebExtension native communication
- **Cross-platform**: Works consistently across Windows, macOS, and Linux

## API Reference

### ThunderbirdNativeClient

```python
client = ThunderbirdNativeClient()

# Connect to native messaging host
await client.connect()

# Send email
result = await client.send_email({
    'to': ['recipient@example.com'],
    'subject': 'Test',
    'body': 'Test body'
})

# Check availability
available = await client.check_availability()

# Get accounts
accounts = await client.get_accounts()

# Disconnect
await client.disconnect()
```

### ThunderbirdNativeConfig

```python
config = ThunderbirdNativeConfig()

# Install native messaging host
config.install_native_host()

# Check installation
config.check_installation()

# Print status
config.print_status()

# Uninstall
config.uninstall_native_host()
```

## Security Considerations

- The native messaging host configuration specifies allowed extensions
- Communication uses JSON over stdin/stdout
- No network communication required
- Extension must be explicitly allowed in the configuration

## Future Improvements

1. **Error Recovery**: Automatic reconnection on connection loss
2. **Message Queue**: Handle multiple concurrent email requests
3. **Progress Updates**: Send progress notifications during email sending
4. **Account Selection**: Allow user to choose which Thunderbird account to use
5. **Draft Support**: Save emails as drafts before sending