# Thunderbird WebExtension Integration

This document describes the Thunderbird WebExtension integration for the Email Automation Desktop application. This integration allows the application to send emails directly through Thunderbird using the Compose API, providing better integration with Thunderbird's email management features.

## Overview

The WebExtension integration consists of three main components:

1. **Thunderbird WebExtension** - A browser extension that runs within Thunderbird
2. **Python WebSocket Server** - A server that communicates with the WebExtension
3. **Email Sender Integration** - Modified email sender that uses the WebExtension instead of SMTP

## Architecture

```
┌─────────────────┐    WebSocket    ┌─────────────────┐
│   Python App    │◄──────────────►│ Thunderbird     │
│                 │   (Port 8765)   │   WebExtension  │
└─────────────────┘                └─────────────────┘
                                       │
                                       ▼
                              ┌─────────────────┐
                              │   Thunderbird   │
                              │  Compose API    │
                              └─────────────────┘
```

## Components

### 1. Thunderbird WebExtension

**Location**: `thunderbird_extension/`

**Files**:
- `manifest.json` - Extension manifest with required permissions
- `background.js` - Main extension script handling compose API
- `content.js` - Content script for additional functionality

**Key Features**:
- Uses Thunderbird's `compose.send` API to send emails
- Supports attachments via `compose.addAttachment`
- Handles compose events (beforeSend, afterSend, etc.)
- Communicates with Python app via WebSocket

**Permissions Required**:
- `compose` - Read and modify email composition
- `compose.send` - Send composed emails
- `compose.save` - Save drafts and templates
- `accountsRead` - Access account information
- `storage` - Store extension settings

### 2. Python WebSocket Server

**Location**: `src/core/thunderbird_extension_client.py`

**Classes**:
- `ThunderbirdExtensionClient` - WebSocket client for communication
- `ThunderbirdWebExtensionSender` - Email sender using WebExtension

**Features**:
- WebSocket server on port 8765
- Message handling for email sending requests
- Account information retrieval
- Error handling and status reporting

### 3. Extension Manager

**Location**: `src/utils/thunderbird_extension_manager.py`

**Features**:
- Extension building and packaging
- Installation to Thunderbird profiles
- Environment validation
- Installation troubleshooting

## Installation

### Step 1: Build the Extension

```bash
# Run the test script to build the extension
python test_thunderbird_webextension.py
```

This will create an `.xpi` file in the project root.

### Step 2: Install to Thunderbird

#### Option A: Automatic Installation
```python
from utils.thunderbird_extension_manager import ThunderbirdExtensionManager

manager = ThunderbirdExtensionManager()
success = manager.install_extension()
```

#### Option B: Manual Installation
1. Open Thunderbird
2. Go to **Tools > Add-ons**
3. Click the gear icon (⚙️) and select **"Install Add-on From File"**
4. Select the generated `.xpi` file
5. Restart Thunderbird

### Step 3: Verify Installation

1. After restart, check **Tools > Add-ons**
2. Look for "Email Automation Thunderbird Extension"
3. Ensure it's enabled
4. Check the Browser Console (Ctrl+Shift+J) for any errors

## Configuration

### Python Application Configuration

The WebExtension integration is automatically configured when you select "thunderbird" as the email client in the application.

**Configuration Options**:
- `thunderbird_profile` - Path to Thunderbird profile (auto-detected if not specified)
- `save_to_thunderbird` - Whether to save emails to Thunderbird Sent folder

### WebExtension Configuration

The WebExtension uses the following default settings:
- **WebSocket Port**: 8765
- **Extension ID**: `email-automation@thunderbird`
- **Required Thunderbird Version**: 78+

## Usage

### Basic Usage

1. Select "thunderbird" as the email client in the application
2. Configure your monitoring settings
3. Start monitoring - emails will be sent via the WebExtension

### Advanced Usage

```python
from core.thunderbird_extension_client import ThunderbirdExtensionClient, ThunderbirdWebExtensionSender

# Create client
client = ThunderbirdExtensionClient()
client.start_websocket_server()

# Create sender
sender = ThunderbirdWebExtensionSender(client)

# Send email
await sender.send_email(
    to_emails=['recipient@example.com'],
    subject='Test Email',
    body='<p>This is a test email</p>',
    attachment_path='/path/to/attachment.pdf'
)
```

## API Reference

### Compose API Functions

The WebExtension provides access to the following Thunderbird Compose API functions:

- `beginNew(details)` - Create new compose window
- `beginReply(messageId, replyType, details)` - Reply to message
- `beginForward(messageId, forwardType, details)` - Forward message
- `sendMessage(tabId, options)` - Send composed message
- `saveMessage(tabId, options)` - Save as draft or template
- `getComposeDetails(tabId)` - Get compose window details
- `setComposeDetails(tabId, details)` - Update compose window
- `addAttachment(tabId, attachment)` - Add attachment
- `listAttachments(tabId)` - List attachments
- `removeAttachment(tabId, attachmentId)` - Remove attachment
- `updateAttachment(tabId, attachmentId, attachment)` - Update attachment

### Event Listeners

The WebExtension listens for the following events:

- `onBeforeSend` - Fired before sending email
- `onAfterSend` - Fired after sending email
- `onAfterSave` - Fired after saving draft/template
- `onAttachmentAdded` - Fired when attachment is added
- `onAttachmentRemoved` - Fired when attachment is removed
- `onIdentityChanged` - Fired when sender identity changes
- `onComposeStateChanged` - Fired when compose state changes

## Troubleshooting

### Common Issues

#### 1. Extension Not Loading
**Symptoms**: Extension doesn't appear in Add-ons list
**Solutions**:
- Check Thunderbird version (requires 78+)
- Verify .xpi file is not corrupted
- Check Browser Console for errors
- Try restarting Thunderbird

#### 2. WebSocket Connection Failed
**Symptoms**: Python app can't connect to WebExtension
**Solutions**:
- Ensure WebSocket server is running on port 8765
- Check firewall settings
- Verify both Python app and Thunderbird are on same machine
- Check for port conflicts

#### 3. Emails Not Sending
**Symptoms**: Send operation fails
**Solutions**:
- Verify Thunderbird account settings
- Check compose API permissions
- Ensure email addresses are valid
- Check attachment file paths and permissions

#### 4. Attachments Not Working
**Symptoms**: Attachments not added to emails
**Solutions**:
- Check file permissions on attachment files
- Verify file paths are correct
- Ensure files are not too large
- Check attachment API usage in background.js

### Debug Information

#### Enable Debug Logging

1. Open Thunderbird
2. Go to **Tools > Developer Tools > Browser Console**
3. Check for extension-related messages
4. Look for WebSocket connection messages

#### Python Debug Logging

```python
import logging
logging.basicConfig(level=logging.DEBUG)

# Enable debug logging for extension client
logger = logging.getLogger('core.thunderbird_extension_client')
logger.setLevel(logging.DEBUG)
```

### Environment Validation

Run the validation script to check your environment:

```python
from utils.thunderbird_extension_manager import ThunderbirdExtensionManager

manager = ThunderbirdExtensionManager()
validation = manager.validate_environment()

print("Validation Results:")
for key, value in validation.items():
    if key == 'issues':
        print(f"Issues: {value}")
    else:
        print(f"{key}: {value}")
```

## Security Considerations

1. **WebSocket Communication**: Uses local WebSocket connection only
2. **File Access**: Extension only accesses files specified by Python app
3. **Network Access**: No external network access required
4. **Permissions**: Minimal required permissions for functionality

## Development

### Building from Source

```bash
# Modify extension files in thunderbird_extension/
# Run test script to build
python test_thunderbird_webextension.py
```

### Testing

```bash
# Run comprehensive test suite
python test_thunderbird_webextension.py

# Test specific components
python -c "from utils.thunderbird_extension_manager import ThunderbirdExtensionManager; m = ThunderbirdExtensionManager(); print(m.validate_environment())"
```

### Extension Development

1. Make changes to files in `thunderbird_extension/`
2. Rebuild the .xpi file
3. Reinstall in Thunderbird
4. Restart Thunderbird to load changes

## Requirements

### System Requirements
- **Thunderbird**: Version 78 or later
- **Python**: 3.7 or later
- **WebSocket Support**: Built into Python 3.7+

### Dependencies
- `websockets` - For WebSocket communication
- `asyncio` - For asynchronous operations (built-in)

### Installation Requirements
- Write access to Thunderbird profile directory
- Ability to install browser extensions
- Local network access for WebSocket communication

## Performance Considerations

1. **WebSocket Overhead**: Minimal overhead for local communication
2. **Memory Usage**: Extension uses minimal memory
3. **Network Usage**: Only local WebSocket traffic
4. **CPU Usage**: Low CPU usage during normal operation

## Future Enhancements

1. **Template Support**: Enhanced template management
2. **Account Management**: Better account configuration
3. **Error Recovery**: Improved error handling and retry logic
4. **Batch Operations**: Support for batch email operations
5. **UI Integration**: Better integration with Thunderbird UI

## Support

For issues related to the WebExtension integration:

1. Check the troubleshooting section above
2. Run the validation script
3. Check the Browser Console in Thunderbird
4. Review Python application logs
5. Ensure all requirements are met

## License

This WebExtension integration is part of the Email Automation Desktop application and follows the same license terms.