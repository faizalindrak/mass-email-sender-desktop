// Background script for Email Automation Thunderbird Extension
// Handles communication with Python application and compose API

let websocket = null;
let composeTabId = null;
const WEBSOCKET_URL = "ws://localhost:8765";
let reconnectInterval = 5000; // 5 seconds

function connect() {
    console.log("Attempting to connect to WebSocket...");
    websocket = new WebSocket(WEBSOCKET_URL);

    websocket.onopen = function(event) {
        console.log("WebSocket connection established");
    };

    websocket.onmessage = function(event) {
        try {
            const message = JSON.parse(event.data);
            handlePythonMessage(message);
        } catch (e) {
            console.error("Error parsing message from Python:", e);
        }
    };

    websocket.onclose = function(event) {
        console.log("WebSocket connection closed. Reconnecting in " + (reconnectInterval / 1000) + " seconds.");
        setTimeout(connect, reconnectInterval);
    };

    websocket.onerror = function(event) {
        console.error("WebSocket error observed:", event);
    };
}

// Handle messages from Python application
async function handlePythonMessage(message) {
  try {
    console.log("Received message from Python:", message);

    // The Python server doesn't send 'action' but 'type'. Let's adapt.
    const action = message.action || message.type;

    switch (action) {
      case "sendEmail":
        await sendEmailViaComposeAPI(message.emailData, message.requestId);
        break;

      case "checkAvailability": // Renamed from checkComposeAvailability for consistency
        const available = await checkComposeAPIAvailability();
        sendMessageToPython({
          type: "availabilityResponse",
          available: available,
          requestId: message.requestId
        });
        break;

      case "getAccounts":
        const accounts = await getThunderbirdAccounts();
        sendMessageToPython({
          type: "accountsResponse",
          accounts: accounts,
          requestId: message.requestId
        });
        break;

      case "pong":
        // Server responded to our ping
        console.log("Received pong from server.");
        break;

      default:
        console.warn("Unknown action:", action);
    }
  } catch (error) {
    console.error("Error handling Python message:", error);
    if (message.requestId) {
        sendMessageToPython({
            type: "error",
            error: error.message,
            requestId: message.requestId
        });
    }
  }
}

function sendMessageToPython(message) {
    if (websocket && websocket.readyState === WebSocket.OPEN) {
        websocket.send(JSON.stringify(message));
    } else {
        console.error("WebSocket is not connected. Cannot send message.");
    }
}

// Send email using Thunderbird Compose API
async function sendEmailViaComposeAPI(emailData, requestId) {
  try {
    console.log("Sending email via Compose API:", emailData);

    // Create new compose window
    const tab = await browser.compose.beginNew({
      to: emailData.to,
      cc: emailData.cc,
      bcc: emailData.bcc,
      subject: emailData.subject,
      body: emailData.body,
      isPlainText: false
    });

    composeTabId = tab.id;
    console.log("Created compose tab:", composeTabId);

    // Add attachment if provided
    if (emailData.attachmentPath) {
      try {
        const attachment = await browser.compose.addAttachment(
          composeTabId,
          {
            file: await fetch(emailData.attachmentPath).then(r => r.blob()),
            name: emailData.attachmentName || "attachment"
          }
        );
        console.log("Attachment added:", attachment);
      } catch (attachmentError) {
        console.warn("Failed to add attachment:", attachmentError);
        // Continue without attachment
      }
    }

    // Listen for compose events
    setupComposeEventListeners(composeTabId, requestId);

    // Send the email
    const sendResult = await browser.compose.sendMessage(composeTabId, {
      mode: "sendNow"
    });

    console.log("Email sent successfully:", sendResult);

    // Notify Python of success
    sendMessageToPython({
      type: "emailSent",
      success: true,
      messageId: sendResult.headerMessageId,
      requestId: requestId
    });

  } catch (error) {
    console.error("Error sending email via Compose API:", error);

    // Notify Python of failure
    sendMessageToPython({
      type: "emailSent",
      success: false,
      error: error.message,
      requestId: requestId
    });
  }
}

// Setup event listeners for compose window
function setupComposeEventListeners(tabId, requestId) {
  // Listen for send events
  browser.compose.onAfterSend.addListener((tab, sendInfo) => {
    if (tab.id === tabId) {
      console.log("Email sent event received:", sendInfo);

      if (sendInfo.error) {
        console.error("Send error:", sendInfo.error);
        sendMessageToPython({
            type: "sendError",
            error: sendInfo.error,
            requestId: requestId
        });
      } else {
        console.log("Email sent successfully");
        sendMessageToPython({
            type: "sendSuccess",
            messageId: sendInfo.headerMessageId,
            requestId: requestId
        });
      }
    }
  });

  // Listen for save events
  browser.compose.onAfterSave.addListener((tab, saveInfo) => {
    if (tab.id === tabId) {
      console.log("Email saved event received:", saveInfo);
    }
  });
}

// Check if Compose API is available
async function checkComposeAPIAvailability() {
  try {
    // Try to create a test compose window
    const testTab = await browser.compose.beginNew({
      to: ["test@example.com"],
      subject: "Test",
      body: "Test"
    });

    // Close the test tab
    await browser.tabs.remove(testTab.id);

    return true;
  } catch (error) {
    console.error("Compose API not available:", error);
    return false;
  }
}

// Get Thunderbird accounts
async function getThunderbirdAccounts() {
  try {
    const accounts = await browser.accounts.list();
    return accounts.map(account => ({
      id: account.id,
      name: account.name,
      type: account.type,
      identities: account.identities.map(identity => ({
        id: identity.id,
        name: identity.name,
        email: identity.email
      }))
    }));
  } catch (error) {
    console.error("Error getting accounts:", error);
    return [];
  }
}

// Initial connection
connect();

// Handle extension installation
browser.runtime.onInstalled.addListener((details) => {
  console.log("Email Automation Extension installed/updated");

  if (details.reason === "install") {
    console.log("Extension installed for the first time");
  } else if (details.reason === "update") {
    console.log("Extension updated");
  }
});

// Handle extension startup
browser.runtime.onStartup.addListener(() => {
  console.log("Email Automation Extension started");
  if (!websocket || websocket.readyState === WebSocket.CLOSED) {
    connect();
  }
});

// Periodic health check
setInterval(() => {
    if (websocket && websocket.readyState === WebSocket.OPEN) {
        sendMessageToPython({ type: "ping", timestamp: Date.now() });
    } else {
        console.log("Health check: WebSocket not open.");
    }
}, 30000); // Every 30 seconds