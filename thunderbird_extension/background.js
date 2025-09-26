// Background script for Email Automation Thunderbird Extension
// Handles communication with Python application via native messaging and compose API

let nativePort = null;
let composeTabId = null;

// Initialize native messaging connection
function initializeNativeMessaging() {
  try {
    // Connect to the native messaging host
    nativePort = browser.runtime.connectNative("email_automation_native_host");
    
    nativePort.onMessage.addListener((message) => {
      handleNativeMessage(message);
    });
    
    nativePort.onDisconnect.addListener(() => {
      if (browser.runtime.lastError) {
        console.error("Native messaging disconnected:", browser.runtime.lastError.message);
      }
      nativePort = null;
      console.log("Native messaging connection disconnected");
    });
    
    console.log("Native messaging connection established");
  } catch (error) {
    console.error("Failed to initialize native messaging:", error);
  }
}

// Handle messages from native messaging host
async function handleNativeMessage(message) {
  try {
    console.log("Received message from native host:", message);

    switch (message.type) {
      case "emailSent":
        handleEmailSentResponse(message);
        break;
        
      case "availability":
        handleAvailabilityResponse(message);
        break;
        
      case "accounts":
        handleAccountsResponse(message);
        break;
        
      case "pong":
        console.log("Received pong from native host");
        break;
        
      case "error":
        console.error("Error from native host:", message.error);
        break;
        
      default:
        console.warn("Unknown message type:", message.type);
    }
  } catch (error) {
    console.error("Error handling native message:", error);
  }
}

// Handle email sent response
function handleEmailSentResponse(message) {
  if (message.success) {
    console.log("Email sent successfully:", message.messageId);
  } else {
    console.error("Failed to send email:", message.error);
  }
}

// Handle availability response
function handleAvailabilityResponse(message) {
  console.log("Native host availability:", message.available);
}

// Handle accounts response
function handleAccountsResponse(message) {
  console.log("Received accounts:", message.accounts);
}

// Send message to native messaging host
function sendToNativeHost(message) {
  if (nativePort) {
    try {
      nativePort.postMessage(message);
      console.log("Sent message to native host:", message);
    } catch (error) {
      console.error("Failed to send message to native host:", error);
    }
  } else {
    console.error("Native messaging port not available");
  }
}

// Send email via native messaging host
async function sendEmailViaNativeHost(emailData) {
  return new Promise((resolve, reject) => {
    const requestId = Date.now().toString();
    
    // Store the promise for later resolution
    const pendingRequest = {
      resolve: resolve,
      reject: reject,
      timeout: setTimeout(() => {
        delete pendingRequests[requestId];
        reject(new Error("Timeout waiting for email send response"));
      }, 30000) // 30 second timeout
    };
    
    pendingRequests[requestId] = pendingRequest;
    
    // Send the email request
    sendToNativeHost({
      type: "sendEmail",
      requestId: requestId,
      emailData: emailData
    });
  });
}

// Store pending requests
let pendingRequests = {};

// Send email using Thunderbird Compose API
async function sendEmailViaComposeAPI(emailData) {
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
        // For native messaging, we need to handle file paths differently
        // The attachment path should be accessible to Thunderbird
        const attachment = await browser.compose.addAttachment(
          composeTabId,
          {
            file: emailData.attachmentPath,
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
    setupComposeEventListeners(composeTabId);

    // Send the email
    const sendResult = await browser.compose.sendMessage(composeTabId, {
      mode: "sendNow"
    });

    console.log("Email sent successfully:", sendResult);

    // Notify native host of success
    sendToNativeHost({
      type: "emailSent",
      success: true,
      messageId: sendResult.headerMessageId
    });

  } catch (error) {
    console.error("Error sending email via Compose API:", error);

    // Notify native host of failure
    sendToNativeHost({
      type: "emailSent",
      success: false,
      error: error.message
    });
  }
}

// Setup event listeners for compose window
function setupComposeEventListeners(tabId) {
  // Listen for send events
  browser.compose.onAfterSend.addListener((tab, sendInfo) => {
    if (tab.id === tabId) {
      console.log("Email sent event received:", sendInfo);

      if (sendInfo.error) {
        console.error("Send error:", sendInfo.error);
        sendToNativeHost({
          type: "sendError",
          error: sendInfo.error
        });
      } else {
        console.log("Email sent successfully");
        sendToNativeHost({
          type: "sendSuccess",
          messageId: sendInfo.headerMessageId
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
});

// Initialize native messaging when extension starts
initializeNativeMessaging();

// Periodic health check
setInterval(() => {
  if (nativePort) {
    sendToNativeHost({
      type: "ping",
      timestamp: Date.now()
    });
  }
}, 30000); // Every 30 seconds