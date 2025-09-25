// Background script for Email Automation Thunderbird Extension
// Handles communication with Python application and compose API

let pythonPort = null;
let composeTabId = null;

// Listen for connections from Python application
browser.runtime.onConnect.addListener(function(port) {
  if (port.name === "python-connection") {
    pythonPort = port;

    port.onMessage.addListener(function(message) {
      handlePythonMessage(message);
    });

    port.onDisconnect.addListener(function() {
      pythonPort = null;
      console.log("Python connection disconnected");
    });

    console.log("Python connection established");
  }
});

// Handle messages from Python application
async function handlePythonMessage(message) {
  try {
    console.log("Received message from Python:", message);

    switch (message.action) {
      case "sendEmail":
        await sendEmailViaComposeAPI(message.emailData);
        break;

      case "checkComposeAvailability":
        const available = await checkComposeAPIAvailability();
        pythonPort.postMessage({
          type: "composeAvailability",
          available: available
        });
        break;

      case "getAccounts":
        const accounts = await getThunderbirdAccounts();
        pythonPort.postMessage({
          type: "accounts",
          accounts: accounts
        });
        break;

      default:
        console.warn("Unknown action:", message.action);
    }
  } catch (error) {
    console.error("Error handling Python message:", error);
    if (pythonPort) {
      pythonPort.postMessage({
        type: "error",
        error: error.message
      });
    }
  }
}

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
    setupComposeEventListeners(composeTabId);

    // Send the email
    const sendResult = await browser.compose.sendMessage(composeTabId, {
      mode: "sendNow"
    });

    console.log("Email sent successfully:", sendResult);

    // Notify Python of success
    if (pythonPort) {
      pythonPort.postMessage({
        type: "emailSent",
        success: true,
        messageId: sendResult.headerMessageId
      });
    }

  } catch (error) {
    console.error("Error sending email via Compose API:", error);

    // Notify Python of failure
    if (pythonPort) {
      pythonPort.postMessage({
        type: "emailSent",
        success: false,
        error: error.message
      });
    }
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
        if (pythonPort) {
          pythonPort.postMessage({
            type: "sendError",
            error: sendInfo.error
          });
        }
      } else {
        console.log("Email sent successfully");
        if (pythonPort) {
          pythonPort.postMessage({
            type: "sendSuccess",
            messageId: sendInfo.headerMessageId
          });
        }
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

// Periodic health check
setInterval(() => {
  if (pythonPort) {
    pythonPort.postMessage({
      type: "ping",
      timestamp: Date.now()
    });
  }
}, 30000); // Every 30 seconds