// Content script for Email Automation Thunderbird Extension
// Handles communication between extension and Python application

console.log("Email Automation Content Script loaded");

// Create connection to background script
const backgroundPort = browser.runtime.connect({name: "content-background"});

// Listen for messages from background script
backgroundPort.onMessage.addListener(function(message) {
  console.log("Content script received message:", message);

  switch (message.type) {
    case "ping":
      // Respond to ping
      backgroundPort.postMessage({
        type: "pong",
        timestamp: message.timestamp
      });
      break;

    default:
      console.log("Unknown message type:", message.type);
  }
});

// Function to send message to Python application
function sendToPython(message) {
  // Try to find Python connection window
  const pythonWindows = browser.extension.getViews({type: "popup"});

  if (pythonWindows.length > 0) {
    // Send message to Python window
    pythonWindows[0].postMessage(message, "*");
  } else {
    console.warn("No Python window found to send message to");
  }
}

// Listen for messages from Python application (via window.postMessage)
window.addEventListener("message", function(event) {
  // Verify origin for security
  if (event.origin !== window.location.origin) {
    return;
  }

  if (event.data.type === "pythonMessage") {
    console.log("Received message from Python:", event.data);

    // Forward to background script
    backgroundPort.postMessage(event.data);
  }
});

// Monitor for compose window events
browser.compose.onBeforeSend.addListener((tab, details) => {
  console.log("Email about to be sent:", details);

  // Send notification to Python
  sendToPython({
    type: "beforeSend",
    tabId: tab.id,
    details: details
  });
});

// Monitor for attachment events
browser.compose.onAttachmentAdded.addListener((tab, attachment) => {
  console.log("Attachment added:", attachment);

  sendToPython({
    type: "attachmentAdded",
    tabId: tab.id,
    attachment: attachment
  });
});

// Monitor for attachment removal events
browser.compose.onAttachmentRemoved.addListener((tab, attachmentId) => {
  console.log("Attachment removed:", attachmentId);

  sendToPython({
    type: "attachmentRemoved",
    tabId: tab.id,
    attachmentId: attachmentId
  });
});

// Monitor for identity changes
browser.compose.onIdentityChanged.addListener((tab, identityId) => {
  console.log("Identity changed:", identityId);

  sendToPython({
    type: "identityChanged",
    tabId: tab.id,
    identityId: identityId
  });
});

// Monitor for compose state changes
browser.compose.onComposeStateChanged.addListener((tab, state) => {
  console.log("Compose state changed:", state);

  sendToPython({
    type: "composeStateChanged",
    tabId: tab.id,
    state: state
  });
});

// Function to get current compose details
async function getComposeDetails(tabId) {
  try {
    const details = await browser.compose.getComposeDetails(tabId);
    return details;
  } catch (error) {
    console.error("Error getting compose details:", error);
    return null;
  }
}

// Function to update compose details
async function updateComposeDetails(tabId, details) {
  try {
    await browser.compose.setComposeDetails(tabId, details);
    return true;
  } catch (error) {
    console.error("Error updating compose details:", error);
    return false;
  }
}

// Function to get compose state
async function getComposeState(tabId) {
  try {
    const state = await browser.compose.getComposeState(tabId);
    return state;
  } catch (error) {
    console.error("Error getting compose state:", error);
    return null;
  }
}

// Function to list attachments
async function listAttachments(tabId) {
  try {
    const attachments = await browser.compose.listAttachments(tabId);
    return attachments;
  } catch (error) {
    console.error("Error listing attachments:", error);
    return [];
  }
}

// Function to remove attachment
async function removeAttachment(tabId, attachmentId) {
  try {
    await browser.compose.removeAttachment(tabId, attachmentId);
    return true;
  } catch (error) {
    console.error("Error removing attachment:", error);
    return false;
  }
}

// Function to update attachment
async function updateAttachment(tabId, attachmentId, attachment) {
  try {
    const updatedAttachment = await browser.compose.updateAttachment(tabId, attachmentId, attachment);
    return updatedAttachment;
  } catch (error) {
    console.error("Error updating attachment:", error);
    return null;
  }
}

// Function to save message as draft
async function saveMessage(tabId, options = {}) {
  try {
    const result = await browser.compose.saveMessage(tabId, options);
    return result;
  } catch (error) {
    console.error("Error saving message:", error);
    return null;
  }
}

// Function to send message
async function sendMessage(tabId, options = {}) {
  try {
    const result = await browser.compose.sendMessage(tabId, options);
    return result;
  } catch (error) {
    console.error("Error sending message:", error);
    throw error;
  }
}

// Function to get attachment file
async function getAttachmentFile(attachmentId) {
  try {
    const file = await browser.compose.getAttachmentFile(attachmentId);
    return file;
  } catch (error) {
    console.error("Error getting attachment file:", error);
    return null;
  }
}

// Function to set active dictionaries
async function setActiveDictionaries(tabId, activeDictionaries) {
  try {
    await browser.compose.setActiveDictionaries(tabId, activeDictionaries);
    return true;
  } catch (error) {
    console.error("Error setting active dictionaries:", error);
    return false;
  }
}

// Function to get active dictionaries
async function getActiveDictionaries(tabId) {
  try {
    const dictionaries = await browser.compose.getActiveDictionaries(tabId);
    return dictionaries;
  } catch (error) {
    console.error("Error getting active dictionaries:", error);
    return null;
  }
}

// Expose functions to background script
backgroundPort.postMessage({
  type: "contentReady",
  functions: [
    "getComposeDetails",
    "updateComposeDetails",
    "getComposeState",
    "listAttachments",
    "removeAttachment",
    "updateAttachment",
    "saveMessage",
    "sendMessage",
    "getAttachmentFile",
    "setActiveDictionaries",
    "getActiveDictionaries"
  ]
});

console.log("Content script initialized and ready");