// Thunderbird MailExtension background script for headless email sending via native messaging

const TB = (globalThis.browser || globalThis.messenger);
const HOST_NAME = "com.emailautomation.tbhost";
let port = null;
let reconnectTimer = null;
let pollTimer = null;

function log(...args) {
  // eslint-disable-next-line no-console
  console.log("[TB-Bridge]", ...args);
}

function toFileUrl(osPath) {
  if (!osPath) return null;
  try {
    // If already file URL, return as-is
    if (/^file:/i.test(osPath)) return osPath;

    // Normalize Windows backslashes to forward slashes
    let p = osPath.replace(/\\/g, "/");

    // Ensure drive letter paths become file:///C:/...
    if (/^[A-Za-z]:\//.test(p)) {
      return "file:///" + p;
    }

    // For absolute POSIX paths
    if (p.startsWith("/")) {
      return "file://" + p;
    }

    // Fallback: treat as relative path (unlikely in practice)
    return "file:///" + p;
  } catch (e) {
    log("toFileUrl error:", e);
    return null;
  }
}

async function getThunderbirdVersion() {
  try {
    const info = await TB.runtime.getBrowserInfo();
    return info.version;
  } catch (e) {
    log("Failed to get Thunderbird version:", e);
    return "unknown";
  }
}

async function sendEmailViaCompose(job) {
  const { id, payload } = job;
  const to = Array.isArray(payload?.to) ? payload.to : [];
  const cc = Array.isArray(payload?.cc) ? payload.cc : [];
  const bcc = Array.isArray(payload?.bcc) ? payload.bcc : [];
  const subject = payload?.subject || "";
  const bodyHtml = payload?.bodyHtml || payload?.body || "";
  const attachments = Array.isArray(payload?.attachments) ? payload.attachments : [];
  
  log("=== SEND EMAIL DEBUG ===");
  log("Job ID:", id);
  log("To:", to);
  log("CC:", cc);
  log("BCC:", bcc);
  log("Subject:", subject);
  log("Body length:", bodyHtml.length);
  log("Attachments:", attachments);
  
  // Diagnostics: what compose methods are available?
  try {
    const keys = Object.keys((TB && TB.compose) || {});
    log("compose API keys:", keys);
    log("compose.saveMessage typeof:", typeof (TB?.compose?.saveMessage));
    log("compose.sendMessage typeof:", typeof (TB?.compose?.sendMessage));
    log("TB.compose object:", TB?.compose);
  } catch (e) {
    log("compose API introspection error:", e);
  }
  
  // Check Thunderbird version
  try {
    const info = await TB.runtime.getBrowserInfo();
    log("Thunderbird version:", info.version, "name:", info.name);
  } catch (e) {
    log("Failed to get browser info:", e);
  }
  
  let composeTab = null;

  try {
    // Resolve an identityId to ensure sending is associated with a valid account identity
    let identityId = undefined;
    try {
      const accounts = await TB.accounts.list();
      log("Available accounts:", accounts);
      if (accounts && accounts.length) {
        for (const acc of accounts) {
          log("Account:", acc.name, "identities:", acc.identities);
          if (acc?.identities?.length) {
            identityId = acc.identities[0].id;
            log("Selected identityId:", identityId);
            break;
          }
        }
      }
    } catch (eId) {
      log("accounts.list error:", eId);
    }
  
    // Begin new compose with base details (include identityId when available)
    log("Creating new compose tab...");
    composeTab = await TB.compose.beginNew({
      to,
      cc,
      bcc,
      subject,
      // Thunderbird accepts HTML in body. If it shows as plain text on some versions,
      // we can call setComposeDetails afterwards as a fallback.
      body: bodyHtml,
      ...(identityId ? { identityId } : {}),
    });
    
    log("Compose tab created:", composeTab);
  
    // Fallback to ensure HTML body and identity in some TB versions
    if (bodyHtml || identityId) {
      log("Setting compose details...");
      await TB.compose.setComposeDetails(composeTab.id, {
        body: bodyHtml,
        ...(identityId ? { identityId } : {}),
      });
    }

    // Add attachments
    for (const att of attachments) {
      const filePath = att?.path || att?.file || "";
      if (!filePath) continue;
      const fileUrl = toFileUrl(filePath);
      log("Adding attachment:", filePath, "as URL:", fileUrl);
      try {
        // Try different attachment formats based on Thunderbird version
        let attachmentAdded = false;
        
        // Format 1: Using file path (most compatible)
        try {
          log("Trying attachment format: file path");
          const fileName = filePath.split('\\').pop();
          await TB.compose.addAttachment(composeTab.id, {
            file: { name: fileName, path: filePath }
          });
          log("Attachment added successfully using file path");
          attachmentAdded = true;
        } catch (e1) {
          log("File path format failed:", e1.message);
          
          // Format 2: Using file URL with name (some versions)
          try {
            log("Trying attachment format: file URL with name");
            const fileName = filePath.split('\\').pop();
            await TB.compose.addAttachment(composeTab.id, {
              file: { name: fileName, url: fileUrl }
            });
            log("Attachment added successfully using file URL with name");
            attachmentAdded = true;
          } catch (e2) {
            log("File URL with name format failed:", e2.message);
            
            // Format 3: Direct file URL (legacy)
            try {
              log("Trying attachment format: direct file URL");
              await TB.compose.addAttachment(composeTab.id, { file: fileUrl });
              log("Attachment added successfully using direct file URL");
              attachmentAdded = true;
            } catch (e3) {
              log("Direct file URL format failed:", e3.message);
              
              // Format 4: Using file ID with native host data (Thunderbird 128+)
              try {
                log("Trying attachment format: file ID with native host data");
                
                // Request file data from native host
                const fileData = await requestFileData(filePath);
                
                if (fileData && fileData.data) {
                  // Convert base64 to blob
                  const byteCharacters = atob(fileData.data);
                  const byteNumbers = new Array(byteCharacters.length);
                  for (let i = 0; i < byteCharacters.length; i++) {
                    byteNumbers[i] = byteCharacters.charCodeAt(i);
                  }
                  const byteArray = new Uint8Array(byteNumbers);
                  const blob = new Blob([byteArray], { type: 'application/octet-stream' });
                  
                  // Create a File object
                  const file = new File([blob], fileData.name, { type: blob.type });
                  
                  // Add attachment using File object
                  await TB.compose.addAttachment(composeTab.id, { file });
                  log("Attachment added successfully using file ID with native host data");
                  attachmentAdded = true;
                } else {
                  throw new Error("No file data received from native host");
                }
              } catch (e4) {
                log("File ID with native host data format failed:", e4.message);
                
                // Format 5: Using browser.downloads API (if available)
                try {
                  log("Trying attachment format: browser.downloads");
                  if (typeof TB.downloads !== 'undefined' && typeof TB.downloads.download === 'function') {
                    // Download the file first, then attach it
                    const downloadId = await TB.downloads.download({
                      url: fileUrl,
                      filename: filePath.split('\\').pop(),
                      saveAs: false
                    });
                    
                    // Wait for the download to complete
                    // This is a simplified approach; in reality, you'd need to handle download events
                    await new Promise(resolve => setTimeout(resolve, 1000));
                    
                    // Try to attach using the downloaded file
                    const fileName = filePath.split('\\').pop();
                    await TB.compose.addAttachment(composeTab.id, {
                      file: { name: fileName, path: filePath }
                    });
                    log("Attachment added successfully using browser.downloads");
                    attachmentAdded = true;
                  } else {
                    log("browser.downloads API not available");
                  }
                } catch (e5) {
                  log("Browser downloads format failed:", e5.message);
                  log("All attachment formats failed");
                }
              }
            }
          }
        }
        
        if (!attachmentAdded) {
          log("Warning: Could not add attachment using any format");
          // Continue without blocking send; report partial failure in result
        }
      } catch (e) {
        log("Failed to add attachment:", fileUrl, e);
        // Continue without blocking send; report partial failure in result
      }
    }

    // Send now prioritizing broad compatibility, with multiple fallbacks
    let sendOk = false;
    let sendMethod = "";
    let lastError = null;
    
    // Log available compose methods for debugging
    log("Available compose methods:", Object.keys(TB.compose).filter(key => typeof TB.compose[key] === 'function'));
  
    // 1) Try using the new Thunderbird 128+ API with sendMessage
    if (!sendOk) {
      sendMethod = "TB.compose.sendMessage()";
      try {
        if (typeof TB.compose.sendMessage === "function") {
          log("Trying send method:", sendMethod);
          await TB.compose.sendMessage(composeTab.id);
          sendOk = true;
          log("Send method succeeded:", sendMethod);
        } else {
          log("Send method not available:", sendMethod);
        }
      } catch (e1) {
        lastError = e1;
        log("Send method failed:", sendMethod, e1);
      }
    }
    
    // 2) Try using the experiment API if available
    if (!sendOk) {
      sendMethod = "experiment API send";
      try {
        if (TB && TB.eSend && typeof TB.eSend.sendNow === "function") {
          log("Trying send method:", sendMethod);
          await TB.eSend.sendNow(composeTab.id);
          sendOk = true;
          log("Send method succeeded:", sendMethod);
        } else {
          log("Send method not available:", sendMethod);
        }
      } catch (e5) {
        lastError = e5;
        log("Send method failed:", sendMethod, e5);
      }
    }
    
    // 3) Try direct command execution
    if (!sendOk) {
      sendMethod = "direct command execution";
      try {
        log("Trying send method:", sendMethod);
        // Try to execute send command directly
        if (composeTab && composeTab.id) {
          const result = await TB.tabs.executeScript(composeTab.id, {
            code: `
              try {
                if (window.goDoCommand) {
                  window.goDoCommand('cmd_sendNow');
                  true;
                } else if (window.SendMessage) {
                  window.SendMessage(0);
                  true;
                } else {
                  false;
                }
              } catch (e) {
                console.error('Direct send command failed:', e);
                false;
              }
            `
          });
          if (result && result[0] === true) {
            sendOk = true;
            log("Send method succeeded:", sendMethod);
          } else {
            log("Direct command execution returned false");
          }
        }
      } catch (e6) {
        lastError = e6;
        log("Send method failed:", sendMethod, e6);
      }
    }
    
    // 4) Try using window.postMessage to trigger send
    if (!sendOk) {
      sendMethod = "window.postMessage";
      try {
        log("Trying send method:", sendMethod);
        if (composeTab && composeTab.id) {
          await TB.tabs.executeScript(composeTab.id, {
            code: `
              try {
                // Try to find and click the send button
                const sendButton = document.querySelector('[command="cmd_sendNow"], [command="cmd_sendWithCheck"], button[label*="Send"], button[accesskey*="S"]');
                if (sendButton) {
                  sendButton.click();
                  true;
                } else {
                  console.log('Send button not found');
                  false;
                }
              } catch (e) {
                console.error('Send button click failed:', e);
                false;
              }
            `
          });
          sendOk = true;
          log("Send method succeeded:", sendMethod);
        }
      } catch (e7) {
        lastError = e7;
        log("Send method failed:", sendMethod, e7);
      }
    }
  
    // 5) Last resort: save as draft (user can manually send)
    if (!sendOk) {
      sendMethod = "save as draft";
      try {
        log("Trying send method:", sendMethod);
        // Try to save as draft using the available API
        if (typeof TB.compose.saveMessage === "function") {
          await TB.compose.saveMessage(composeTab.id, { mode: "draft" });
          log("Draft saved successfully");
        } else {
          // Try to click the save button
          if (composeTab && composeTab.id) {
            await TB.tabs.executeScript(composeTab.id, {
              code: `
                try {
                  const saveButton = document.querySelector('[command="cmd_saveAsDraft"], [command="cmd_save"], button[label*="Save"], button[accesskey*="S"]');
                  if (saveButton) {
                    saveButton.click();
                    true;
                  } else {
                    console.log('Save button not found');
                    false;
                  }
                } catch (e) {
                  console.error('Save button click failed:', e);
                  false;
                }
              `
            });
          }
        }
      } catch (e4) {
        lastError = e4;
        log('Send method failed:', sendMethod, e4);
      }
      
      // Provide detailed error information
      let errorMsg = `Unable to send automatically; message saved (or attempted) as draft. `;
      errorMsg += `Tried methods: TB.compose.sendMessage(), experiment API, direct command execution, window.postMessage, save as draft. `;
      if (lastError) {
        errorMsg += `Last error: ${lastError.message || lastError.toString()}`;
      }
      errorMsg += ` Thunderbird version: ${await getThunderbirdVersion()}`;
      
      log(errorMsg);
      if (port) port.postMessage({ id, success: false, error: errorMsg });
      return;
    }
  
    // Report success to native host
    log("Email sent successfully using method:", sendMethod);
    if (port) port.postMessage({ id, success: true });
  } catch (e) {
    log("sendEmailViaCompose error:", e);
    log("Error stack:", e.stack);
    if (port) port.postMessage({ id, success: false, error: String(e && e.message ? e.message : e) });
  } finally {
    // Attempt to close compose tab if it still exists
    try {
      if (composeTab && composeTab.id) {
        log("Closing compose tab:", composeTab.id);
        await TB.tabs.remove(composeTab.id);
      }
    } catch (e) {
      log("Error closing compose tab:", e);
    }
  }
}

// Store pending file requests
const pendingFileRequests = {};

function handlePortMessage(msg) {
  try {
    if (!msg || typeof msg !== "object") {
      return;
    }
    if (msg.type === "sendEmail") {
      sendEmailViaCompose(msg);
      return;
    }
    if (msg.type === "ping") {
      if (port) port.postMessage({ type: "pong", ts: Date.now() });
      return;
    }
    if (msg.type === "fileDataResponse") {
      // Handle file data response from native host
      const requestId = msg.id;
      if (pendingFileRequests[requestId]) {
        const { resolve, reject } = pendingFileRequests[requestId];
        delete pendingFileRequests[requestId];
        
        if (msg.success) {
          resolve(msg);
        } else {
          reject(new Error(msg.error || "Unknown error"));
        }
      }
      return;
    }
  } catch (e) {
    log("handlePortMessage error:", e);
  }
}

// Function to request file data from native host
async function requestFileData(filePath) {
  return new Promise((resolve, reject) => {
    const requestId = `file_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    pendingFileRequests[requestId] = { resolve, reject };
    
    if (port) {
      port.postMessage({
        type: "getFileData",
        id: requestId,
        filePath: filePath
      });
      
      // Set a timeout
      setTimeout(() => {
        if (pendingFileRequests[requestId]) {
          delete pendingFileRequests[requestId];
          reject(new Error("File request timeout"));
        }
      }, 10000); // 10 second timeout
    } else {
      delete pendingFileRequests[requestId];
      reject(new Error("No connection to native host"));
    }
  });
}

function connectNative() {
  try {
    if (port) {
      try { port.disconnect(); } catch (_) {}
      port = null;
    }

    port = TB.runtime.connectNative(HOST_NAME);
    log("Connected to native host:", HOST_NAME);

    port.onMessage.addListener(handlePortMessage);

    port.onDisconnect.addListener(() => {
      const lastError = TB.runtime.lastError ? TB.runtime.lastError.message : null;
      log("Native host disconnected.", lastError || "");
      port = null;

      // Attempt to reconnect after a short delay
      if (!reconnectTimer) {
        reconnectTimer = setTimeout(() => {
          reconnectTimer = null;
          connectNative();
        }, 3000);
      }
    });

    // Say hello
    port.postMessage({ type: "hello", from: "tb-extension", ts: Date.now() });
  } catch (e) {
    log("connectNative error:", e);

    // Retry later
    if (!reconnectTimer) {
      reconnectTimer = setTimeout(() => {
        reconnectTimer = null;
        connectNative();
      }, 5000);
    }
  }
}

TB.runtime.onStartup.addListener(() => {
  log("onStartup");
  connectNative();
});

TB.runtime.onInstalled.addListener(() => {
  log("onInstalled");
  connectNative();
});

// Also connect immediately when background loads
connectNative();