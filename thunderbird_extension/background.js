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

async function sendEmailViaCompose(job) {
  const { id, payload } = job;
  const to = Array.isArray(payload?.to) ? payload.to : [];
  const cc = Array.isArray(payload?.cc) ? payload.cc : [];
  const bcc = Array.isArray(payload?.bcc) ? payload.bcc : [];
  const subject = payload?.subject || "";
  const bodyHtml = payload?.bodyHtml || payload?.body || "";
  const attachments = Array.isArray(payload?.attachments) ? payload.attachments : [];
  
  // Diagnostics: what compose methods are available?
  try {
    const keys = Object.keys((TB && TB.compose) || {});
    log("compose API keys:", keys);
    log("compose.saveMessage typeof:", typeof (TB?.compose?.saveMessage));
    log("compose.sendMessage typeof:", typeof (TB?.compose?.sendMessage));
  } catch (e) {
    log("compose API introspection error:", e);
  }
  
  let composeTab = null;

  try {
    // Resolve an identityId to ensure sending is associated with a valid account identity
    let identityId = undefined;
    try {
      const accounts = await TB.accounts.list();
      if (accounts && accounts.length) {
        for (const acc of accounts) {
          if (acc?.identities?.length) {
            identityId = acc.identities[0].id;
            break;
          }
        }
      }
      log("Using identityId:", identityId);
    } catch (eId) {
      log("accounts.list error:", eId);
    }
  
    // Begin new compose with base details (include identityId when available)
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
  
    // Fallback to ensure HTML body and identity in some TB versions
    if (bodyHtml || identityId) {
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
      try {
        await TB.compose.addAttachment(composeTab.id, { file: fileUrl });
      } catch (e) {
        log("Failed to add attachment:", fileUrl, e);
        // Continue without blocking send; report partial failure in result
      }
    }

    // Send now prioritizing broad compatibility, with multiple fallbacks
    let sendOk = false;
  
    // 1) Try saveMessage('sendNow') - supported in several TB versions
    try {
      if (typeof TB.compose.saveMessage === "function") {
        await TB.compose.saveMessage(composeTab.id, { mode: "sendNow" });
        sendOk = true;
      }
    } catch (e0) {
      log("compose.saveMessage(sendNow) failed:", e0);
      // continue fallbacks
    }
  
    // 2) Fallback to sendMessage({mode:'sendNow'}) - newer TB
    if (!sendOk) {
      try {
        if (typeof TB.compose.sendMessage === "function") {
          await TB.compose.sendMessage(composeTab.id, { mode: "sendNow" });
          sendOk = true;
        }
      } catch (e1) {
        log("compose.sendMessage({mode:\"sendNow\"}) failed:", e1);
      }
    }
  
    // 3) Fallback to sendMessage('sendNow') - legacy signature
    if (!sendOk) {
      try {
        if (typeof TB.compose.sendMessage === "function") {
          await TB.compose.sendMessage(composeTab.id, "sendNow");
          sendOk = true;
        }
      } catch (e2) {
        log('compose.sendMessage("sendNow") failed:', e2);
      }
    }
  
    // 4) Fallback to saveMessage('sendLater') to Outbox/Unsent
    if (!sendOk) {
      try {
        if (typeof TB.compose.saveMessage === "function") {
          await TB.compose.saveMessage(composeTab.id, { mode: "sendLater" });
          sendOk = true;
        }
      } catch (e3) {
        log('compose.saveMessage("sendLater") failed:', e3);
      }
    }
  
    // 5) Try experiment API eSend.sendNow (fallback when compose API lacks send)
    if (!sendOk) {
      try {
        if (TB && TB.eSend && typeof TB.eSend.sendNow === "function") {
          await TB.eSend.sendNow(composeTab.id);
          sendOk = true;
        }
      } catch (e5) {
        log('eSend.sendNow failed:', e5);
      }
    }
  
    // 6) Last resort: save as draft (user can manually send)
    if (!sendOk) {
      try {
        if (typeof TB.compose.saveMessage === "function") {
          await TB.compose.saveMessage(composeTab.id, { mode: "draft" });
        }
      } catch (e4) {
        log('compose.saveMessage("draft") failed:', e4);
      }
      if (port) port.postMessage({ id, success: false, error: "Unable to send automatically; message saved (or attempted) as draft." });
      return;
    }
  
    // Report success to native host
    if (port) port.postMessage({ id, success: true });
  } catch (e) {
    log("sendEmailViaCompose error:", e);
    if (port) port.postMessage({ id, success: false, error: String(e && e.message ? e.message : e) });
  } finally {
    // Attempt to close compose tab if it still exists
    try {
      if (composeTab && composeTab.id) {
        await TB.tabs.remove(composeTab.id);
      }
    } catch (_) {
      // ignore
    }
  }
}

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
  } catch (e) {
    log("handlePortMessage error:", e);
  }
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