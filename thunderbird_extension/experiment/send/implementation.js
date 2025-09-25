// Experiment API to programmatically trigger "Send Now" on a compose window.
// Used as a compatibility fallback when compose.sendMessage/saveMessage are unavailable.

"use strict";

var { ExtensionCommon } = ChromeUtils.import("resource://gre/modules/ExtensionCommon.jsm");
var { Services } = ChromeUtils.import("resource://gre/modules/Services.jsm");

const { ExtensionAPI } = ExtensionCommon;

var eSend = class extends ExtensionAPI {
  getAPI(context) {
    return {
      eSend: {
        async sendNow(tabId) {
          try {
            // Best-effort: map tabId to a compose window. As a minimal fallback,
            // use the most recent compose window.
            let win = Services.wm.getMostRecentWindow("msgcompose");
            if (!win) {
              throw new Error("No compose window found");
            }

            // nsIMsgCompDeliverMode.Now
            const { Ci } = Components;

            // Prefer GenericSendMessage if available
            if (typeof win.GenericSendMessage === "function") {
              win.GenericSendMessage(Ci.nsIMsgCompDeliverMode.Now);
              return true;
            }

            // Older compose windows expose SendMessage
            if (typeof win.SendMessage === "function") {
              win.SendMessage(Ci.nsIMsgCompDeliverMode.Now);
              return true;
            }

            // As a last resort, try triggering menu command if exposed
            try {
              if (typeof win.goDoCommand === "function") {
                win.goDoCommand("cmd_sendNow");
                return true;
              }
            } catch (_) {
              // ignore
            }

            throw new Error("No supported compose send function found");
          } catch (e) {
            // Surface detailed error back to the extension
            return Promise.reject(String(e && e.message ? e.message : e));
          }
        },
      },
    };
  }
};