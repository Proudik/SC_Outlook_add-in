/* global Office */

import { onMessageSendHandler } from "./onMessageSendHandler";

console.log("[commands.ts] Script loaded");

let associated = false;

function associateHandlers() {
  if (associated) return;
  associated = true;

  try {
    if (!Office?.actions?.associate) {
      console.warn("[commands.ts] Office.actions.associate not available");
      return;
    }

    console.log("[commands.ts] Associating onMessageSendHandler");
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    console.log("[commands.ts] Handler associated successfully");
  } catch (e) {
    console.error("[commands.ts] Failed to associate handler:", e);
  }
}

async function boot() {
  try {
    if (typeof Office?.onReady === "function") {
      await Office.onReady();
      console.log("[commands.ts] Office.onReady fired");
      console.log("[commands.ts] Office.context:", Office.context);
    } else {
      console.warn("[commands.ts] Office.onReady not available");
    }
  } catch (e) {
    console.error("[commands.ts] Office.onReady failed:", e);
  } finally {
    associateHandlers();
  }
}

// Start immediately, but also try again onReady.
// This avoids cases where the script runs before Office runtime is fully initialised.
boot();
try {
  if (typeof Office?.onReady === "function") {
    Office.onReady(() => {
      associateHandlers();
    });
  }
} catch {
  // ignore
}