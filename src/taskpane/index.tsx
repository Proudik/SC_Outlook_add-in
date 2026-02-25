import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { installDebugCommands } from "../utils/debugCache";
import { initTelemetry, emit } from "../telemetry/telemetry";
import { deriveAnonymousUserId, hashWorkspaceId } from "../telemetry/hashing";
import { getAuth } from "../services/auth";
import { getStored } from "../utils/storage";
import { STORAGE_KEYS } from "../utils/constants";

/* global document, Office, module, require, HTMLElement */

const title = "Contoso Task Pane Add-in";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(async () => {
  // Install debug commands for console access
  installDebugCommands();

  // ── Telemetry bootstrap ────────────────────────────────────────────────────
  // Compute anonymised IDs asynchronously so initTelemetry() receives hashes,
  // never raw PII. Failures here are swallowed — telemetry must not block render.
  try {
    const { email } = getAuth();
    const workspaceHostRaw = await getStored(STORAGE_KEYS.workspaceHost).catch(() => "");
    const workspaceHost = String(workspaceHostRaw || "").trim().toLowerCase();

    const [anonymousUserId, workspaceId] = await Promise.all([
      deriveAnonymousUserId(email || "", workspaceHost),
      hashWorkspaceId(workspaceHost),
    ]);

    initTelemetry({ anonymousUserId, workspaceId });

    // "addin.load" fires once per Office.onReady, which is once per session.
    // platform comes from Office.context.diagnostics if available.
    const platform = String(
      (Office as any)?.context?.diagnostics?.platform || "unknown"
    );
    const composeMode =
      (Office as any)?.context?.mailbox?.item?.itemType === "message" &&
      (Office as any)?.context?.mailbox?.item?.displayReplyAllForm !== undefined
        ? false // rough heuristic — exact value set later in MainWorkspace
        : false;

    emit("addin.load", { composeMode, platform });
  } catch {
    // Silently ignore — telemetry must never block the add-in
  }
  // ──────────────────────────────────────────────────────────────────────────

  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
