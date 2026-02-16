import * as React from "react";
import { PublicClientApplication, type AuthenticationResult } from "@azure/msal-browser";
import { makeStyles } from "@fluentui/react-components";

declare const Office: any;

const CLIENT_ID = "82c4dffb-91eb-49d4-a723-bae6232465f5";
const TENANT_ID = "f73679c2-a878-4ab2-a898-1d9405ccb695";

// Must be registered in Entra as SPA Redirect URI exactly
const REDIRECT_URI = `${window.location.origin}/dialog.html?mode=graph`;

const SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "Mail.ReadWrite",
  "MailboxSettings.ReadWrite",
];

const msal = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
});

// One time initialisation gate (newer MSAL needs initialize(), older does not have it)
let msalInitPromise: Promise<void> | null = null;
function ensureMsalReady(): Promise<void> {
  if (!msalInitPromise) {
    const init = (msal as any).initialize;
    msalInitPromise = typeof init === "function" ? init.call(msal) : Promise.resolve();
  }
  return msalInitPromise;
}

const useStyles = makeStyles({
  page: {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "18px",
    background: "rgb(242, 238, 231)",
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif',
  },
  card: {
    width: "100%",
    maxWidth: "520px",
    borderRadius: "14px",
    border: "1px solid rgba(0,0,0,0.08)",
    background: "rgba(255,255,255,0.92)",
    boxShadow: "0 12px 30px rgba(0,0,0,0.06)",
    padding: "18px",
    boxSizing: "border-box",
  },
  title: { margin: 0, fontSize: "14px", fontWeight: 700, color: "rgba(17,24,39,0.92)" },
  hint: { margin: "8px 0 0 0", fontSize: "12px", color: "rgba(17,24,39,0.58)", lineHeight: 1.35 },
  error: {
    marginTop: "10px",
    borderRadius: "10px",
    padding: "10px 12px",
    border: "1px solid rgba(220, 38, 38, 0.22)",
    background: "rgba(220, 38, 38, 0.08)",
    color: "rgba(153, 27, 27, 0.95)",
    fontSize: "12px",
    lineHeight: 1.35,
    wordBreak: "break-word",
  },
  actions: { display: "flex", justifyContent: "flex-end", gap: "10px", marginTop: "14px" },
  btn: {
    height: "36px",
    borderRadius: "10px",
    border: "1px solid rgba(0,0,0,0.12)",
    background: "rgba(255,255,255,0.75)",
    padding: "0 12px",
    cursor: "pointer",
    fontWeight: 600,
    fontSize: "12px",
    color: "rgba(17,24,39,0.86)",
  },
  primary: {
    border: "none",
    background: "#00204A",
    color: "white",
    WebkitTextFillColor: "white",
  },
});

function messageParent(payload: unknown) {
  const msg = JSON.stringify(payload);
  try {
    Office?.context?.ui?.messageParent?.(msg);
  } catch {
    // ignore
  }
}

function isGraphMode(): boolean {
  try {
    const url = new URL(window.location.href);
    const mode = String(url.searchParams.get("mode") || "").toLowerCase();
    if (mode === "graph") return true;

    const h = String(window.location.hash || "").toLowerCase();
    if (h.includes("graph")) return true;

    return false;
  } catch {
    return false;
  }
}

export default function GraphAuthDialog() {
  const styles = useStyles();
  const [error, setError] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);

  React.useEffect((): void | (() => void) => {
    if (!isGraphMode()) return undefined;

    let cancelled = false;

    const run = async (): Promise<void> => {
      try {
        setBusy(true);
        setError(null);

        await ensureMsalReady();
        if (cancelled) return;

        const redirectResult: AuthenticationResult | null = await msal.handleRedirectPromise();
        if (cancelled) return;

        const account = redirectResult?.account || msal.getAllAccounts()[0] || null;

        if (!account) {
          await msal.loginRedirect({ scopes: SCOPES });
          return;
        }

        const tokenRes = await msal.acquireTokenSilent({ account, scopes: SCOPES });
        if (cancelled) return;

        messageParent({
          type: "graph_auth_complete",
          accessToken: tokenRes.accessToken,
          account: { username: account.username, homeAccountId: account.homeAccountId },
        });
      } catch (e) {
        if (cancelled) return;
        setError(e instanceof Error ? e.message : String(e));
      } finally {
        if (!cancelled) setBusy(false);
      }
    };

    void run();

    return () => {
      cancelled = true;
    };
  }, []);

  const startLogin = async (): Promise<void> => {
    setError(null);
    setBusy(true);
    try {
      await ensureMsalReady();
      await msal.loginRedirect({ scopes: SCOPES });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      setBusy(false);
    }
  };

  return (
    <div className={styles.page}>
      <div className={styles.card}>
        <p className={styles.title}>Microsoft login</p>
        <p className={styles.hint}>This is used only to count unfiled emails in Inbox and Sent via Microsoft Graph.</p>

        {error ? <div className={styles.error}>{error}</div> : null}

        <div className={styles.actions}>
          <button type="button" className={styles.btn} onClick={() => messageParent({ type: "graph_auth_cancel" })} disabled={busy}>
            Cancel
          </button>

          <button type="button" className={`${styles.btn} ${styles.primary}`} onClick={startLogin} disabled={busy}>
            {busy ? "Working" : "Continue"}
          </button>
        </div>
      </div>
    </div>
  );
}