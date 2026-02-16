import * as React from "react";
import { makeStyles } from "@fluentui/react-components";

declare const Office: any;

type Workspace = {
  workspaceId: string;
  name: string;
  host: string;
};

const useStyles = makeStyles({
  page: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    justifyContent: "center",
    padding: "18px",
    background: "rgba(18, 34, 55, 0.8)",
    boxSizing: "border-box",
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif',
  },
  card: {
    width: "100%",
    maxWidth: "720px",
    marginLeft: "auto",
    marginRight: "auto",
    borderRadius: "14px",
    border: "1px solid rgba(0,0,0,0.10)",
    background: "rgba(255,255,255,1)",
    boxShadow: "0 12px 30px rgba(0,0,0,0.06)",
    padding: "22px",
    boxSizing: "border-box",
  },
  error: {
    borderRadius: "10px",
    padding: "10px 12px",
    border: "1px solid rgba(180, 0, 0, 0.25)",
    background: "rgba(220, 0, 0, 0.06)",
    color: "rgba(120, 0, 0, 0.95)",
    fontSize: "13px",
    lineHeight: 1.35,
    marginBottom: "12px",
  },

  consentTop: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "14px",
    textAlign: "center",
    marginBottom: "18px",
  },
  connectorImg: {
    width: "100%",
    maxWidth: "220px",
    height: "auto",
    display: "block",
    objectFit: "contain",
  },
  consentTitle: {
    margin: 0,
    fontSize: "20px",
    fontWeight: 800,
    color: "rgba(17,24,39,0.95)",
  },
  consentTitleEm: { fontWeight: 900 },

  consentSection: { marginTop: "18px" },
  sectionHeading: {
    margin: "0 0 10px 0",
    fontSize: "13px",
    fontWeight: 700,
    opacity: 0.85,
    textTransform: "none",
  },
  list: { display: "flex", flexDirection: "column", borderTop: "1px solid rgba(0,0,0,0.10)" },
  row: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
    padding: "12px 0",
    borderBottom: "1px solid rgba(0,0,0,0.10)",
  },
  iconBox: {
    width: "34px",
    height: "34px",
    borderRadius: "10px",
    border: "1px solid rgba(0,0,0,0.10)",
    background: "rgba(255,255,255,0.65)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "16px",
    flex: "0 0 auto",
  },
  rowText: { flex: "1 1 auto", fontSize: "14px", color: "rgba(17,24,39,0.92)" },
  chevron: { flex: "0 0 auto", fontSize: "18px", opacity: 0.55 },

  checkboxRow: {
    display: "flex",
    gap: "10px",
    alignItems: "flex-start",
    marginTop: "14px",
    padding: "10px 12px",
    borderRadius: "12px",
    border: "1px solid rgba(0,0,0,0.10)",
    background: "rgba(255,255,255,0.60)",
  },
  checkbox: { marginTop: "3px" },
  agreementText: {
    margin: 0,
    fontSize: "13px",
    lineHeight: "1.35",
    opacity: 0.9,
    textAlign: "left",
  },

  consentActions: {
    display: "flex",
    justifyContent: "center",
    gap: "12px",
    marginTop: "22px",
  },
  cancelBtn: {
    height: "40px",
    minWidth: "110px",
    borderRadius: "12px",
    border: "1px solid rgba(0,0,0,0.18)",
    background: "transparent",
    cursor: "pointer",
    fontWeight: 600,
  },
  allowBtn: {
    height: "40px",
    minWidth: "110px",
    borderRadius: "12px",
    border: "none",
    background: "#2EA44F",
    color: "white",
    cursor: "pointer",
    fontWeight: 700,
  },
});

function messageParent(payload: unknown) {
  const msg = JSON.stringify(payload);

  try {
    if (typeof Office !== "undefined" && Office?.context?.ui?.messageParent) {
      Office.context.ui.messageParent(msg);
      return;
    }
  } catch {
    // ignore
  }

  alert("Office dialog messaging is not available.");
}

function getQueryParam(name: string): string {
  try {
    const url = new URL(window.location.href);
    return url.searchParams.get(name) || "";
  } catch {
    return "";
  }
}

export default function DialogAuth() {
  const styles = useStyles();

  const [workspaceHost, setWorkspaceHost] = React.useState("");
  const [token, setToken] = React.useState("");
  const [selectedWorkspace, setSelectedWorkspace] = React.useState<Workspace | null>(null);
  const [authedEmail, setAuthedEmail] = React.useState<string>("");
  const [agreed, setAgreed] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    const wh = (getQueryParam("workspaceHost") || "").trim();
    const tk = (getQueryParam("token") || "").trim();

    setWorkspaceHost(wh);
    setToken(tk);

    if (wh) {
      setSelectedWorkspace({
        workspaceId: wh,
        name: wh,
        host: wh,
      });
    }
  }, []);

  const onAllow = () => {
    setError(null);

    if (!workspaceHost) {
      setError("Workspace is missing.");
      return;
    }
    if (!token) {
      setError("Token is missing.");
      return;
    }
    if (!selectedWorkspace) {
      setError("Workspace is missing.");
      return;
    }
    if (!agreed) {
      setError("You must confirm authorisation to continue.");
      return;
    }

    messageParent({
      type: "auth_complete",
      token,
      email: authedEmail || "",
      workspace: selectedWorkspace,
      agreed: true,
    });
  };

  const onCancel = () => {
    messageParent({ type: "auth_cancel" });
  };

  return (
    <div className={styles.page}>
      <div className={styles.card}>
        <div className={styles.consentTop}>
          <img src="assets/outlooktosc.png" alt="Outlook to SingleCase" className={styles.connectorImg} />
          <h2 className={styles.consentTitle}>
            Outlook is requesting permission to access the <span className={styles.consentTitleEm}>SingleCase</span>{" "}
            workspace
          </h2>
        </div>

        {selectedWorkspace ? (
          <div style={{ fontSize: "12px", opacity: 0.75, textAlign: "center", marginBottom: "10px" }}>
            Workspace: {selectedWorkspace.host}
          </div>
        ) : null}

        {error ? <div className={styles.error}>{error}</div> : null}

        <div className={styles.consentSection}>
          <div className={styles.sectionHeading}>What will this add in be able to view?</div>
          <div className={styles.list}>
            <div className={styles.row}>
              <div className={styles.iconBox}>👤</div>
              <div className={styles.rowText}>Content and info about you</div>
              <div className={styles.chevron}>›</div>
            </div>
            <div className={styles.row}>
              <div className={styles.iconBox}>💬</div>
              <div className={styles.rowText}>Content and info about emails and conversations</div>
              <div className={styles.chevron}>›</div>
            </div>
            <div className={styles.row}>
              <div className={styles.iconBox}>▦</div>
              <div className={styles.rowText}>Content and info about your workspace</div>
              <div className={styles.chevron}>›</div>
            </div>
          </div>
        </div>

        <div className={styles.consentSection}>
          <div className={styles.sectionHeading}>What will this add in be able to do?</div>
          <div className={styles.list}>
            <div className={styles.row}>
              <div className={styles.iconBox}>⚡</div>
              <div className={styles.rowText}>Upload selected emails and attachments to SingleCase cases</div>
              <div className={styles.chevron}>›</div>
            </div>
          </div>
        </div>

        <div className={styles.checkboxRow}>
          <input
            className={styles.checkbox}
            type="checkbox"
            checked={agreed}
            onChange={(e) => setAgreed(e.target.checked)}
          />

          <div>
            <p className={styles.agreementText}>
              I confirm that I have authorisation to connect this Outlook mailbox with SingleCase and upload emails and
              attachments to the selected workspace and case.
            </p>
          </div>
        </div>

        <div className={styles.consentActions}>
          <button className={styles.cancelBtn} onClick={onCancel} type="button">
            Cancel
          </button>
          <button className={styles.allowBtn} onClick={onAllow} type="button">
            Allow
          </button>
        </div>
      </div>
    </div>
  );
}
