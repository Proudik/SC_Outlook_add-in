import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import Header from "./Header";
import { getStored, setStored } from "../../utils/storage";
import { STORAGE_KEYS } from "../../utils/constants";

type Props = {
  title: string;
  onAuthCompleted: (payload: { token: string; email: string; agreed: boolean }) => void;
  onOpenLoginDialog: (args: { workspaceHost: string; token: string }) => void;
};

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: "rgb(240, 240, 240)",
  },
  content: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "24px 16px 0",
    boxSizing: "border-box",
  },
  headerWrap: {
    marginTop: "36px",
    display: "flex",
    justifyContent: "center",
    width: "100%",
  },
  scaleImage: {
    width: "100%",
    maxWidth: "320px",
    margin: "64px auto 24px",
    opacity: 0.6,
    display: "block",
    pointerEvents: "none",
    userSelect: "none",
  },

  actionsWrap: {
    position: "sticky",
    bottom: "20px",
    width: "100%",
    display: "flex",
    justifyContent: "center",
  },
  actionsInner: {
    width: "100%",
    maxWidth: "420px",
    padding: "0 16px",
    display: "flex",
    flexDirection: "column",
    gap: "14px",
  },

  formWrap: {
    width: "100%",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },
  label: { fontSize: "12px", opacity: 0.8 },
  input: {
    height: "40px",
    borderRadius: "12px",
    border: "1px solid rgba(0,0,0,0.12)",
    padding: "0 12px",
    outline: "none",
    background: "rgba(255,255,255,0.9)",
    boxSizing: "border-box",
    width: "100%",
  },

  error: {
    borderRadius: "10px",
    padding: "10px 12px",
    border: "1px solid rgba(180, 0, 0, 0.25)",
    background: "rgba(220, 0, 0, 0.06)",
    color: "rgba(120, 0, 0, 0.95)",
    fontSize: "13px",
    lineHeight: 1.35,
  },

  primaryBtn: {
    height: "44px",
    borderRadius: "12px",
    border: "none",
    padding: "0 16px",
    background: "#00204A",
    color: "white",
    cursor: "pointer",
    fontWeight: 600,
    width: "100%",
  },
  secondaryBtn: {
    height: "44px",
    borderRadius: "12px",
    border: "1px solid rgba(0,0,0,0.12)",
    padding: "0 16px",
    background: "transparent",
    cursor: "pointer",
    fontWeight: 500,
    width: "100%",
  },

  legalRow: {
    marginTop: "8px",
    fontSize: "12px",
    lineHeight: "1.35",
    opacity: 0.75,
    textAlign: "center",
  },
  legalLink: {
    border: "none",
    background: "transparent",
    padding: 0,
    cursor: "pointer",
    textDecoration: "underline",
    fontSize: "12px",
    fontWeight: 700,
    color: "#00204A",
  },
});

function normalizeWorkspaceHost(input: string): string {
  const raw = (input || "").trim();
  if (!raw) return "";

  let v = raw;
  v = v.replace(/^https?:\/\//i, "");
  v = v.split("/")[0];
  v = v.split("?")[0];
  v = v.split("#")[0];

  if (v && !v.includes(".")) {
    v = `${v}.singlecase.com`;
  }

  return v.toLowerCase();
}

export default function AuthScreen({ title, onOpenLoginDialog }: Props) {
  const styles = useStyles();

  const [workspaceHostInput, setWorkspaceHostInput] = React.useState("");
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    (async () => {
      const storedHost = await getStored(STORAGE_KEYS.workspaceHost);
      if (storedHost) setWorkspaceHostInput(storedHost);
    })();
  }, []);

  const workspaceHost = normalizeWorkspaceHost(workspaceHostInput);
  const canContinue = workspaceHost.length > 0;

  const openWorkspaceWebsite = () => {
    const url = workspaceHost ? `https://${workspaceHost}` : "https://singlecase.com";
    const ui = (Office as any)?.context?.ui as any;

    if (ui?.openBrowserWindow) {
      ui.openBrowserWindow(url);
      return;
    }

    window.open(url, "_blank", "noopener,noreferrer");
  };

  const openLegal = (url: string) => {
    const ui = (Office as any)?.context?.ui as any;

    if (ui?.openBrowserWindow) {
      ui.openBrowserWindow(url);
      return;
    }

    window.open(url, "_blank", "noopener,noreferrer");
  };

  const onContinue = async () => {
    setError(null);

    if (!workspaceHost) {
      setError("Please enter a workspace URL.");
      return;
    }

    const storedToken = await getStored(STORAGE_KEYS.publicToken);
    const token = (storedToken || "").trim();
    if (!token) {
      setError("Missing public token in local storage.");
      return;
    }

    await setStored(STORAGE_KEYS.workspaceHost, workspaceHost);

    onOpenLoginDialog({ workspaceHost, token });
  };

  return (
    <div className={styles.root}>
      <div className={styles.content}>
        <div className={styles.headerWrap}>
          <Header logo="assets/SingleCaseFullLogo.svg" title={title} size="small" />
        </div>

        <img src="assets/Scale.webp" alt="" className={styles.scaleImage} />
      </div>

      <div className={styles.actionsWrap}>
        <div className={styles.actionsInner}>
          <div className={styles.formWrap}>
            <div>
              <div className={styles.label}>Workspace URL</div>
              <input
                className={styles.input}
                value={workspaceHostInput}
                onChange={(e) => setWorkspaceHostInput(e.target.value)}
                placeholder="acme.singlecase.com or acme"
                autoComplete="off"
                spellCheck={false}
              />
            </div>

            {error ? <div className={styles.error}>{error}</div> : null}
          </div>

          <button className={styles.primaryBtn} onClick={onContinue} disabled={!canContinue} type="button">
            Continue
          </button>

          <button className={styles.secondaryBtn} onClick={openWorkspaceWebsite} type="button">
            Open workspace
          </button>

          <div className={styles.legalRow}>
            <button type="button" className={styles.legalLink} onClick={() => openLegal("https://singlecase.com/terms")}>
              Terms
            </button>{" "}
            and{" "}
            <button type="button" className={styles.legalLink} onClick={() => openLegal("https://singlecase.com/privacy")}>
              Privacy Policy
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
