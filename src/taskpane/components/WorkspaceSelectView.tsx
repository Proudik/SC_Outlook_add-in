import * as React from "react";
import { makeStyles, Button, Input } from "@fluentui/react-components";
import { getStored, setStored } from "../../utils/storage";

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
    background: "rgb(63, 74, 86)",
    boxSizing: "border-box",
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif',
  },

  shell: {
    width: "100%",
    maxWidth: "720px",
    marginLeft: "auto",
    marginRight: "auto",
    borderRadius: "16px",
    border: "1px solid rgba(0,0,0,0.10)",
    background: "white",
    boxShadow: "0 12px 30px rgba(0,0,0,0.12)",
    padding: "22px",
    boxSizing: "border-box",
  },

  brandRow: {
    display: "flex",
    justifyContent: "center",
    marginBottom: "14px",
  },

  logo: {
    height: "40px",
    width: "auto",
    display: "block",
  },

  title: {
    margin: 0,
    textAlign: "center",
    fontSize: "30px",
    lineHeight: 1.12,
    letterSpacing: "-0.02em",
    fontWeight: 800,
    color: "rgba(17,24,39,0.95)",
  },

  subtitle: {
    margin: "10px 0 0 0",
    textAlign: "center",
    fontSize: "13.5px",
    lineHeight: 1.45,
    opacity: 0.72,
  },

  formBlock: {
    marginTop: "18px",
    maxWidth: "640px",
    marginLeft: "auto",
    marginRight: "auto",
  },

  fieldLabel: {
    fontSize: "12.5px",
    fontWeight: 800,
    opacity: 0.85,
    marginBottom: "8px",
  },

  inputWrap: {
    borderRadius: "14px",
    border: "1px solid rgba(0,0,0,0.14)",
    background: "white",
    padding: "8px",
    boxShadow: "0 1px 0 rgba(0,0,0,0.02)",
  },

  input: {
    width: "100%",
    "& input::placeholder": {
      color: "rgba(0,0,0,0.35)",
    },
  },

  continueBtn: {
    marginTop: "12px",
    width: "100%",
    minHeight: "42px",
    borderRadius: "14px",
    fontWeight: 800,
  },

  errorBox: {
    marginTop: "12px",
    borderRadius: "12px",
    padding: "10px 12px",
    border: "1px solid rgba(180, 0, 0, 0.25)",
    background: "rgba(220, 0, 0, 0.05)",
    color: "rgba(120, 0, 0, 0.95)",
    fontSize: "12.5px",
    lineHeight: 1.35,
  },

  signedInTitle: {
    marginTop: "18px",
    textAlign: "center",
    fontSize: "13px",
    fontWeight: 900,
    opacity: 0.9,
  },

  connectedList: {
    marginTop: "12px",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    maxWidth: "640px",
    marginLeft: "auto",
    marginRight: "auto",
  },

  connectedCard: {
    borderRadius: "16px",
    border: "1px solid rgba(0,0,0,0.10)",
    background: "white",
    padding: "12px 14px",
    display: "flex",
    alignItems: "center",
    gap: "12px",
    boxShadow: "0 1px 0 rgba(0,0,0,0.02)",
    transition: "background 120ms ease",
    ":hover": {
      background: "rgba(0,0,0,0.015)",
    },
  },

  iconImg: {
    width: "40px",
    height: "40px",
    borderRadius: "14px",
    border: "1px solid rgba(0,0,0,0.10)",
    background: "white",
    padding: "8px",
    boxSizing: "border-box",
    flex: "0 0 auto",
  },

  cardMain: {
    flex: "1 1 auto",
    minWidth: 0,
  },

  hostLine: {
    fontSize: "14px",
    fontWeight: 900,
    color: "rgba(17,24,39,0.95)",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },

  nameLine: {
    marginTop: "3px",
    fontSize: "12.5px",
    opacity: 0.72,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },

  openBtn: {
    flex: "0 0 auto",
    minHeight: "36px",
    padding: "0 16px",
    borderRadius: "999px",
    fontWeight: 800,
    fontSize: "13px",
    background: "rgba(37, 99, 235, 0.08)",
    color: "rgba(37, 99, 235, 0.95)",
    border: "1px solid rgba(37, 99, 235, 0.25)",
    boxShadow: "0 1px 0 rgba(0,0,0,0.02)",
  },
});

const STORAGE_KEYS_LOCAL = {
  connectedWorkspaces: "sc:connectedWorkspaces",
} as const;

function normaliseHost(input: string): string {
  const raw = (input || "").trim();
  if (!raw) return "";

  const noProto = raw.replace(/^https?:\/\//i, "");
  const hostOnly = noProto.split(/[\/?#]/)[0].trim();

  // Do not auto append any domain.
  // You are switching between .singlecase.app and .singlecase.ch and others.
  // Force user to enter the real host.
  return hostOnly.toLowerCase();
}

function workspaceFromHost(host: string): Workspace {
  return {
    workspaceId: host,
    name: host,
    host,
  };
}

type Props = {
  email: string;
  onSelect: (workspace: Workspace) => void;
  onHelpClick?: () => void;
};

export default function WorkspaceSelectView({ email, onSelect }: Props) {
  const styles = useStyles();

  const [input, setInput] = React.useState("");
  const [error, setError] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);

  const [connected, setConnected] = React.useState<Workspace[]>([]);

  React.useEffect(() => {
    (async () => {
      const raw = await getStored(STORAGE_KEYS_LOCAL.connectedWorkspaces);
      if (!raw) return;
      try {
        const parsed = JSON.parse(raw) as Workspace[];
        if (Array.isArray(parsed) && parsed.length > 0) setConnected(parsed);
      } catch {
        // ignore
      }
    })();
  }, []);

  const continueWithHost = React.useCallback(
    async (rawInput: string) => {
      setError(null);

      const host = normaliseHost(rawInput);
      if (!host) return;

      setBusy(true);
      try {
        const w = workspaceFromHost(host);

        onSelect(w);

        const next = [w, ...connected.filter((x) => x.workspaceId !== w.workspaceId)].slice(0, 10);
        setConnected(next);
        await setStored(STORAGE_KEYS_LOCAL.connectedWorkspaces, JSON.stringify(next));
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      } finally {
        setBusy(false);
      }
    },
    [connected, onSelect]
  );

  const onContinue = () => {
    void continueWithHost(input.trim());
  };

  return (
    <div className={styles.page}>
      <div className={styles.shell}>
        <div className={styles.brandRow}>
          <img src="assets/SingleCaseFullLogo.svg" alt="SingleCase" className={styles.logo} />
        </div>

        <h1 className={styles.title}>Select your workspace</h1>
        <p className={styles.subtitle}>Signed in as {email}. Enter your SingleCase workspace host.</p>

        <div className={styles.formBlock}>
          <div className={styles.fieldLabel}>Workspace host</div>

          <div className={styles.inputWrap}>
            <Input
              value={input}
              onChange={(_, data) => setInput(data.value)}
              placeholder="valfor-demo.singlecase.ch"
              appearance="outline"
              className={styles.input}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !busy && input.trim().length > 0) {
                  onContinue();
                }
              }}
            />
          </div>

          <Button
            appearance="primary"
            onClick={onContinue}
            disabled={busy || input.trim().length === 0}
            className={styles.continueBtn}
          >
            {busy ? "Checking..." : "Continue"}
          </Button>

          {error ? <div className={styles.errorBox}>{error}</div> : null}
        </div>

        {connected.length > 0 ? (
          <>
            <div className={styles.signedInTitle}>Recent workspaces</div>

            <div className={styles.connectedList}>
              {connected.slice(0, 2).map((w) => (
                <div key={w.workspaceId} className={styles.connectedCard}>
                  <img src="assets/icon.svg" alt="" className={styles.iconImg} />

                  <div className={styles.cardMain}>
                    <div className={styles.hostLine}>{w.host}</div>
                    <div className={styles.nameLine}>{w.name}</div>
                  </div>

                  <Button appearance="secondary" onClick={() => onSelect(w)} className={styles.openBtn}>
                    Open
                  </Button>
                </div>
              ))}
            </div>
          </>
        ) : null}
      </div>
    </div>
  );
}
