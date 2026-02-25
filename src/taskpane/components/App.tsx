import * as React from "react";
import { makeStyles } from "@fluentui/react-components";

import AuthScreen from "./AuthScreen";
import MainWorkspace from "./MainWorkspace/MainWorkspace";
import WorkspaceSelectView from "./WorkspaceSelectView";
import TimesheetsPanel from "./TimesheetsPanel";
import SettingsModal, { AddinSettings } from "./SettingsModal";
import { getAuth, clearAuth, clearAuthIfExpired, setAuth } from "../../services/auth";
import { getStored, setStored } from "../../utils/storage";
import { STORAGE_KEYS } from "../../utils/constants";
import { loadSettings, saveSettings } from "../../utils/settingsStorage";
import QuickActionsPanel from "./QuickActionsPanel";
import { emit } from "../../telemetry/telemetry";

interface AppProps {
  title: string;
}

const DEFAULT_SETTINGS: AddinSettings = {
  caseListScope: "all",
  rememberLastCase: false,

  duplicates: "warn",
  filingOnSend: "warn",
  internalEmailHandling: "treatNormally",
};

type DialogPayload =
  | {
      type: "auth_complete";
      token: string;
      email: string;
      agreed: boolean;
      workspace?: { workspaceId: string; name: string; host: string };
    }
  | { type: "auth_cancel" };

type TabId = "cases" | "quick" | "timesheets" | "tasks";

const NAV_HEIGHT = 74;

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: "rgb(255, 255, 255)",
  },

  content: {
    flex: 1,
    padding: "12px",
    paddingBottom: `${12 + NAV_HEIGHT}px`,
    minHeight: 0,
  },

  topRow: {
    position: "relative",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: "10px",
    minHeight: "48px",
  },

  logoCenterWrap: {
    position: "absolute",
    left: "50%",
    transform: "translateX(-50%)",
  },

  topLeft: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    minWidth: 0,
  },
  logoCenter: {
    position: "absolute",
    left: "50%",
    transform: "translateX(-50%)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    pointerEvents: "none", // prevents blocking clicks on buttons
  },
  backBtn: {
    height: "40px",
    padding: "0 16px",
    borderRadius: "999px",
    border: "1px solid rgba(0,0,0,0.12)",
    backgroundColor: "rgba(255,255,255,0.92)",
    cursor: "pointer",
    fontSize: "14px",
    fontWeight: 700,
    color: "rgba(17,24,39,0.92)",
    boxShadow: "0 10px 26px rgba(0,0,0,0.06)",
    transitionProperty: "transform, background-color, border-color",
    transitionDuration: "120ms",
    transitionTimingFunction: "ease",

    ":hover": {
      backgroundColor: "rgba(255,255,255,1)",
      transform: "translateY(-1px)",
    },

    ":active": {
      transform: "translateY(0px)",
      backgroundColor: "rgba(255,255,255,0.94)",
    },

    ":focus-visible": {
      outlineStyle: "none",
      boxShadow: "0 0 0 3px rgba(0,32,74,0.18), 0 10px 26px rgba(0,0,0,0.06)",
    },
  },

  brandLogo: {
    height: "38px",
    width: "auto",
    objectFit: "contain",
    userSelect: "none",
  },

  settingsBtn: {
    width: "44px",
    height: "44px",
    borderRadius: "14px",
    border: "1px solid rgba(0,0,0,0.10)",
    backgroundColor: "rgba(255,255,255,0.92)",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    boxShadow: "0 10px 26px rgba(0,0,0,0.06)",
    transitionProperty: "transform, background-color, border-color",
    transitionDuration: "120ms",
    transitionTimingFunction: "ease",
    padding: 0,
    color: "#111827",

    ":hover": {
      backgroundColor: "rgba(255,255,255,1)",
      transform: "translateY(-1px)",
    },

    ":active": {
      transform: "translateY(0px)",
      backgroundColor: "rgba(255,255,255,0.94)",
    },

    ":focus-visible": {
      outlineStyle: "none",
      boxShadow: "0 0 0 3px rgba(0,32,74,0.18), 0 10px 26px rgba(0,0,0,0.06)",
    },
  },

  settingsIcon: {
    fontSize: "18px",
    lineHeight: 1,
    opacity: 0.9,
    userSelect: "none",
  },

  bottomNavWrap: {
    position: "fixed",
    left: 0,
    right: 0,
    bottom: 0,
    padding: "10px 12px 12px",
    zIndex: 50,
    pointerEvents: "none",
  },

  bottomNav: {
    pointerEvents: "auto",
    maxWidth: "720px",
    marginLeft: "auto",
    marginRight: "auto",
    backgroundColor: "rgba(255,255,255,0.92)",
    border: "1px solid rgba(0,0,0,0.08)",
    borderRadius: "18px",
    boxShadow: "0 14px 34px rgba(0,0,0,0.12)",
    height: `${NAV_HEIGHT}px`,
    display: "grid",
    gridTemplateColumns: "repeat(4, 1fr)",
    alignItems: "center",
    padding: "6px",
    gap: "6px",
  },

  navItem: {
    height: "100%",
    borderRadius: "14px",
    border: "1px solid transparent",
    background: "transparent",
    cursor: "pointer",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "6px",
    padding: 0,
    color: "rgba(17,24,39,0.65)",
    transitionProperty: "background-color, border-color, transform",
    transitionDuration: "120ms",
    transitionTimingFunction: "ease",

    ":hover": {
      backgroundColor: "rgba(0,0,0,0.04)",
      transform: "translateY(-1px)",
    },

    ":active": {
      transform: "translateY(0px)",
    },

    ":focus-visible": {
      outlineStyle: "none",
      boxShadow: "0 0 0 3px rgba(0,32,74,0.14)",
    },
  },

  navItemActive: {
    backgroundColor: "rgba(0,32,74,0.06)",
    color: "rgba(17,24,39,0.92)",
  },

  navIcon: {
    fontSize: "18px",
    lineHeight: 1,
    userSelect: "none",
  },

  navLabel: {
    fontSize: "12px",
    fontWeight: 700,
    lineHeight: 1,
  },
});

function BottomNav(props: { active: TabId; onSelect: (id: TabId) => void }) {
  const styles = useStyles();

  const items: Array<{ id: TabId; icon: string; label: string }> = [
    { id: "quick", icon: "⚡", label: "Quick Actions" },
  ];

  return (
    <div className={styles.bottomNavWrap}>
      <div
        className={styles.bottomNav}
        style={{ gridTemplateColumns: `repeat(${items.length}, 1fr)` }}
      >
        {items.map((it) => {
          const active = it.id === props.active;
          return (
            <button
              key={it.id}
              type="button"
              className={[styles.navItem, active ? styles.navItemActive : ""]
                .filter(Boolean)
                .join(" ")}
              onClick={() => props.onSelect(it.id)}
              aria-current={active ? "page" : undefined}
            >
              <span className={styles.navIcon} aria-hidden="true">
                {it.icon}
              </span>
              <span className={styles.navLabel}>{it.label}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
}

const App: React.FC<AppProps> = (props) => {
  const styles = useStyles();

  const [forceWorkspaceSelect, setForceWorkspaceSelect] = React.useState(false);
  const [auth, setAuthState] = React.useState(() => getAuth());
  const [workspaceId, setWorkspaceId] = React.useState<string | null>(null);

  const [bootstrapped, setBootstrapped] = React.useState(false);

  const [settings, setSettings] = React.useState<AddinSettings>(() =>
    loadSettings(DEFAULT_SETTINGS)
  );
  const [isSettingsOpen, setIsSettingsOpen] = React.useState(false);

  const [activeTab, setActiveTab] = React.useState<TabId>("cases");

  // home = MainWorkspace, app = any screen with bottom nav
  const [postLoginStep, setPostLoginStep] = React.useState<"home" | "app">("home");

  const dialogRef = React.useRef<Office.Dialog | null>(null);

  React.useEffect(() => {
    clearAuthIfExpired();
    setAuthState(getAuth());
  }, []);

  React.useEffect(() => {
    (async () => {
      const newToken = process.env.SINGLECASE_PUBLIC_TOKEN;

      if (!newToken) {
        console.error("SINGLECASE_PUBLIC_TOKEN is not defined in environment variables");
        return;
      }

      await setStored(STORAGE_KEYS.publicToken, newToken);

      setAuth(newToken, getAuth().email || "unknown@singlecase.local");
      setAuthState(getAuth());
    })();
  }, []);

  React.useEffect(() => {
    (async () => {
      if (!auth.token) {
        setWorkspaceId(null);
        setForceWorkspaceSelect(false);
        setPostLoginStep("home");
        setBootstrapped(true);
        return;
      }

      if (forceWorkspaceSelect) {
        setWorkspaceId(null);
        setPostLoginStep("home");
        setBootstrapped(true);
        return;
      }

      const storedWorkspaceIdRaw = await getStored(STORAGE_KEYS.workspaceId);
      const storedWorkspaceId =
        storedWorkspaceIdRaw && storedWorkspaceIdRaw.length > 0 ? storedWorkspaceIdRaw : null;

      setWorkspaceId(storedWorkspaceId);

      if (!storedWorkspaceId) {
        setPostLoginStep("home");
        setBootstrapped(true);
        return;
      }

      // Always start on MainWorkspace when taskpane opens
      setPostLoginStep("home");
      setBootstrapped(true);
    })();
  }, [auth.token, forceWorkspaceSelect]);

  const closeDialogIfOpen = () => {
    try {
      dialogRef.current?.close();
    } catch {
      // ignore
    } finally {
      dialogRef.current = null;
    }
  };

  const handleAuthCompleted = async (payload: {
    token: string;
    email: string;
    agreed: boolean;
    workspace?: { workspaceId: string; name: string; host: string };
  }) => {
if (!payload.agreed) return;

setIsSettingsOpen(false); // ADD THIS LINE
setAuth(payload.token, payload.email);
    emit("auth.refreshed", { reason: "login_completed" });
    await setStored(STORAGE_KEYS.agreementAccepted, "true");

    if (payload.workspace?.workspaceId) {
      await setStored(STORAGE_KEYS.workspaceId, payload.workspace.workspaceId);
      await setStored(STORAGE_KEYS.workspaceName, payload.workspace.name || "");
      await setStored(STORAGE_KEYS.workspaceHost, payload.workspace.host || "");

      setWorkspaceId(payload.workspace.workspaceId);
      setForceWorkspaceSelect(false);

      await setStored(STORAGE_KEYS.onboardingDone, "false");
      setPostLoginStep("home");
    } else {
      await setStored(STORAGE_KEYS.onboardingDone, "false");
      await setStored(STORAGE_KEYS.workspaceId, "");
      await setStored(STORAGE_KEYS.workspaceName, "");
      await setStored(STORAGE_KEYS.workspaceHost, "");

      setForceWorkspaceSelect(true);
      setWorkspaceId(null);
      setPostLoginStep("home");
    }

    setAuthState(getAuth());
  };

  const onOpenLoginDialog = async (args: { workspaceHost: string; token: string }) => {
    try {
      if (!Office?.context?.ui?.displayDialogAsync) {
        console.warn("Office UI dialog API not ready yet.");
        return;
      }

      const storedTokenRaw = await getStored(STORAGE_KEYS.publicToken);
      const storedToken = (storedTokenRaw || "").trim();
      if (!storedToken) {
        console.error("Missing sc:publicToken in storage.");
        return;
      }

      const params = new URLSearchParams();
      params.set("workspaceHost", args.workspaceHost);
      params.set("token", storedToken);

      const dialogUrl = `${window.location.origin}/dialog.html?${params.toString()}`;

      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 70, width: 45, displayInIframe: false },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("displayDialogAsync failed", result.error);
            return;
          }

          const dialog = result.value;
          dialogRef.current = dialog;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg: any) => {
            if (!arg || typeof arg.message !== "string") return;

            let data: DialogPayload | any;
            try {
              data = JSON.parse(arg.message);
            } catch {
              return;
            }

            if (data?.type === "auth_cancel") {
              closeDialogIfOpen();
              return;
            }

            if (data?.type === "auth_complete") {
              await handleAuthCompleted({
                token: String(data.token || ""),
                email: String(data.email || ""),
                agreed: Boolean(data.agreed),
                workspace: data.workspace
                  ? {
                      workspaceId: String(data.workspace.workspaceId || ""),
                      name: String(data.workspace.name || ""),
                      host: String(data.workspace.host || ""),
                    }
                  : undefined,
              });
              closeDialogIfOpen();
            }
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            closeDialogIfOpen();
          });
        }
      );
    } catch (e) {
      console.error("Open dialog crashed", e);
    }
  };

  const onSignOut = async () => {
    clearAuth();

    await setStored(STORAGE_KEYS.onboardingDone, "false");
    await setStored(STORAGE_KEYS.workspaceId, "");
    await setStored(STORAGE_KEYS.workspaceName, "");
    await setStored(STORAGE_KEYS.workspaceHost, "");
    await setStored(STORAGE_KEYS.agreementAccepted, "false");

    setAuthState(getAuth());
    setWorkspaceId(null);
    setPostLoginStep("home");
    setBootstrapped(true);
  };

  const showShell = Boolean(auth.token) && Boolean(workspaceId);

  // Show global back button whenever you are in the "app" shell.
  // This guarantees "back to MainWorkspace from anywhere".
  const showBackToHome = showShell && postLoginStep === "app";

  if (!bootstrapped) {
    return <div className={styles.root} />;
  }

  return (
    <div className={styles.root}>
      {!auth.token ? (
        <AuthScreen
          title={props.title}
          onOpenLoginDialog={onOpenLoginDialog}
          onAuthCompleted={() => {
            // not used by launcher
          }}
        />
      ) : (
        <>
          <SettingsModal
            isOpen={isSettingsOpen}
            settings={settings}
            onClose={() => setIsSettingsOpen(false)}
            onChange={(s) => {
              const changedKeys = (Object.keys(s) as (keyof AddinSettings)[]).filter(
                (k) => s[k] !== settings[k]
              );
              saveSettings(s);
              setSettings(s);
              emit("settings.changed", { changedKeys });
            }}
            onReset={() => { saveSettings(DEFAULT_SETTINGS); setSettings(DEFAULT_SETTINGS); }}
            onSignOut={onSignOut}
          />

          <div className={styles.content}>
            {showShell ? (
              <div className={styles.topRow}>
                <div className={styles.topLeft}>
                  {showBackToHome ? (
                    <button
                      type="button"
                      className={styles.backBtn}
                      onClick={() => setPostLoginStep("home")}
                    >
                      ‹ Back
                    </button>
                  ) : (
                    <div />
                  )}
                </div>

                <div className={styles.logoCenter}>
                  <img
                    src="assets/sc.png"
                    alt="SingleCase"
                    className={styles.brandLogo}
                  />
                </div>

                <button
                  type="button"
                  className={styles.settingsBtn}
                  onClick={() => setIsSettingsOpen(true)}
                  aria-label="Nastavení"
                  title="Nastavení"
                >
                  <span className={styles.settingsIcon} aria-hidden="true">
                    ⚙
                  </span>
                </button>
              </div>
            ) : null}

            {!workspaceId ? (
              <WorkspaceSelectView
                email={auth.email as string}
                onSelect={async (w) => {
                  await setStored(STORAGE_KEYS.workspaceId, w.workspaceId);
                  await setStored(STORAGE_KEYS.workspaceName, w.name);
                  await setStored(STORAGE_KEYS.workspaceHost, w.host);

                  setWorkspaceId(w.workspaceId);
                  setForceWorkspaceSelect(false);

                  const onboardingDone = (await getStored(STORAGE_KEYS.onboardingDone)) === "true";
                  setPostLoginStep(onboardingDone ? "app" : "home");
                }}
              />
            ) : postLoginStep === "home" ? (
              <MainWorkspace
                email={auth.email as string}
                token={auth.token as string}
                settings={settings}
                onChangeSettings={setSettings}
                onSignOut={onSignOut}
                onOpenTab={(tab) => {
                  setPostLoginStep("app");
                  setActiveTab(tab);
                }}
              />
            ) : activeTab === "timesheets" ? (
              <TimesheetsPanel
                token={auth.token as string}
                onBack={() => setPostLoginStep("home")}
              />
              ) : activeTab === "quick" ? (
                <QuickActionsPanel
                  onOpenTab={(tab) => {
                    setPostLoginStep("app");
                    setActiveTab(tab as any);
                  }}
                />
              ) : null}
          </div>

          {showShell ? (
            <BottomNav
              active={activeTab}
              onSelect={(id) => {
                setPostLoginStep("app");
                setActiveTab(id);
              }}
            />
          ) : null}
        </>
      )}
    </div>
  );
};

export default App;
