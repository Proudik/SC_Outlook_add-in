import * as React from "react";

export type CaseListScope = "favourites" | "all";
export type DuplicatesHandling = "off" | "warn" | "block";
export type FilingOnSend = "off" | "warn" | "always";
export type InternalEmailHandling = "treatNormally" | "doNotSuggest";

export type AddinSettings = {
  caseListScope: CaseListScope;
  rememberLastCase: boolean;

  duplicates: DuplicatesHandling;
  filingOnSend: FilingOnSend;
  internalEmailHandling: InternalEmailHandling;
};

type Props = {
  isOpen: boolean;
  settings: AddinSettings;
  onClose: () => void;
  onChange: (s: AddinSettings) => void;
  onReset: () => void;

  onSignOut?: () => void;
};


export default function SettingsModal(props: Props) {
  const { isOpen, settings, onClose, onChange, onReset } = props;

  React.useEffect(() => {
    if (!isOpen) return undefined;

    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  const row: React.CSSProperties = {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 12,
  };

  const label: React.CSSProperties = {
    fontSize: 12,
    opacity: 0.85,
    lineHeight: "1.3",
  };

  const help: React.CSSProperties = {
    fontSize: 12,
    opacity: 0.65,
    marginTop: 4,
    lineHeight: "1.35",
  };

  const toggle: React.CSSProperties = {
    width: 18,
    height: 18,
  };

  const selectStyle: React.CSSProperties = {
    height: 32,
    borderRadius: 10,
    border: "1px solid rgba(0,0,0,0.12)",
    padding: "0 10px",
    background: "rgba(255,255,255,0.9)",
    fontSize: 12,
  };

  const card: React.CSSProperties = {
    borderRadius: 12,
    border: "1px solid rgba(0,0,0,0.08)",
    padding: 12,
  };

  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.25)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 16,
        zIndex: 9999,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          width: "100%",
          maxWidth: 420,
          borderRadius: 14,
          border: "1px solid rgba(0,0,0,0.10)",
          background: "rgba(255,255,255,0.98)",
          boxShadow: "0 10px 30px rgba(0,0,0,0.18)",
          padding: 14,
        }}
      >
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            gap: 12,
          }}
        >
          <div style={{ fontWeight: 600 }}>Settings</div>
          <button
            onClick={onClose}
            style={{
              border: "none",
              background: "transparent",
              cursor: "pointer",
              fontSize: 16,
              lineHeight: 1,
              padding: 6,
              opacity: 0.75,
            }}
            aria-label="Close"
            type="button"
          >
            ×
          </button>
        </div>

        <div style={{ marginTop: 12, display: "flex", flexDirection: "column", gap: 12 }}>

          {/* ── Case display ── */}
          <div style={{ ...card, background: "rgba(17,24,39,0.02)" }}>
            <div style={row}>
              <div>
                <div style={label}>Default case group</div>
                <div style={help}>
                  Which case group is shown by default when opening the add-in.
                </div>
              </div>
              <select
                style={selectStyle}
                value={settings.caseListScope}
                onChange={(e) =>
                  onChange({ ...settings, caseListScope: e.target.value as CaseListScope })
                }
              >
                <option value="favourites">Favourites</option>
                <option value="all">All</option>
              </select>
            </div>

            <div style={{ marginTop: 10, ...row }}>
              <div>
                <div style={label}>Remember last selected case</div>
                <div style={help}>Start with your last case when opening new emails.</div>
              </div>
              <input
                style={toggle}
                type="checkbox"
                checked={settings.rememberLastCase}
                onChange={(e) => onChange({ ...settings, rememberLastCase: e.target.checked })}
              />
            </div>
          </div>

          {/* ── Filing behaviour ── */}
          <div style={card}>
            <div style={row}>
              <div>
                <div style={label}>File on send</div>
                <div style={help}>
                  {settings.filingOnSend === "warn"
                    ? "Shows a warning before you send. After sending, you can choose whether to file to SingleCase."
                    : settings.filingOnSend === "always"
                    ? "Files this email to SingleCase automatically when you press Send."
                    : "Auto filing is disabled. Emails will not be filed when you press Send."}
                </div>
              </div>
              <select
                style={selectStyle}
                value={settings.filingOnSend}
                onChange={(e) =>
                  onChange({ ...settings, filingOnSend: e.target.value as FilingOnSend })
                }
              >
                <option value="off">Off</option>
                <option value="warn">Warn each time</option>
                <option value="always">Always file</option>
              </select>
            </div>

            <div style={{ marginTop: 10, ...row }}>
              <div>
                <div style={label}>Duplicate handling</div>
                <div style={help}>
                  What to do when the same email was already filed to the same case.
                </div>
              </div>
              <select
                style={selectStyle}
                value={settings.duplicates}
                onChange={(e) =>
                  onChange({ ...settings, duplicates: e.target.value as DuplicatesHandling })
                }
              >
                <option value="off">Off</option>
                <option value="warn">Warn</option>
                <option value="block">Block</option>
              </select>
            </div>
          </div>

          {/* ── Internal emails ── */}
          <div style={card}>
            <div style={row}>
              <div>
                <div style={label}>Internal email handling</div>
                <div style={help}>
                  When on, emails where everyone shares your domain show an info message and skip case suggestions.
                </div>
              </div>
              <input
                style={toggle}
                type="checkbox"
                checked={settings.internalEmailHandling === "doNotSuggest"}
                onChange={(e) =>
                  onChange({
                    ...settings,
                    internalEmailHandling: e.target.checked ? "doNotSuggest" : "treatNormally",
                  })
                }
              />
            </div>
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", gap: 10 }}>
            {props.onSignOut ? (
              <button
                onClick={props.onSignOut}
                style={{
                  height: 34,
                  borderRadius: 10,
                  border: "1px solid rgba(220,38,38,0.25)",
                  padding: "0 12px",
                  background: "rgba(220,38,38,0.06)",
                  color: "rgb(153,27,27)",
                  cursor: "pointer",
                  fontWeight: 600,
                }}
                type="button"
              >
                Sign out
              </button>
            ) : (
              <div />
            )}

            <div style={{ display: "flex", gap: 10 }}>
              <button
                onClick={onReset}
                style={{
                  height: 34,
                  borderRadius: 10,
                  border: "1px solid rgba(0,0,0,0.12)",
                  padding: "0 12px",
                  background: "transparent",
                  cursor: "pointer",
                }}
                type="button"
              >
                Reset
              </button>
              <button
                onClick={onClose}
                style={{
                  height: 34,
                  borderRadius: 10,
                  border: "none",
                  padding: "0 12px",
                  background: "#111827",
                  color: "white",
                  cursor: "pointer",
                }}
                type="button"
              >
                Done
              </button>
            </div>
          </div>

        </div>
      </div>
    </div>
  );
}
