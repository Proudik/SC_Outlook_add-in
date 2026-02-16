import * as React from "react";
import "./EmailContextPanel.css";

type Props = {
  emailError?: string | null;

  emailItemId: string;
  fromName: string;
  fromEmail: string;
  subject: string;

  isSwitchingEmail: boolean;
};

export default function EmailContextPanel({
  emailError,
  emailItemId,
  fromName,
  fromEmail,
  subject,
  isSwitchingEmail,
}: Props) {
  const [expanded, setExpanded] = React.useState(false);

  // Reset collapse when user switches to a different email
  React.useEffect(() => {
    setExpanded(false);
  }, [emailItemId]);

  if (emailError) {
    return <div className="email-context-panel-error">{emailError}</div>;
  }

  if (!emailItemId) {
    return <div className="email-context-panel-empty">Open an email first.</div>;
  }

  const fromText = isSwitchingEmail
    ? "Updating…"
    : `${fromName}${fromEmail ? ` (${fromEmail})` : ""}`;

  return (
    <button
      type="button"
      className={["email-context-compact", expanded ? "email-context-compact--expanded" : ""]
        .filter(Boolean)
        .join(" ")}
      onClick={() => setExpanded((v) => !v)}
      aria-expanded={expanded}
    >
      <div className="email-context-compact-left">
        <div className="email-context-compact-row">
          <span className="email-context-compact-label">From:</span>
          <span className="email-context-compact-value">{fromText}</span>
        </div>

        {expanded ? (
          <div className="email-context-compact-row email-context-compact-row--secondary">
            <span className="email-context-compact-label">Subject:</span>
            <span className="email-context-compact-value">{isSwitchingEmail ? "Updating…" : subject}</span>
          </div>
        ) : null}
      </div>

      <div className="email-context-compact-chevron" aria-hidden="true">
        {expanded ? "⌃" : "⌄"}
      </div>
    </button>
  );
}
