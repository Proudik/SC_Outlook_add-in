import * as React from "react";
import { makeStyles } from "@fluentui/react-components";

type Props = {
  label: string;
  tone?: "success" | "info" | "neutral";
  title?: string;

  // Optional: show version like v3
  version?: number;
};

const useStyles = makeStyles({
  pill: {
    height: "24px",
    borderRadius: "999px",
    padding: "0 10px",
    display: "inline-flex",
    alignItems: "center",
    gap: "8px",
    fontSize: "12px",
    fontWeight: 700,
    border: "1px solid rgba(0,0,0,0.10)",
    background: "rgba(255,255,255,0.6)",
    color: "rgba(17,24,39,0.9)",
    userSelect: "none",
    whiteSpace: "nowrap",
  },

  dot: {
    width: "8px",
    height: "8px",
    borderRadius: "999px",
    background: "rgba(0,0,0,0.35)",
    flex: "0 0 auto",
  },

  versionBadge: {
    marginLeft: "4px",
    height: "18px",
    borderRadius: "999px",
    padding: "0 8px",
    display: "inline-flex",
    alignItems: "center",
    fontSize: "11px",
    fontWeight: 800,
    border: "1px solid rgba(0,0,0,0.10)",
    background: "rgba(255,255,255,0.55)",
    opacity: 0.9,
  },

  success: {
    border: "1px solid rgba(16,185,129,0.28)",
    background: "rgba(16,185,129,0.10)",
    color: "rgb(6,95,70)",
  },
  successDot: { background: "rgb(16,185,129)" },

  info: {
    border: "1px solid rgba(37,99,235,0.25)",
    background: "rgba(37,99,235,0.08)",
    color: "rgba(29,78,216,0.95)",
  },
  infoDot: { background: "rgba(37,99,235,0.95)" },
});

export default function StatusPill({ label, tone = "neutral", title, version }: Props) {
  const s = useStyles();

  const toneClass = tone === "success" ? s.success : tone === "info" ? s.info : undefined;
  const dotClass = tone === "success" ? s.successDot : tone === "info" ? s.infoDot : s.dot;

  const showVersion = typeof version === "number" && version >= 2;

  return (
    <span className={[s.pill, toneClass].filter(Boolean).join(" ")} title={title}>
      <span className={dotClass} />
      <span>{label}</span>
      {showVersion ? <span className={s.versionBadge}>{`v${version}`}</span> : null}
    </span>
  );
}
