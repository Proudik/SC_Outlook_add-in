import * as React from "react";

export type PreparedTimerEntry = {
  note: string;
  seconds: number;
};

type Props = {
  onEntryChange: (entry: PreparedTimerEntry | null) => void;
};

/**
 * Strips non-digit characters from `value`, parses it as an integer, and
 * clamps the result to [min, max].
 *
 * - Empty / non-numeric input → { display: "", int: 0 }
 *   (blank field means 0 for calculation, per spec)
 * - Out-of-range values are clamped to the nearest bound.
 */
export function parseAndClampInt(
  value: string,
  min: number,
  max: number
): { display: string; int: number } {
  const digits = value.replace(/[^0-9]/g, "");
  if (!digits) return { display: "", int: 0 };
  const n = Math.min(max, Math.max(min, parseInt(digits, 10)));
  return { display: String(n), int: n };
}

/**
 * Returns today's date as "YYYY-MM-DD 00:00:00" (midnight, local time).
 */
export function formatDateMidnight(): string {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day} 00:00:00`;
}

// ─── Styles ────────────────────────────────────────────────────────────────

const panelStyle: React.CSSProperties = {
  marginTop: 8,
  padding: "10px 12px",
  borderRadius: 12,
  background: "rgba(16, 185, 129, 0.06)",
  border: "1px solid rgba(16, 185, 129, 0.2)",
  fontSize: 12,
};

const labelStyle: React.CSSProperties = {
  display: "block",
  fontWeight: 500,
  marginBottom: 3,
  marginTop: 6,
  color: "#333",
};

const inputStyle: React.CSSProperties = {
  width: "100%",
  boxSizing: "border-box",
  border: "1px solid rgba(0,0,0,0.15)",
  borderRadius: 6,
  padding: "5px 7px",
  fontSize: 12,
  fontFamily: "inherit",
  background: "white",
  textAlign: "center",
};

const subLabelStyle: React.CSSProperties = {
  fontSize: 10,
  color: "#999",
  textAlign: "center",
  marginTop: 2,
};

// ─── Component ─────────────────────────────────────────────────────────────

const TimeEntryOnSendPanel: React.FC<Props> = ({ onEntryChange }) => {
  const [description, setDescription] = React.useState("");
  const [hours, setHours] = React.useState("0");
  const [minutes, setMinutes] = React.useState("5");

  // Derive totalSeconds and propagate to parent on every field change
  React.useEffect(() => {
    const note = description.trim();
    const h = hours === "" ? 0 : Math.min(23, Math.max(0, parseInt(hours, 10) || 0));
    const m = minutes === "" ? 0 : Math.min(59, Math.max(0, parseInt(minutes, 10) || 0));
    const totalSeconds = h * 3600 + m * 60;

    if (note && totalSeconds >= 60) {
      onEntryChange({ note, seconds: totalSeconds });
    } else {
      onEntryChange(null);
    }
  // onEntryChange is a stable callback from parent — intentionally not in deps
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [description, hours, minutes]);

  // ── Derived state for UI feedback ────────────────────────────────────────
  const hoursInt = hours === "" ? 0 : Math.min(23, Math.max(0, parseInt(hours, 10) || 0));
  const minutesInt = minutes === "" ? 0 : Math.min(59, Math.max(0, parseInt(minutes, 10) || 0));
  const totalSeconds = hoursInt * 3600 + minutesInt * 60;
  const hasTime = hours !== "" || minutes !== "";
  const timeHint = hasTime && totalSeconds < 60 ? "Must be at least 1 minute" : null;
  const isReady = totalSeconds >= 60 && Boolean(description.trim());

  // ── Shared key handler — blocks decimal, sign, and exponent characters ───
  const blockInvalidKeys = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if ([".", "-", "e", "E", "+"].includes(e.key)) {
      e.preventDefault();
    }
  };

  // onChange: strip non-digit chars (handles paste of letters/symbols)
  const handleHoursChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setHours(e.target.value.replace(/[^0-9]/g, ""));
  };
  const handleMinutesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setMinutes(e.target.value.replace(/[^0-9]/g, ""));
  };

  // onBlur: clamp to valid range (e.g. "99" → "59" for minutes)
  const handleHoursBlur = (e: React.FocusEvent<HTMLInputElement>) => {
    setHours(parseAndClampInt(e.target.value, 0, 23).display);
  };
  const handleMinutesBlur = (e: React.FocusEvent<HTMLInputElement>) => {
    setMinutes(parseAndClampInt(e.target.value, 0, 59).display);
  };

  return (
    <div style={panelStyle}>
      <div style={{ fontWeight: 600, color: "#111", marginBottom: 2 }}>
        Log time to this case{" "}
        <span style={{ fontWeight: 400, color: "#888" }}>(optional)</span>
      </div>
      <div style={{ fontSize: 11, color: "#888", marginBottom: 8 }}>
        Fill in and click Send in Outlook to log time with this email.
      </div>

      <label style={labelStyle} htmlFor="te-description">
        Description
      </label>
      <textarea
        id="te-description"
        rows={3}
        style={{ ...inputStyle, resize: "vertical", textAlign: "left" }}
        placeholder="What did you work on?"
        value={description}
        onChange={(e) => setDescription(e.target.value)}
      />

      <label style={labelStyle}>Time</label>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
        <div>
          <input
            id="te-hours"
            type="number"
            min={0}
            max={23}
            step={1}
            style={inputStyle}
            placeholder="h"
            value={hours}
            onKeyDown={blockInvalidKeys}
            onChange={handleHoursChange}
            onBlur={handleHoursBlur}
          />
          <div style={subLabelStyle}>hours</div>
        </div>
        <div>
          <input
            id="te-minutes"
            type="number"
            min={0}
            max={59}
            step={1}
            style={inputStyle}
            placeholder="min"
            value={minutes}
            onKeyDown={blockInvalidKeys}
            onChange={handleMinutesChange}
            onBlur={handleMinutesBlur}
          />
          <div style={subLabelStyle}>minutes</div>
        </div>
      </div>

      {timeHint ? (
        <div style={{ fontSize: 10, color: "#d97706", marginTop: 4 }}>{timeHint}</div>
      ) : null}
      {isReady ? (
        <div style={{ fontSize: 10, color: "rgba(16, 185, 129, 0.9)", marginTop: 4 }}>
          ✓ Time entry ready — will be logged on Send
        </div>
      ) : null}
    </div>
  );
};

export default TimeEntryOnSendPanel;
