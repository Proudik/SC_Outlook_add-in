import * as React from "react";
import { makeStyles } from "@fluentui/react-components";

type QuickActionId = "add_timesheet" | "add_task" | "create_note" | "summarise_email";

type Props = {
  onOpenTab?: (tab: "timesheets" | "tasks" | "cases" | "quick") => void;
};

type Msg = {
  id: string;
  role: "assistant" | "user";
  text: string;
};

const useStyles = makeStyles({
    page: {
        display: "flex",
        flexDirection: "column",
        height: "100%",
        minHeight: 0,
      },

  header: {
    marginBottom: "10px",
  },

  title: {
    margin: 0,
    fontSize: "18px",
    fontWeight: 800,
    color: "rgba(17,24,39,0.92)",
  },

  subtitle: {
    margin: "4px 0 0 0",
    fontSize: "13px",
    color: "rgba(17,24,39,0.62)",
    lineHeight: 1.35,
  },

  actionsDock: {
    position: "fixed",
    left: 0,
    right: 0,
  
    // This is the key number. 74 is your NAV_HEIGHT.
    // Add 12-16px for breathing room.
    bottom: "120px",
  
    padding: "10px 12px 12px",
    zIndex: 20,
  
    // optional, but usually looks better as a dock
    backgroundColor: "rgba(255,255,255,0.92)",
    borderTop: "1px solid rgba(0,0,0,0.06)",
    backdropFilter: "blur(8px)",
  },

  chat: {
    flex: 1,
    minHeight: 0,
    overflowY: "auto",
    padding: "6px 0 120px 0", // room so last bubble not hidden by dock
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },

  bubble: {
    maxWidth: "92%",
    borderRadius: "16px",
    padding: "10px 12px",
    border: "1px solid rgba(0,0,0,0.08)",
    boxShadow: "0 10px 24px rgba(0,0,0,0.06)",
    fontSize: "13px",
    lineHeight: 1.4,
    color: "rgba(17,24,39,0.92)",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },

  bubbleAssistant: {
    alignSelf: "flex-start",
    backgroundColor: "rgba(255,255,255,0.92)",
  },

  bubbleUser: {
    alignSelf: "flex-end",
    backgroundColor: "rgba(0,32,74,0.06)",
  },

  pillsWrap: {
    paddingTop: "10px",
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
  },

  pill: {
    height: "34px",
    padding: "0 12px",
    borderRadius: "999px",
    border: "1px solid rgba(0,0,0,0.12)",
    backgroundColor: "rgba(255,255,255,0.92)",
    cursor: "pointer",
    fontSize: "12px",
    fontWeight: 800,
    color: "rgba(17,24,39,0.86)",
    boxShadow: "0 10px 24px rgba(0,0,0,0.06)",
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
      boxShadow: "0 0 0 3px rgba(0,32,74,0.16), 0 10px 24px rgba(0,0,0,0.06)",
    },
  },

  footerHint: {
    marginTop: "10px",
    fontSize: "12px",
    color: "rgba(17,24,39,0.55)",
    lineHeight: 1.35,
  },
});

function uid() {
  return `${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function labelForAction(id: QuickActionId): string {
  switch (id) {
    case "add_timesheet":
      return "Add timesheet";
    case "add_task":
      return "Add task";
    case "create_note":
      return "Create note";
    case "summarise_email":
      return "Summarise email";
    default:
      return "Action";
  }
}

function assistantReplyForAction(id: QuickActionId): string {
  switch (id) {
    case "add_timesheet":
      return "OK. This will create a timesheet entry.\n\nNext step: choose case, date, and duration (prototype only).";
    case "add_task":
      return "OK. This will create a task.\n\nNext step: pick case, assignee, due date (prototype only).";
    case "create_note":
      return "OK. This will create a note linked to the current case or email.\n\nNext step: enter the note text (prototype only).";
    case "summarise_email":
      return "OK. This will summarise the currently selected email.\n\nNext step: generate a short summary and key actions (prototype only).";
    default:
      return "OK.";
  }
}

export default function QuickActionsPanel(props: Props) {
  const styles = useStyles();

  const [messages, setMessages] = React.useState<Msg[]>([]);

  const chatEndRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    if (messages.length === 0) return;
    chatEndRef.current?.scrollIntoView({ block: "end", behavior: "smooth" });
  }, [messages.length]);

  const runAction = (actionId: QuickActionId) => {
    const userText = labelForAction(actionId);

    setMessages((prev) => [
      ...prev,
      { id: uid(), role: "user", text: userText },
      { id: uid(), role: "assistant", text: assistantReplyForAction(actionId) },
    ]);

    // Optional: jump to real screens for the actions you already have
    if (actionId === "add_timesheet") {
      props.onOpenTab?.("timesheets");
    }
  };

  return (
    <div className={styles.page}>
      <div className={styles.header}>
        <p className={styles.title}>Quick Actions</p>
      </div>

      <div className={styles.chat} role="log" aria-label="Quick actions chat">
  {messages.length > 0
    ? messages.map((m) => (
        <div
          key={m.id}
          className={[
            styles.bubble,
            m.role === "assistant" ? styles.bubbleAssistant : styles.bubbleUser,
          ].join(" ")}
        >
          {m.text}
        </div>
      ))
    : null}

  <div ref={chatEndRef} />
</div>

      <div className={styles.actionsDock}>
  <div className={styles.pillsWrap} aria-label="Quick action options">
    <button type="button" className={styles.pill} onClick={() => runAction("add_timesheet")}>
      Add timesheet
    </button>
    <button type="button" className={styles.pill} onClick={() => runAction("add_task")}>
      Add task
    </button>
    <button type="button" className={styles.pill} onClick={() => runAction("create_note")}>
      Create note
    </button>
    <button type="button" className={styles.pill} onClick={() => runAction("summarise_email")}>
      Summarise email
    </button>
  </div>

</div>

    
    </div>
  );
}