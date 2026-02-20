import * as React from "react";

type QuickAction = {
  id: string;
  label: string;
  onClick: () => void;
};

type Props = {
  text: string;
  isUnfiled: boolean;
  actions?: QuickAction[];
  tone?: "default" | "success" | "warning";
  note?: string;
};

export default function PromptBubble({ text, isUnfiled, actions, tone = "default", note }: Props) {
  let className = isUnfiled ? "mwChatBubble" : "mwChatMuted";
  if (tone === "success" && isUnfiled) {
    className = "mwChatBubbleSuccess";
  }
  if (tone === "warning" && isUnfiled) {
    className = "mwChatBubbleWarning";
  }

  return (
    <div className={className}>
      <div>
        {text || "Select an email and I’ll show you relevant suggestions."}
      </div>
      {note ? <div className="mwPromptNote">{note}</div> : null}

      {actions && actions.length > 0 && (
        <div className="mwQuickReplies">
          {actions.map((a) => (
            <button
              key={a.id}
              type="button"
              className="mwQuickReplyBtn"
              onClick={a.onClick}
            >
              {a.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}