import * as React from "react";
import "./AttachmentsPicker.css";

export type OutlookAttachment = {
  id: string;
  name: string;
  size: number;
  isInline: boolean;
  contentType?: string;
};

type Props = {
  enabled: boolean;

  attachments: OutlookAttachment[];
  selectedAttachmentIds: Set<string>;
  isLoadingAttachments: boolean;

  onToggleAttachment: (id: string) => void;
  onSelectAll: () => void;
  onClearAll: () => void;
};

export default function AttachmentsPicker({
  enabled,
  attachments,
  selectedAttachmentIds,
  isLoadingAttachments,
  onToggleAttachment,
  onSelectAll,
  onClearAll,
}: Props) {
  if (!enabled) return null;

  return (
    <div className="attachments-picker-wrap">
      <div className="attachments-picker-top">
        <div className="attachments-picker-section-title">Attachments</div>

        {attachments.length ? (
          <div className="attachments-picker-actions">
            <button
              type="button"
              onClick={onSelectAll}
              className="attachments-picker-small-btn"
              disabled={isLoadingAttachments}
            >
              Select all
            </button>
            <button
              type="button"
              onClick={onClearAll}
              className="attachments-picker-small-btn"
              disabled={isLoadingAttachments}
            >
              Clear
            </button>
          </div>
        ) : null}
      </div>

      {isLoadingAttachments ? (
        <div className="attachments-picker-loading">Loading attachments...</div>
      ) : attachments.length === 0 ? (
        <div className="attachments-picker-empty">No attachments detected.</div>
      ) : (
        <div className="attachments-picker-list">
          {attachments.map((a) => {
            const checked = selectedAttachmentIds.has(a.id);

            return (
              <label
                key={a.id}
                className="attachments-picker-item"
              >
                <input type="checkbox" checked={checked} onChange={() => onToggleAttachment(a.id)} />
                <div className="attachments-picker-item-body">
                  <div
                    className="attachments-picker-item-name"
                    title={a.name}
                  >
                    {a.name}
                  </div>

                  <div
                    className="attachments-picker-item-meta"
                    title={`${a.size ? `${Math.round(a.size / 1024)} KB` : "Size unknown"}${
                      a.contentType ? ` · ${a.contentType}` : ""
                    }`}
                  >
                    {a.size ? `${Math.round(a.size / 1024)} KB` : "Size unknown"}
                    {a.contentType ? ` · ${a.contentType}` : ""}
                  </div>
                </div>
              </label>
            );
          })}
        </div>
      )}
    </div>
  );
}
