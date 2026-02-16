import * as React from "react";

export type AttachmentLike = {
  id: string;
  name: string;
  size?: number;
  isInline?: boolean;
};

type Props = {
  attachmentsLite: AttachmentLike[];
  attachmentIds: string[];
  selectedAttachments: string[];
  onSelectionChange: (ids: string[]) => void;
  filingMode: "attachments" | "both";
  onFilingModeChange: (mode: "attachments" | "both") => void;
  containerRef: React.RefObject<HTMLDivElement | null>;
};

export default function AttachmentsStep({
  attachmentsLite,
  attachmentIds,
  selectedAttachments,
  onSelectionChange,
  filingMode,
  onFilingModeChange,
  containerRef,
}: Props) {
  return (
    <div ref={containerRef} className="mwAttachmentsStep">
      <div className="mwAttachmentsStepTitle">Co chcete uložit?</div>

      <div className="mwAttachmentsStepButtons">
        <button
          type="button"
          className={`mwFilingBtn ${filingMode === "attachments" ? "mwFilingBtnActive" : ""}`}
          onClick={() => {
            onFilingModeChange("attachments");
            onSelectionChange(attachmentIds);
          }}
        >
          Pouze přílohy
        </button>

        <button
          type="button"
          className={`mwFilingBtn ${filingMode === "both" ? "mwFilingBtnActive" : ""}`}
          onClick={() => {
            onFilingModeChange("both");
            onSelectionChange(attachmentIds);
          }}
        >
          Email i přílohy
        </button>
      </div>

      <div className="mwAttachmentsStepList">
        {attachmentsLite.map((att) => {
          const id = String(att.id);
          return (
            <label key={id} className="mwAttachmentsStepLabel">
              <input
                type="checkbox"
                checked={selectedAttachments.includes(id)}
                onChange={(e) => {
                  onSelectionChange(
                    e.target.checked ? [...selectedAttachments, id] : selectedAttachments.filter((x) => x !== id)
                  );
                }}
              />
              {att.name}
            </label>
          );
        })}
      </div>
    </div>
  );
}
