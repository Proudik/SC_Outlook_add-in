import * as React from "react";
import type { SentPillData } from "../../../../utils/sentPillStore";
import DocumentHoverCard from "./DocumentHoverCard";

export type UploadedItem = {
  id: string;
  name: string;
  url: string;
  kind: "email" | "attachment";
  atIso: string;
  uploadedBy?: string;

  isLocked?: boolean;
  lockedBy?: string;
  lockedUntilIso?: string;
};

type Props = {
  caseUrl: string;
  filedCaseName: string;
  sentPill: SentPillData | null;
  documentsToShow: UploadedItem[];
  workspaceHost: string;
  onLockedDocAttempt: (msg: string) => void;
  onOpenUrl: (url: string) => void;
  buildLiveEditUrl: (host: string, documentId: string) => string;
  fmtCs: (iso?: string) => string;
};

export default function FiledSummaryCard({
  caseUrl,
  filedCaseName,
  sentPill,
  documentsToShow,
  workspaceHost,
  onOpenUrl,
  buildLiveEditUrl,
    onLockedDocAttempt,

  fmtCs,
}: Props) {
  return (
    <div className="mwFiledSummaryCard">
      <div className="mwFiledSummaryHeader">
        <div className="mwFiledSummaryHeaderLeft">
     <span className="mwFiledSummaryLabel">Case:</span>
          <div className="mwFiledSummaryTitleWrap">
            {caseUrl ? (
              <button
                type="button"
                className="mwFiledCaseLink"
                onClick={() => onOpenUrl(caseUrl)}
                title="Open case"
              >
                {filedCaseName}
              </button>
            ) : (
              <span className="mwFiledCaseName">{filedCaseName}</span>
            )}
          </div>
        </div>
      </div>

      <div className="mwFiledSummaryMeta">
        {sentPill?.atIso ? (
          <span className="mwFiledSummaryMetaItem">{`Filed: ${fmtCs(sentPill.atIso)}`}</span>
        ) : null}

        <span className="mwFiledSummaryMetaItem">{`Documents: ${documentsToShow.length}`}</span>
      </div>

      {documentsToShow.length > 0 ? (
        <div className="mwFiledDocsMini">
          {documentsToShow.map((doc) => (
  <DocumentHoverCard
    key={doc.id}
    doc={doc}
    workspaceHost={workspaceHost}
    onOpenUrl={onOpenUrl}
    buildLiveEditUrl={buildLiveEditUrl}
    onLockedDocAttempt={onLockedDocAttempt}
  />
))}
        </div>
      ) : null}
    </div>
  );
}
