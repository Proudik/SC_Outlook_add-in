import * as React from "react";
import { Open24Regular, Edit24Regular } from "@fluentui/react-icons";

export type UploadedItem = {
  id: string;
  name: string;
  url: string;
  kind: "email" | "attachment";
  atIso: string;
  uploadedBy?: string;
};

type Props = {
  doc: UploadedItem;
  workspaceHost: string;
  onOpenUrl: (url: string) => void;
  buildLiveEditUrl: (host: string, documentId: string) => string;
};

export default function DocumentHoverCard({
  doc,
  workspaceHost,
  onOpenUrl,
  buildLiveEditUrl,
}: Props) {
  return (
    <div className="mwFiledDocMiniRow">
    <div className="mwFiledDocMiniText" title={doc.name}>
  <span className="mwFiledDocMiniName">{doc.name}</span>

  {doc.uploadedBy ? (
    <span className="mwFiledDocMiniMeta" title={doc.uploadedBy}>
      Nahrál: {doc.uploadedBy}
    </span>
  ) : null}
</div>

      <div className="mwFiledDocMiniActions">
        <button
          type="button"
          className="mwMiniIconBtn"
          title="Open"
          onClick={() => onOpenUrl(doc.url)}
        >
          <Open24Regular className="mwIconSvg" />
        </button>

        {workspaceHost && doc.id ? (
          <button
            type="button"
            className="mwMiniIconBtnPrimary"
            title="Edit"
            onClick={() =>
              onOpenUrl(buildLiveEditUrl(workspaceHost, String(doc.id)))
            }
          >
            <Edit24Regular className="mwIconSvgPrimary" />
          </button>
        ) : null}
      </div>
    </div>
  );
}