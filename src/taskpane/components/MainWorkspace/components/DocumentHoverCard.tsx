import * as React from "react";
import { Open24Regular, Edit24Regular } from "@fluentui/react-icons";
import { getDocumentMetaRaw, probeDocumentLock, isDocumentLockedError } from "../../../../services/singlecaseDocuments";

export type UploadedItem = {
  id: string;
  name: string;
  url: string;
  kind: "email" | "attachment";
  atIso: string;
  uploadedBy?: string;

  // lock awareness (these already exist in MainWorkspace UploadedItem)
  isLocked?: boolean;
  lockedBy?: string;
  lockedUntilIso?: string;
};

type Props = {
  doc: UploadedItem;
  workspaceHost: string;
  onOpenUrl: (url: string) => void;
  buildLiveEditUrl: (host: string, documentId: string) => string;
  onLockedDocAttempt?: (msg: string) => void;
};

function extractDocMeta(raw: any): any {
  if (!raw) return null;
  if (raw.document) return raw.document;
  if (raw.data) return raw.data;
  if (raw.result) return raw.result;
  return raw;
}

function extractLockInfo(meta: any): { isLocked: boolean; lockedBy?: string; lockedUntilIso?: string } {
  if (!meta) return { isLocked: false };

  // meta.lockInfo is the server's native shape (highest priority)
  const li = meta.lockInfo;

  const locked =
    Boolean(li?.isLocked) ||
    Boolean(meta.isLocked) ||
    Boolean(meta.is_locked) ||
    Boolean(meta.locked) ||
    Boolean(meta.lock?.is_locked) ||
    Boolean(meta.lock?.locked) ||
    Boolean(meta.opened_and_locked) ||
    Boolean(meta.openedAndLocked);

  const lockedBy =
    String(
      li?.lockOwnerName ??
        meta.lockOwnerName ??
        meta.locked_by?.name ??
        meta.locked_by?.email ??
        meta.lockedBy?.name ??
        meta.lockedBy ??
        meta.lock?.locked_by?.name ??
        meta.lock?.locked_by ??
        meta.lock?.user ??
        ""
    ).trim() || undefined;

  const lockedUntilIso =
    String(
      meta.locked_until ??
        meta.lockedUntil ??
        meta.lock?.locked_until ??
        meta.lock?.until ??
        ""
    ).trim() || undefined;

  return { isLocked: locked, lockedBy, lockedUntilIso };
}


function buildLockMessage(lock: { lockedBy?: string; lockedUntilIso?: string }, fallbackName?: string): string {
  const by = String(lock.lockedBy || "").trim();
  const docLabel = fallbackName ? `"${fallbackName}"` : "This document";

  if (by) return `${docLabel} is currently opened by ${by} and cannot be opened right now.`;
  return `${docLabel} is currently opened in SingleCase and cannot be opened right now.`;
}

export default function DocumentHoverCard({
  doc,
  workspaceHost,
  onOpenUrl,
  buildLiveEditUrl,
  onLockedDocAttempt,
}: Props) {
  const [checking, setChecking] = React.useState(false);

  const blockIfLockedFast = React.useCallback((): boolean => {
    if (!doc?.isLocked) return true;

    if (onLockedDocAttempt) {
      onLockedDocAttempt(
        buildLockMessage(
          { lockedBy: doc.lockedBy, lockedUntilIso: doc.lockedUntilIso },
          doc.name
        )
      );
    }
    return false;
  }, [doc?.isLocked, doc?.lockedBy, doc?.lockedUntilIso, doc?.name, onLockedDocAttempt]);

  const preflightLockCheck = React.useCallback(async (): Promise<boolean> => {
    // If no callback, still block fast based on local signal, but do not hard fail.
    // This keeps behaviour safe even if parent forgot to pass handler.
    if (!blockIfLockedFast()) return false;

    // If we cannot check server, allow open (do not block on network issues).
    if (!doc?.id) return true;

    setChecking(true);
    try {
      const metaRaw = await getDocumentMetaRaw(String(doc.id));
      // Log full raw response so we can confirm the server shape in DevTools
      console.log("[preflight] metaRaw", metaRaw);

      // Lock fields (lockInfo, isLocked, lockOwnerName) live at the root of the
      // server response. extractDocMeta() may return an inner envelope object
      // (e.g. raw.document) that does NOT contain those fields.
      // We therefore check lock info at both levels and take whichever signals locked.
      const lockAtRoot = extractLockInfo(metaRaw);
      const lockInEnvelope = extractLockInfo(extractDocMeta(metaRaw));
      const lock = lockAtRoot.isLocked ? lockAtRoot : lockInEnvelope;
      console.log("[preflight] lock", lock);

      if (lock.isLocked) {
        console.log("[preflight] blocked due to lock (meta)", lock);
        if (onLockedDocAttempt) onLockedDocAttempt(buildLockMessage(lock, doc.name));
        return false;
      }

      // Metadata endpoint did not return lock info — probe the version upload
      // endpoint which runs server-side lock validation before body validation.
      console.log("[preflight] meta has no lock info, probing version endpoint");
      const isLocked = await probeDocumentLock(String(doc.id));
      console.log("[preflight] probe result", { isLocked });

      if (isLocked) {
        console.log("[preflight] blocked due to lock (probe 423)");
        if (onLockedDocAttempt) onLockedDocAttempt(buildLockMessage({}, doc.name));
        return false;
      }
      return true;
    } catch (e) {
      if (isDocumentLockedError(e)) {
        if (onLockedDocAttempt) onLockedDocAttempt(buildLockMessage({}, doc.name));
        return false;
      }
      // Network or other error: don't block the user
      return true;
    } finally {
      setChecking(false);
    }
  }, [doc?.id, doc?.name, blockIfLockedFast, onLockedDocAttempt]);

const handleOpen = React.useCallback(async () => {
  console.log("[DocumentHoverCard] handleOpen click", {
    id: doc.id,
    name: doc.name,
    isLocked: doc.isLocked,
    lockedBy: doc.lockedBy,
    lockedUntilIso: doc.lockedUntilIso,
    hasLockedHandler: Boolean(onLockedDocAttempt),
    url: doc.url,
  });

  if (checking) return;
  const ok = await preflightLockCheck();
  if (!ok) return;
  onOpenUrl(doc.url);
}, [checking, preflightLockCheck, onOpenUrl, doc.url, doc.id, doc.name, doc.isLocked, doc.lockedBy, doc.lockedUntilIso, onLockedDocAttempt]);

const handleLiveEdit = React.useCallback(async () => {
  console.log("[DocumentHoverCard] handleLiveEdit click", {
    id: doc.id,
    name: doc.name,
    isLocked: doc.isLocked,
    hasLockedHandler: Boolean(onLockedDocAttempt),
    workspaceHost,
  });

  if (checking) return;
  const ok = await preflightLockCheck();
  if (!ok) return;

  const liveUrl = buildLiveEditUrl(workspaceHost, String(doc.id));
  console.log("[DocumentHoverCard] liveUrl", liveUrl);

  if (liveUrl) onOpenUrl(liveUrl);
}, [checking, preflightLockCheck, buildLiveEditUrl, workspaceHost, doc.id, doc.name, doc.isLocked, onOpenUrl, onLockedDocAttempt]);

  return (
    <div className="mwFiledDocMiniRow">
      <div className="mwFiledDocMiniText" title={doc.name}>
        <span className="mwFiledDocMiniName">{doc.name}</span>

        {doc.uploadedBy ? (
          <span className="mwFiledDocMiniMeta" title={doc.uploadedBy}>
            Uploaded by: {doc.uploadedBy}
          </span>
        ) : null}

        {doc.isLocked ? (
          <span className="mwFiledDocMiniMeta" title="Document is locked">
            Locked
          </span>
        ) : null}
      </div>
      

      <div className="mwFiledDocMiniActions">
        <button
          type="button"
          className="mwMiniIconBtn"
          title={checking ? "Checking..." : "Open"}
          onClick={handleOpen}
          disabled={checking}
        >
          <Open24Regular className="mwIconSvg" />
        </button>

        {workspaceHost && doc.id ? (
          <button
            type="button"
            className="mwMiniIconBtnPrimary"
            title={checking ? "Checking..." : "Edit"}
            onClick={handleLiveEdit}
            disabled={checking}
          >
            
            <Edit24Regular className="mwIconSvgPrimary" />
          </button>
        ) : null}
      </div>
    </div>
  );
}