import { getStored, setStored } from "./storage";

export type SentPillData = {
  sent: boolean;
  atIso?: string;
  caseId?: string;
  singlecaseRecordId?: string;

  documentId?: string;
  revisionNumber?: number;

  filedBy?: string; // NEW
};

const KEY_PREFIX = "sc:sent:";

export function getEmailKey(itemId: string): string {
  return `${KEY_PREFIX}${itemId}`;
}

/**
 * Normalises any previously stored shape
 * and guarantees stable types for the rest of the app.
 */
function normaliseSentPill(raw: unknown): SentPillData | null {
  if (!raw || typeof raw !== "object") return null;
  const v: any = raw;
  const filedBy = typeof v.filedBy === "string" ? v.filedBy : undefined;
  const sent = Boolean(v.sent);

  const atIso = typeof v.atIso === "string" ? v.atIso : undefined;
  const caseId = typeof v.caseId === "string" ? v.caseId : undefined;
  const singlecaseRecordId =
    typeof v.singlecaseRecordId === "string" ? v.singlecaseRecordId : undefined;

  // documentId can be number or string historically
  const documentId =
    typeof v.documentId === "string"
      ? v.documentId
      : typeof v.documentId === "number"
        ? String(v.documentId)
        : undefined;

  // revisionNumber can be string or number historically
  const rev =
    typeof v.revisionNumber === "number"
      ? v.revisionNumber
      : typeof v.revisionNumber === "string"
        ? Number(v.revisionNumber)
        : undefined;

  const revisionNumber = Number.isFinite(rev) ? rev : undefined;

  return {
    sent,
    ...(atIso ? { atIso } : {}),
    ...(caseId ? { caseId } : {}),
    ...(singlecaseRecordId ? { singlecaseRecordId } : {}),
    ...(documentId ? { documentId } : {}),
    ...(revisionNumber !== undefined ? { revisionNumber } : {}),
    ...(filedBy ? { filedBy } : {}), // NEW
  };
}

export async function loadSentPill(itemId: string): Promise<SentPillData | null> {
  if (!itemId) return null;

  const raw = await getStored(getEmailKey(itemId));

  // IMPORTANT: empty string means "cleared"
  if (!raw || raw.trim() === "") return null;

  try {
    const parsed = JSON.parse(raw);
    return normaliseSentPill(parsed);
  } catch {
    return null;
  }
}

export async function saveSentPill(itemId: string, data: SentPillData): Promise<void> {
  if (!itemId) return;

  const normalised = normaliseSentPill(data) ?? { sent: Boolean(data.sent) };
  await setStored(getEmailKey(itemId), JSON.stringify(normalised));
}

/**
 * Clears sent pill state for this email.
 * We intentionally store an empty string because
 * OfficeRuntime.storage has no removeItem.
 */
export async function clearSentPill(itemId: string): Promise<void> {
  if (!itemId) return;
  await setStored(getEmailKey(itemId), "");
}
