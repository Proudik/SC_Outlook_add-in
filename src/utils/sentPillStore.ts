// sentPillStore.ts uses localStorage directly (not setStored/getStored) because:
// - Sent pill indicators are display-only UI state for the current device — no cross-device sync needed.
// - setStored falls back to roamingSettings in OWA (OfficeRuntime.storage is Desktop-only),
//   and writing sc:sentPills to roamingSettings contributes to the 32KB overflow.

export type SentPillData = {
  sent: boolean;
  atIso?: string;
  caseId?: string;
  singlecaseRecordId?: string;

  documentId?: string;
  revisionNumber?: number;

  filedBy?: string; // NEW
};

// Blob-based storage: ONE localStorage key holds all sentPills.
// This avoids accumulating individual "sc:sent:${itemId}" keys.
const SENT_PILLS_KEY = "sc:sentPills";
const MAX_SENT_PILLS = 10;

// Legacy key prefix — kept only for backward-compatible reads
const KEY_PREFIX = "sc:sent:";

type SentPillEntry = SentPillData & { _savedAt: number };
type SentPillsBlob = Record<string, SentPillEntry>;

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
    ...(filedBy ? { filedBy } : {}),
  };
}

function loadBlob(): SentPillsBlob {
  try {
    const raw = typeof localStorage !== "undefined" ? localStorage.getItem(SENT_PILLS_KEY) : null;
    if (raw) return JSON.parse(raw) as SentPillsBlob;
  } catch {}
  return {};
}

export async function loadSentPill(itemId: string): Promise<SentPillData | null> {
  if (!itemId) return null;

  // 1. Try blob format
  try {
    const blob = loadBlob();
    const entry = blob[itemId];
    if (entry) {
      const { _savedAt: _unused, ...pillData } = entry;
      return normaliseSentPill(pillData);
    }
  } catch {}

  // 2. Fallback to legacy individual-key format (for emails saved before this version)
  try {
    const legacyRaw = typeof localStorage !== "undefined" ? localStorage.getItem(`${KEY_PREFIX}${itemId}`) : null;
    if (legacyRaw && legacyRaw.trim() !== "") {
      return normaliseSentPill(JSON.parse(legacyRaw));
    }
  } catch {}

  return null;
}

export async function saveSentPill(itemId: string, data: SentPillData): Promise<void> {
  if (!itemId) return;

  const normalised = normaliseSentPill(data) ?? { sent: Boolean(data.sent) };
  let blob = loadBlob();

  blob[itemId] = { ...normalised, _savedAt: Date.now() };

  // Prune to MAX_SENT_PILLS most recent entries
  const entries = Object.entries(blob);
  if (entries.length > MAX_SENT_PILLS) {
    entries.sort((a, b) => (b[1]._savedAt || 0) - (a[1]._savedAt || 0));
    const pruned: SentPillsBlob = {};
    entries.slice(0, MAX_SENT_PILLS).forEach(([k, v]) => { pruned[k] = v; });
    blob = pruned;
  }

  try {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(SENT_PILLS_KEY, JSON.stringify(blob));
    }
  } catch {
    // localStorage full or unavailable — silently ignore, pills are display-only
  }
}

/**
 * Clears sent pill state for this email.
 */
export async function clearSentPill(itemId: string): Promise<void> {
  if (!itemId) return;

  // Remove from blob
  try {
    const blob = loadBlob();
    if (blob[itemId] !== undefined) {
      delete blob[itemId];
      if (typeof localStorage !== "undefined") {
        localStorage.setItem(SENT_PILLS_KEY, JSON.stringify(blob));
      }
    }
  } catch {}

  // Also clear legacy individual key from localStorage
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.removeItem(`${KEY_PREFIX}${itemId}`);
    }
  } catch {}
}
