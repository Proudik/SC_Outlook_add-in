// recipientHistory uses localStorage directly (not setStored) because:
// - Recipient → case history is per-device suggestion state — no cross-device sync needed.
// - setStored falls back to roamingSettings in OWA and triggers 32KB overflow errors.
import { STORAGE_KEYS } from "./constants";

export type RecipientHistoryEntry = {
  email: string;       // normalised lower case
  caseId: string;
  count: number;
  lastUsedIso: string;
};

function normEmail(v: string): string {
  return String(v || "").trim().toLowerCase();
}

async function loadMap(): Promise<Record<string, RecipientHistoryEntry>> {
  try {
    const raw = typeof localStorage !== "undefined" ? localStorage.getItem(STORAGE_KEYS.recipientHistory) : null;
    if (!raw) return {};
    const obj = JSON.parse(String(raw));
    return obj && typeof obj === "object" ? (obj as any) : {};
  } catch {
    return {};
  }
}

async function saveMap(map: Record<string, RecipientHistoryEntry>) {
  try {
    if (typeof localStorage !== "undefined") localStorage.setItem(STORAGE_KEYS.recipientHistory, JSON.stringify(map));
  } catch { /* ignore */ }
}

export async function recordRecipientsFiledToCase(emails: string[], caseId: string) {
  const cid = String(caseId || "").trim();
  if (!cid) return;

  const nowIso = new Date().toISOString();
  const map = await loadMap();

  for (const e of emails) {
    const email = normEmail(e);
    if (!email) continue;

    const prev = map[email];
    const next: RecipientHistoryEntry = {
      email,
      caseId: cid,
      count: (prev?.caseId === cid ? (prev?.count || 0) : 0) + 1,
      lastUsedIso: nowIso,
    };
    map[email] = next;
  }

  await saveMap(map);
}

export async function findBestCaseForRecipients(
  emails: string[]
): Promise<{ caseId: string; score: number } | null> {
  const map = await loadMap();
  const votes: Record<string, number> = {};

  for (const e of emails) {
    const email = normEmail(e);
    if (!email) continue;

    const hit = map[email];
    if (!hit?.caseId) continue;

    const w = Math.min(10, Math.max(1, Number(hit.count || 1))); // cap weight
    votes[hit.caseId] = (votes[hit.caseId] || 0) + w;
  }

  let bestId = "";
  let bestScore = 0;

  for (const [caseId, score] of Object.entries(votes)) {
    if (score > bestScore) {
      bestScore = score;
      bestId = caseId;
    }
  }

  if (!bestId) return null;
  return { caseId: bestId, score: bestScore };
}