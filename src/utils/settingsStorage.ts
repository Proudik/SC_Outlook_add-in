import { AddinSettings } from "../taskpane/components/SettingsModal";

export type DuplicateCache = Record<string, string[]>; // caseId -> itemId list

function getMailboxKeySafe(): string {
  try {
    const p = (Office as any)?.context?.mailbox?.userProfile;
    const email = String(p?.emailAddress || "").toLowerCase().trim();
    return email || "default";
  } catch {
    return "default";
  }
}

function settingsStorageKey(): string {
  return `sc_case_intake_settings:${getMailboxKeySafe()}`;
}

function lastCaseStorageKey(): string {
  return `sc_case_intake_last_case:${getMailboxKeySafe()}`;
}

function duplicateCacheKey(): string {
  return `sc_case_intake_attached_cache:${getMailboxKeySafe()}`;
}

export function loadSettings(defaultSettings: AddinSettings): AddinSettings {
  try {
    const raw = localStorage.getItem(settingsStorageKey());
    if (!raw) return defaultSettings;
    const parsed = JSON.parse(raw);
    const merged = { ...defaultSettings, ...parsed };
    // Migrate old "ask" value → "warn" (renamed in v2)
    if ((merged as any).filingOnSend === "ask") (merged as any).filingOnSend = "warn";
    return merged as AddinSettings;
  } catch {
    return defaultSettings;
  }
}
function everFiledKey(): string {
  return `sc_ever_filed_v1:${getMailboxKeySafe()}`;
}

function loadEverFiled(): Record<string, number> {
  try {
    const raw = localStorage.getItem(everFiledKey());

    const obj = raw ? JSON.parse(raw) : {};
    if (!obj || typeof obj !== "object") return {};
    return obj as Record<string, number>;
  } catch {
    return {};
  }
}

function saveEverFiled(map: Record<string, number>): void {
  try {
    localStorage.setItem(everFiledKey(), JSON.stringify(map));

  } catch {
    // ignore
  }
}

export function hasEverFiled(emailItemId: string): boolean {
  if (!emailItemId) return false;
  const map = loadEverFiled();
  return Boolean(map[emailItemId]);
}

export function markEverFiled(emailItemId: string): void {
  if (!emailItemId) return;
  const map = loadEverFiled();
  map[emailItemId] = Date.now();
  saveEverFiled(map);
}


export function saveSettings(s: AddinSettings) {
  try {
    localStorage.setItem(settingsStorageKey(), JSON.stringify(s));
  } catch {
    // ignore
  }
}

export function loadLastCaseId(): string {
  try {
    return localStorage.getItem(lastCaseStorageKey()) || "";
  } catch {
    return "";
  }
}

export function saveLastCaseId(id: string) {
  try {
    localStorage.setItem(lastCaseStorageKey(), id);
  } catch {
    // ignore
  }
}

export function loadDuplicateCache(): DuplicateCache {
  try {
    const raw = localStorage.getItem(duplicateCacheKey());
    if (!raw) return {};
    return JSON.parse(raw) as DuplicateCache;
  } catch {
    return {};
  }
}

export function saveDuplicateCache(cache: DuplicateCache) {
  try {
    localStorage.setItem(duplicateCacheKey(), JSON.stringify(cache));
  } catch {
    // ignore
  }
}

export function hasAttached(cache: DuplicateCache, caseId: string, itemId: string): boolean {
  const arr = cache[caseId] || [];
  return arr.includes(itemId);
}

export function markAttached(cache: DuplicateCache, caseId: string, itemId: string) {
  const arr = cache[caseId] || [];
  if (!arr.includes(itemId)) {
    cache[caseId] = [...arr, itemId].slice(-200);
  }
}

/**
 * Compute a stable content fingerprint for an email.
 * Two emails with the same subject, sender, and body will produce the same key,
 * regardless of their Office itemId or conversationId.
 * Used as the primary key in DuplicateCache so duplicate detection works across
 * separate copies of "identical" emails (e.g. sent + received).
 */
function djb2Hash(s: string): string {
  let h = 5381;
  for (let i = 0; i < s.length; i++) {
    h = (((h << 5) + h) ^ s.charCodeAt(i)) >>> 0;
  }
  return h.toString(36);
}

export function computeEmailFingerprint(
  subject: string,
  senderEmail: string,
  bodyExcerpt: string
): string {
  const norm = (v: string) =>
    String(v || "").toLowerCase().trim().replace(/\s+/g, " ");
  const key = [
    norm(subject),
    norm(senderEmail),
    norm(bodyExcerpt).substring(0, 200),
  ].join("\x01");
  return `fp:${djb2Hash(key)}`;
}


function discardEmailKey(itemId: string): string {
  return `sc_case_intake_discarded:${getMailboxKeySafe()}:${String(itemId || "").trim()}`;
}

export function isDiscardedEmail(itemId: string): boolean {
  try {
    if (!itemId) return false;
    return localStorage.getItem(discardEmailKey(itemId)) === "true";
  } catch {
    return false;
  }
}

export function setDiscardedEmail(itemId: string, value: boolean) {
  try {
    if (!itemId) return;
    localStorage.setItem(discardEmailKey(itemId), value ? "true" : "false");
  } catch {
    // ignore
  }
}