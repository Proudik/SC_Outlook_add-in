// MainWorkspace.tsx (part 1 of 5)

import * as React from "react";
import {
  markEverFiled,
  loadLastCaseId,
  saveLastCaseId,
  loadDuplicateCache,
  saveDuplicateCache,
  hasAttached,
  markAttached,
  computeEmailFingerprint,
} from "../../../utils/settingsStorage";
import { loadSentPill, saveSentPill, SentPillData } from "../../../utils/sentPillStore";
import type { AddinSettings, CaseListScope } from "../SettingsModal";
import CaseSelector from "../CaseSelector";
import {
  listCases,
  listClients,
  submitEmailToCase,
  CaseOption,
} from "../../../services/singlecase";
import { useCaseSuggestions } from "../../../hooks/useCaseSuggestions";
import { suggestCasesByContent } from "../../../utils/caseSuggestionEngine";
import { isInternalEmail } from "../../../utils/internalEmailGuard";
import {
  uploadDocumentToCase,
  uploadDocumentVersion,
  getDocumentMeta,
  checkDuplicateFilename,
} from "../../../services/singlecaseDocuments";
import { recordSuccessfulAttach } from "../../../utils/suggestionStorage";
import { getStored, setStored, removeStored } from "../../../utils/storage";
import { STORAGE_KEYS } from "../../../utils/constants";
import { loadUploadedLinks, saveUploadedLinks } from "../../../utils/uploadedLinksStore";
import FiledSummaryCard from "./components/FiledSummaryCard";
import AttachmentsStep from "./components/AttachmentsStep";
import PromptBubble from "./components/PromptBubble";
import {
  applyFiledCategoryToCurrentEmailOfficeJs,
  applyUnfiledCategoryToCurrentEmailOfficeJs,
  applyUnfiledCategoryToCurrentEmailGraph,
  getCurrentEmailCategoriesGraph,
} from "../../../services/graphMail";

import "./MainWorkspace.css";

// Feature flag: Set to false to silence verbose logging (helps with render loops)
const VERBOSE_LOGGING = false;

declare const Office: any;
declare const OfficeRuntime: any;

type TabId = "cases" | "quick" | "timesheets" | "tasks";

type Props = {
  email: string;
  token: string;
  settings: AddinSettings;
  onChangeSettings: React.Dispatch<React.SetStateAction<AddinSettings>>;
  onSignOut: () => Promise<void> | void;
  onOpenTab: (tab: TabId) => void;
};

type UploadedItem = {
  id: string;
  name: string;
  url: string;
  kind: "email" | "attachment";
  atIso: string;
  uploadedBy?: string;
  caseId?: string;

  // lock awareness
  isLocked?: boolean;
  lockedBy?: string;
  lockedUntilIso?: string;
};

type AttachmentLike = {
  id: string;
  name: string;
  size?: number;
  isInline?: boolean;
};

type PromptState =
  | { itemId: string; kind: "none"; text: string }
  | { itemId: string; kind: "unfiled"; text: string }
  | { itemId: string; kind: "filed"; text: string }
  | { itemId: string; kind: "deleted"; text: string };

type ViewMode = "prompt" | "pickCase" | "sending" | "sent";
type FilingMode = "attachments" | "both";
type PickStep = "case" | "attachments";

type ChatStep =
  | "idle"
  | "compose_wait_recipients"
  | "compose_offer_frequent"
  | "compose_choose_case"
  | "compose_ask_attachments"
  | "compose_ready";

type QuickAction = {
  id: string;
  label: string;
  intent:
    | "accept_frequent_case"
    | "pick_another_case"
    | "toggle_auto_file"
    | "skip_attachments"
    | "select_attachments"
    | "cancel_compose"
    | "view_in_singlecase"
    | "file_manually"
    | "show_suggestions"
    | "refresh_filing_status"
    | "mark_as_unfiled"
    | "file_now_from_weak_signal"
    | "enable_filing_on_send"
    | "confirm_file_now"
    | "skip_pending_filing"
    | "file_now";
};

const FILED_CATEGORY = "SC: Filed";
const UNFILED_CATEGORY = "SC: Unfiled";

/**
 * Recipient history (local) so draft To: can trigger a frequent-case suggestion.
 * Stored under STORAGE_KEYS.recipientHistory if present, otherwise a safe fallback string key.
 */
type RecipientHistoryEntry = {
  email: string;
  caseId: string;
  count: number;
  lastUsedIso: string;
};

const CONV_CASE_KEY_PREFIX = "sc_conv_case:";

const LAST_FILED_CASE_KEY = "sc_last_filed_case";

const LAST_FILED_CTX_KEY = "sc_last_filed_ctx";

const CONV_CTX_KEY_PREFIX = "sc_conv_ctx:";

async function saveConversationFiledCtx(conversationKey: string, ctx: LastFiledCtx) {
  const ck = String(conversationKey || "").trim();
  const caseId = String(ctx.caseId || "").trim();
  const emailDocId = String(ctx.emailDocId || "").trim();
  if (!ck || !caseId || !emailDocId) return;
  await setStored(`${CONV_CTX_KEY_PREFIX}${ck}`, JSON.stringify({ caseId, emailDocId }));
}

type LastFiledCtx = {
  caseId: string;
  emailDocId: string;
};

async function saveLastFiledCtx(ctx: LastFiledCtx) {
  const caseId = String(ctx.caseId || "").trim();
  const emailDocId = String(ctx.emailDocId || "").trim();
  if (!caseId || !emailDocId) return;
  await setStored(LAST_FILED_CTX_KEY, JSON.stringify({ caseId, emailDocId }));
}

async function saveLastFiledCase(caseId: string) {
  const cid = String(caseId || "").trim();
  if (!cid) return;
  await setStored(LAST_FILED_CASE_KEY, cid);
}

async function loadLastFiledCase(): Promise<string> {
  const v = await getStored(LAST_FILED_CASE_KEY);
  return String(v || "").trim();
}

async function saveConversationFiledCase(conversationKey: string, caseId: string) {
  const ck = String(conversationKey || "").trim();
  const cid = String(caseId || "").trim();
  if (!ck || !cid) return;
  await setStored(`${CONV_CASE_KEY_PREFIX}${ck}`, cid);
}

async function loadConversationFiledCase(conversationKey: string): Promise<string> {
  const ck = String(conversationKey || "").trim();
  if (!ck) return "";
  const v = await getStored(`${CONV_CASE_KEY_PREFIX}${ck}`);
  return String(v || "").trim();
}

const RECIPIENT_HISTORY_KEY = (STORAGE_KEYS as any)?.recipientHistory || "recipientHistory";

function normEmail(v: string): string {
  return String(v || "")
    .trim()
    .toLowerCase();
}

async function loadRecipientHistory(): Promise<Record<string, RecipientHistoryEntry>> {
  try {
    const raw = await getStored(RECIPIENT_HISTORY_KEY);
    if (!raw) return {};
    const obj = JSON.parse(String(raw));
    return obj && typeof obj === "object" ? (obj as any) : {};
  } catch {
    return {};
  }
}

async function getOfficeMailCategoriesNorm(): Promise<string[]> {
  try {
    const item = Office?.context?.mailbox?.item as any;
    if (!item?.categories?.getAsync) return [];

    const catsOfficeAny: any[] = await new Promise((resolve) => {
      item.categories.getAsync((res: any) => {
        if (res?.status === Office.AsyncResultStatus.Succeeded) {
          resolve(Array.isArray(res.value) ? res.value : []);
        } else resolve([]);
      });
    });

    const rawNames = catsOfficeAny.map(catToName).filter(Boolean);

    console.log("[officeCats] raw", rawNames);
    console.log("[officeCats] norm", rawNames.map(normaliseCat));

    return rawNames.map(normaliseCat).filter(Boolean);
  } catch {
    return [];
  }
}

async function getOutlookSubjectAsync(): Promise<string> {
  try {
    const item = Office?.context?.mailbox?.item as any;
    if (!item) return "";

    // Read mode often gives a plain string
    if (typeof item.subject === "string") return String(item.subject || "");

    // Compose mode gives a Subject object with getAsync
    if (item?.subject?.getAsync) {
      const v: string = await new Promise((resolve) => {
        item.subject.getAsync((res: any) => {
          if (res?.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || ""));
          else resolve("");
        });
      });
      return v || "";
    }

    // fallback
    return String(item?.subject || "");
  } catch {
    return "";
  }
}

async function saveRecipientHistory(map: Record<string, RecipientHistoryEntry>) {
  try {
    await setStored(RECIPIENT_HISTORY_KEY, JSON.stringify(map));
  } catch {
    // ignore
  }
}

async function recordRecipientsFiledToCase(emails: string[], caseId: string) {
  const cid = String(caseId || "").trim();
  if (!cid) return;

  const nowIso = new Date().toISOString();
  const map = await loadRecipientHistory();

  for (const e of emails) {
    const email = normEmail(e);
    if (!email) continue;

    const prev = map[email];
    const next: RecipientHistoryEntry = {
      email,
      caseId: cid,
      count: (prev?.caseId === cid ? Number(prev?.count || 0) : 0) + 1,
      lastUsedIso: nowIso,
    };
    map[email] = next;
  }

  await saveRecipientHistory(map);
}

// Per-draft autoFileOnSend persistence (localStorage, taskpane-only)
const DRAFT_AUTOFILE_PREFIX = "sc_draft_autofile:";
function getDraftAutoFile(key: string): boolean | null {
  if (!key) return null;
  try {
    const raw = localStorage.getItem(`${DRAFT_AUTOFILE_PREFIX}${key}`);
    if (raw === null) return null;
    return raw === "1";
  } catch { return null; }
}
function setDraftAutoFile(key: string, enabled: boolean): void {
  if (!key) return;
  try { localStorage.setItem(`${DRAFT_AUTOFILE_PREFIX}${key}`, enabled ? "1" : "0"); } catch {}
}

// Draft intent is still useful for future hosts, but on Mac you cannot rely on OnMessageSend.
async function saveComposeIntent(params: {
  itemKey: string;
  caseId: string;
  autoFileOnSend: boolean;
  filingOnSend?: string;
  duplicates?: string;
  baseCaseId?: string;
  baseEmailDocId?: string;
}) {
  try {
    const value = JSON.stringify({
      caseId: params.caseId,
      autoFileOnSend: params.autoFileOnSend,
      filingOnSend: String(params.filingOnSend || ""),
      duplicates: String(params.duplicates || ""),
      baseCaseId: String(params.baseCaseId || "").trim(),
      baseEmailDocId: String(params.baseEmailDocId || "").trim(),
    });

    const key = `sc_intent:${params.itemKey}`;

    console.log("[saveComposeIntent] Saving intent", {
      key,
      caseId: params.caseId,
      autoFileOnSend: params.autoFileOnSend,
      baseCaseId: String(params.baseCaseId || "").trim(),
      baseEmailDocId: String(params.baseEmailDocId || "").trim(),
    });

    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime?.storage) {
      await OfficeRuntime.storage.setItem(key, value);
      console.log("[saveComposeIntent] Saved to OfficeRuntime.storage");
    } else if (Office?.context?.roamingSettings) {
      Office.context.roamingSettings.set(key, value);
      await new Promise<void>((resolve, reject) => {
        Office.context.roamingSettings.saveAsync((result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("[saveComposeIntent] Saved to roamingSettings");
            resolve();
          } else {
            console.error("[saveComposeIntent] roamingSettings.saveAsync failed:", result.error);
            reject(new Error(result.error?.message || "saveAsync failed"));
          }
        });
      });
    } else {
      localStorage.setItem(key, value);
      console.warn("[saveComposeIntent] Saved to localStorage (cross-context won't work)");
    }

    // ALSO save fallback key for new compose emails (no stable itemId before send)
    try {
      const fallbackKey = "sc_intent:last_compose";
      if (typeof OfficeRuntime !== "undefined" && OfficeRuntime?.storage) {
        await OfficeRuntime.storage.setItem(fallbackKey, value);
        console.log("[saveComposeIntent] Saved fallback to OfficeRuntime.storage");
      } else if (Office?.context?.roamingSettings) {
        Office.context.roamingSettings.set(fallbackKey, value);
        await new Promise<void>((resolve) => {
          Office.context.roamingSettings.saveAsync(() => resolve());
        });
        console.log("[saveComposeIntent] Saved fallback to roamingSettings");
      } else {
        localStorage.setItem(fallbackKey, value);
        console.log("[saveComposeIntent] Saved fallback to localStorage");
      }
    } catch (e) {
      console.warn("[saveComposeIntent] Failed to save fallback key:", e);
    }

    // Also save under real itemId if it exists and differs
    try {
      const realId = String(Office?.context?.mailbox?.item?.itemId || "").trim();
      if (realId && realId !== params.itemKey) {
        const altKey = `sc_intent:${realId}`;
        if (typeof OfficeRuntime !== "undefined" && OfficeRuntime?.storage) {
          await OfficeRuntime.storage.setItem(altKey, value);
        } else if (Office?.context?.roamingSettings) {
          Office.context.roamingSettings.set(altKey, value);
          await new Promise<void>((resolve) => {
            Office.context.roamingSettings.saveAsync(() => resolve());
          });
        } else {
          localStorage.setItem(altKey, value);
        }
      }
    } catch {
      // ignore
    }
  } catch (e) {
    console.error("[saveComposeIntent] Failed to save intent:", e);
  }
}

const RECIPIENT_HISTORY_MAX_AGE_MS = 90 * 24 * 60 * 60 * 1000; // 90 days
const RECIPIENT_HISTORY_MIN_COUNT = 2; // or within age window

async function findBestCaseForRecipients(
  emails: string[]
): Promise<{ caseId: string; score: number } | null> {
  const map = await loadRecipientHistory();
  const votes: Record<string, number> = {};
  const now = Date.now();

  for (const e of emails) {
    const email = normEmail(e);
    if (!email) continue;

    const hit = map[email];
    if (!hit?.caseId) continue;

    // Strong-match filter: only use if recent (≤90 days) OR count ≥ 2
    const count = Number(hit.count || 1);
    const lastUsed = hit.lastUsedIso ? new Date(hit.lastUsedIso).getTime() : 0;
    const isRecent = lastUsed > 0 && now - lastUsed <= RECIPIENT_HISTORY_MAX_AGE_MS;
    if (!isRecent && count < RECIPIENT_HISTORY_MIN_COUNT) continue;

    const w = Math.min(10, Math.max(1, count));
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

function getGreetingCz(d: Date): string {
  const h = d.getHours();
  if (h >= 5 && h < 12) return "Good morning";
  if (h >= 12 && h < 18) return "Good afternoon";
  return "Good evening";
}

function openUrl(url: string) {
  if (!url) return;
  try {
    const ui = Office?.context?.ui as any;
    if (ui && typeof ui.openBrowserWindow === "function") {
      ui.openBrowserWindow(url);
      return;
    }
  } catch {
    // ignore
  }
  try {
    window.open(url, "_blank", "noopener,noreferrer");
  } catch {
    // ignore
  }
}

function buildLiveEditUrl(host: string, documentId: string): string {
  const h = (host || "")
    .trim()
    .replace(/^https?:\/\//i, "")
    .split("/")[0];
  if (!h || !documentId) return "";
  return `https://${h}/liveEdit/electronOnlineEditLatestVersion/${encodeURIComponent(documentId)}`;
}

function getOutlookDisplayName(): string {
  try {
    const profile = Office.context.mailbox.userProfile;
    if (profile?.displayName) return profile.displayName;

    const email = profile?.emailAddress;
    if (!email) return "user";

    const local = email.split("@")[0];
    return local
      .replace(/[._-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .split(" ")
      .map((p) => p.charAt(0).toUpperCase() + p.slice(1))
      .join(" ");
  } catch {
    return "user";
  }
}

function getCurrentItemIdSafe(): string {
  try {
    const item = Office?.context?.mailbox?.item;
    return String(item?.itemId || "");
  } catch {
    return "";
  }
}

function isComposeMode(): boolean {
  try {
    const item = Office?.context?.mailbox?.item as any;
    return Boolean(item?.body?.setAsync);
  } catch {
    return false;
  }
}

async function getCurrentItemKey(): Promise<string> {
  try {
    const item = Office?.context?.mailbox?.item as any;
    if (!item) return "";

    const compose = Boolean(item?.body?.setAsync);

    // IMPORTANT: In compose, keep a stable draft key.
   if (compose) {
  const conv = String(item.conversationId || item.conversationKey || "").trim();
  if (conv) {
    const k = `draft:${conv}`;
    if (VERBOSE_LOGGING) console.log("[getCurrentItemKey] compose conv", { conv, k });
    return k;
  }

  const created = String(item.dateTimeCreated || "").trim();
  if (created) {
    const k = `draft:${created}`;
    if (VERBOSE_LOGGING) console.log("[getCurrentItemKey] compose created", { created, k });
    return k;
  }

  if (VERBOSE_LOGGING) console.log("[getCurrentItemKey] compose fallback", { k: "draft:current" });
  return "draft:current";
}

    // READ MODE: now it is safe to use itemId/getItemIdAsync
    const direct = String(item.itemId || "").trim();
    if (direct) return direct;

    if (typeof item.getItemIdAsync === "function") {
      const id: string = await new Promise((resolve) => {
        item.getItemIdAsync((res: any) => {
          if (res?.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || ""));
          else resolve("");
        });
      });
      if (id) return id;
    }

    return "";
  } catch {
    return "";
  }
}

function isMailItem(): boolean {
  try {
    const item = Office?.context?.mailbox?.item;
    return String(item?.itemType || "").toLowerCase() === "message";
  } catch {
    return false;
  }
}

function normaliseCat(s: string): string {
  return String(s || "")
    .normalize("NFKC")
    .replace(/[\u00A0\u2007\u202F]/g, " ") // NBSP and friends
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function catToName(x: any): string {
  if (!x) return "";
  if (typeof x === "string") return x;

  // Office sometimes returns category objects
  const name =
    x.displayName ??
    x.name ??
    x.label ??
    x.value ??
    "";

  return String(name || "");
}

async function getCurrentMailCategoriesNorm(
  
  filedCatNorm: string,
  unfiledCatNorm: string
): Promise<string[]> {
const normaliseList = (arr: any[]) =>
  arr.map((c) => normaliseCat(catToName(c))).filter(Boolean);

  const hasScSignal = (catsNorm: string[]) =>
    catsNorm.includes(filedCatNorm) || catsNorm.includes(unfiledCatNorm);

  let officeNorm: string[] = [];
  let graphNorm: string[] = [];

  // 1) Office read (best effort)
  try {
    const item = Office?.context?.mailbox?.item as any;
    if (item?.categories?.getAsync) {
      const catsOffice: any[] = await new Promise((resolve) => {
  item.categories.getAsync((res: any) => {
    if (res?.status === Office.AsyncResultStatus.Succeeded) {
      resolve(Array.isArray(res.value) ? res.value : []);
    } else resolve([]);
  });
});
officeNorm = normaliseList(catsOffice);
    }
  } catch {
    // ignore
  }

  // 2) Graph read (best effort)
  try {
    const catsGraph = await getCurrentEmailCategoriesGraph();
    graphNorm = normaliseList(Array.isArray(catsGraph) ? catsGraph : []);
  } catch {
    // ignore
  }

  console.log("[cats] sourcePref (after reads)", { officeNorm, graphNorm, hasOfficeSignal: hasScSignal(officeNorm), hasGraphSignal: hasScSignal(graphNorm) });

// Prefer Office if it has our SC labels (it matches what user sees in UI)
if (hasScSignal(officeNorm)) return officeNorm;

// Fallback to Graph (can be delayed or eventually consistent)
if (hasScSignal(graphNorm)) return graphNorm;

  return Array.from(new Set([...(officeNorm || []), ...(graphNorm || [])]));
}

async function syncForceUnfiledFromOutlook(
  filedCatNorm: string,
  unfiledCatNorm: string,
  setForceUnfiledLabel: (v: boolean) => void
) {
  const cats = await getCurrentMailCategoriesNorm(filedCatNorm, unfiledCatNorm);

  if (cats.includes(unfiledCatNorm)) {
    setForceUnfiledLabel(true);
    return;
  }

  if (cats.includes(filedCatNorm)) {
    setForceUnfiledLabel(false);
    return;
  }

  setForceUnfiledLabel(false);
}

function isClosedStatus(status?: string | null): boolean {
  const s = (status || "").toLowerCase();
  if (!s) return false;
  return s.includes("closed") || s.includes("uzav") || s.includes("archiv") || s.includes("done");
}

async function getEmailBodySnippet(maxLen: number): Promise<string> {
  try {
    const item = Office?.context?.mailbox?.item as any;
    if (!item?.body?.getAsync) return "";

    const text: string = await new Promise((resolve) => {
      item.body.getAsync(Office.CoercionType.Text, (res: any) => {
        if (res?.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || ""));
        else resolve("");
      });
    });

    const trimmed = String(text || "").trim();
    if (!trimmed) return "";
    return trimmed.length > maxLen ? trimmed.slice(0, maxLen) : trimmed;
  } catch {
    return "";
  }
}

function getOutlookFromEmail(): string {
  try {
    const item = Office?.context?.mailbox?.item as any;
    const from =
      item?.from?.emailAddress ||
      item?.from?.address ||
      item?.sender?.emailAddress ||
      item?.sender?.address ||
      "";
    return String(from || "");
  } catch {
    return "";
  }
}

function getOutlookFromName(): string {
  try {
    const item = Office?.context?.mailbox?.item as any;
    const name =
      item?.from?.displayName ||
      item?.from?.name ||
      item?.sender?.displayName ||
      item?.sender?.name ||
      "";
    return String(name || "");
  } catch {
    return "";
  }
}


async function getDraftRecipientsEmailsAsync(): Promise<string[]> {
  try {
    const item = Office?.context?.mailbox?.item as any;
    if (!item) return [];

    const readField = (field: any) =>
      new Promise<any[]>((resolve) => {
        try {
          if (!field?.getAsync) return resolve([]);
          field.getAsync((res: any) => {
            if (res?.status === Office.AsyncResultStatus.Succeeded) {
              resolve(Array.isArray(res.value) ? res.value : []);
            } else {
              resolve([]);
            }
          });
        } catch {
          resolve([]);
        }
      });

    const to = await readField(item.to);
    const cc = await readField(item.cc);
    const bcc = await readField(item.bcc);

    const all = [...to, ...cc, ...bcc]
      .map((r: any) => normEmail(r?.emailAddress || r?.address || ""))
      .filter(Boolean);

    return Array.from(new Set(all));
  } catch {
    return [];
  }
}

/**
 * Read recipients (To + Cc) from an opened email in read mode.
 * In read mode these are plain arrays, not async getters like compose mode.
 * BCC is not available in read mode.
 */
function getReadModeRecipientEmails(): string[] {
  try {
    const item = Office?.context?.mailbox?.item as any;
    if (!item) return [];
    const toArr = Array.isArray(item.to) ? item.to : [];
    const ccArr = Array.isArray(item.cc) ? item.cc : [];
    const all = [...toArr, ...ccArr]
      .map((r: any) => normEmail(r?.emailAddress || r?.address || ""))
      .filter(Boolean);
    return Array.from(new Set(all));
  } catch {
    return [];
  }
}

/**
 * Extract case keys from SingleCase submail addresses in recipients
 * Example: "2023-0006@valfor-demo.singlecase.ch" -> ["2023-0006"]
 */
function extractSubmailCaseKeys(recipients: string[], workspaceHost: string): string[] {
  if (!workspaceHost || recipients.length === 0) return [];

  // Normalize workspace host (remove protocol, trailing slash)
  const normalizedHost = workspaceHost.toLowerCase().replace(/^https?:\/\//, "").replace(/\/$/, "");

  const caseKeys: string[] = [];

  for (const email of recipients) {
    const emailLower = email.toLowerCase().trim();
    if (!emailLower || !emailLower.includes("@")) continue;

    const [localPart, domain] = emailLower.split("@");
    if (!domain || !localPart) continue;

    // Check if domain matches workspace host
    if (domain === normalizedHost) {
      // Extract case key from local part (e.g., "2023-0006")
      const caseKey = localPart.trim();
      if (caseKey) {
        caseKeys.push(caseKey);
        console.log("[extractSubmailCaseKeys] Found submail case key:", caseKey, "from:", email);
      }
    }
  }

  return caseKeys;
}

/**
 * Normalize case key by replacing dashes with periods for comparison
 * Example: "2023-0005-001" -> "2023.0005.001"
 */
function normalizeCaseKey(key: string): string {
  return key.toLowerCase().trim().replace(/-/g, ".");
}

/**
 * Resolve case key to case ID from cases list
 * Matches case key (e.g., "2023-0006" or "2023-0005-001") against case names/titles
 * Normalizes dashes/periods for comparison (e.g., "2023-0005-001" matches "2023-0005.001")
 * Returns { caseId, caseName, caseKey } if found, null otherwise
 */
function resolveSubmailCaseKey(
  caseKey: string,
  cases: CaseOption[]
): { caseId: string; caseName: string; caseKey: string } | null {
  if (!caseKey || !cases || cases.length === 0) return null;

  const keyNorm = normalizeCaseKey(caseKey);

  // Try to find case where name/title contains the case key in parentheses
  // Example: "Internal Know How (2023-0006)" matches key "2023-0006"
  // Example: "Airbus aircraft lease contract (2023-0005.001)" matches key "2023-0005-001"
  const matches = cases.filter((c) => {
    const name = String((c as any)?.name || (c as any)?.title || "").toLowerCase();

    // Look for match in parentheses at end: "(2023-0006)" or "(2023-0005.001)"
    const parenMatch = name.match(/\(([^)]+)\)\s*$/);
    if (parenMatch) {
      const caseKeyInParen = normalizeCaseKey(parenMatch[1]);
      if (caseKeyInParen === keyNorm) {
        return true;
      }
    }

    // Also check if normalized case key appears anywhere in the name
    const nameNorm = name.replace(/-/g, ".");
    return nameNorm.includes(keyNorm);
  });

  if (matches.length === 0) {
    console.log("[resolveSubmailCaseKey] No case found for key:", caseKey);
    return null;
  }

  if (matches.length > 1) {
    console.warn("[resolveSubmailCaseKey] Multiple cases match key:", caseKey, matches.length);
    return null; // Don't auto-select if ambiguous
  }

  const matchedCase = matches[0];
  const caseId = String((matchedCase as any)?.id || "");
  const caseName = String((matchedCase as any)?.name || (matchedCase as any)?.title || "");

  console.log("[resolveSubmailCaseKey] Resolved case:", { caseKey, caseId, caseName });

  return { caseId, caseName, caseKey };
}

function getConversationKey(): string {
  try {
    const item = Office?.context?.mailbox?.item as any;
    return String(item?.conversationId || item?.conversationKey || "");
  } catch {
    return "";
  }
}

function getOutlookAttachmentsLite(): AttachmentLike[] {
  try {
    const item = Office?.context?.mailbox?.item as any;
    const atts = Array.isArray(item?.attachments) ? item.attachments : [];

    return atts
      .filter((a: any) => !a?.isInline)
      .map((a: any) => ({
        id: String(a.id),
        name: String(a.name || ""),
        size: Number(a.size || 0),
        isInline: Boolean(a.isInline),
      }));
  } catch {
    return [];
  }
}

async function getAttachmentBase64(
  attachmentId: string,
  fallbackName?: string
): Promise<{ base64: string; name: string; mime: string }> {
  const item = Office?.context?.mailbox?.item as any;

  return new Promise((resolve, reject) => {
    try {
      if (!item?.getAttachmentContentAsync)
        return reject(new Error("Attachment API not available"));

      item.getAttachmentContentAsync(attachmentId, (res: any) => {
        if (!res || res.status !== Office.AsyncResultStatus.Succeeded) return reject(res?.error);

        const v = res.value || {};

        // Prefer Office response name, then passed fallback, then try item.attachments, then last resort
        const nameFromItem = Array.isArray(item?.attachments)
          ? String(
              item.attachments.find((a: any) => String(a?.id) === String(attachmentId))?.name || ""
            )
          : "";

        const finalName = String(v.name || fallbackName || nameFromItem || "attachment");

        resolve({
          base64: String(v.content || ""),
          name: finalName,
          mime: String(v.contentType || "application/octet-stream"),
        });
      });
    } catch (e) {
      reject(e);
    }
  });
}

async function clearLocalFiling(itemKey: string) {
  try {
    await saveSentPill(itemKey, null as any);
  } catch {
    // ignore
  }
  try {
    await saveUploadedLinks(itemKey, [] as any);
  } catch {
    // ignore
  }
}

/**
 * Checks whether the documents referenced by this pill/links still exist in SingleCase.
 * Returns:
 *   true  — at least one document confirmed present (API returned 200)
 *   false — document confirmed absent (API returned 404, document genuinely deleted)
 *   null  — cannot determine; API threw (auth/network error, temporary outage)
 *
 * Callers must treat null as "uncertain" and avoid showing a "deleted" message.
 */
async function hasAnyRealDocuments(pill: SentPillData | null, itemKey: string): Promise<boolean | null> {
  // Check pill's document ID first
  const docId = String((pill as any)?.documentId || "").trim();
  if (docId) {
    try {
      const meta = await getDocumentMeta(docId);
      // meta === null means 404 (document confirmed gone)
      // meta with .id means document exists
      return meta !== null;
    } catch {
      // API threw — auth expired, network error, server error.
      // We cannot confirm whether the document exists.
      return null;
    }
  }

  // Check uploaded links
  const raw = (await loadUploadedLinks(itemKey).catch(() => [])) as any[];
  const links = Array.isArray(raw) ? raw : [];

  // No local record of any document ID — we can't confirm deletion.
  // This happens when email was filed on another device/session.
  // Return null (uncertain) to avoid false "deleted" message.
  if (links.length === 0) return null;

  for (const it of links.slice(0, 5)) {
    const id = String(it?.id || "").trim();
    if (!id) continue;
    try {
      const meta = await getDocumentMeta(id);
      if (meta?.id) return true;
      // meta === null → this link is 404, continue checking others
    } catch {
      // API error → can't determine
      return null;
    }
  }

  // All links checked, all returned 404
  return false;
}

function toBase64Utf8(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

function safeFileName(value: string): string {
  const v = (value || "").trim();
  const cleaned = v
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned.slice(0, 80) || "email";
}

function buildDocumentUrl(host: string, documentId: string): string {
  const h = (host || "")
    .trim()
    .replace(/^https?:\/\//i, "")
    .split("/")[0];
  if (!h || !documentId) return "";
  return `https://${h}/?/documents/view/${encodeURIComponent(documentId)}`;
}

const CASE_URL_PATH_PREFIX = "/?/cases/view/";

function buildCaseUrl(host: string, caseId: string): string {
  const h = (host || "")
    .trim()
    .replace(/^https?:\/\//i, "")
    .split("/")[0];
  if (!h || !caseId) return "";
  return `https://${h}${CASE_URL_PATH_PREFIX}${encodeURIComponent(caseId)}`;
}

function tryGetCaseUrlFromCaseOption(host: string, c: any): string {
  if (!c) return "";

  const direct = c.url || c.webUrl || c.web_url || c.href || c.link;

  if (direct && typeof direct === "string") {
    if (/^https?:\/\//i.test(direct)) return direct;

    const h = (host || "")
      .trim()
      .replace(/^https?:\/\//i, "")
      .split("/")[0];
    if (!h) return "";
    return `https://${h}${direct.startsWith("/") ? direct : `/${direct}`}`;
  }

  return "";
}

function fmtCs(iso?: string): string {
  if (!iso) return "";
  try {
    return new Date(iso).toLocaleString("cs-CZ");
  } catch {
    return "";
  }
}

function allSettled<T>(promises: Array<Promise<T>>) {
  return Promise.all(
    promises.map((p) =>
      p.then(
        (value) => ({ status: "fulfilled" as const, value }),
        (reason) => ({ status: "rejected" as const, reason })
      )
    )
  );
}

function extractDocMeta(raw: any): any {
  if (!raw) return null;
  if (raw.document) return raw.document;
  if (raw.data) return raw.data;
  if (raw.result) return raw.result;
  return raw;
}

function extractLatestRevisionNumber(doc: any): number | null {
  const n = doc?.latest_version?.revision_number;
  const num = Number(n);
  return Number.isFinite(num) && num > 0 ? num : null;
}

function extractLockInfo(meta: any): { isLocked: boolean; lockedBy?: string; lockedUntilIso?: string } {
  if (!meta) return { isLocked: false };

  // Accept several possible shapes without guessing too hard.
  // You can refine once you see real meta payload fields.
  const locked =
    Boolean(meta.locked) ||
    Boolean(meta.is_locked) ||
    Boolean(meta.isLocked) ||
    Boolean(meta.lock?.is_locked) ||
    Boolean(meta.lock?.locked) ||
    Boolean(meta.opened_and_locked) ||
    Boolean(meta.openedAndLocked);

  const lockedBy =
    String(
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

function extractRevisionFromVersionUploadResponse(raw: any): number | null {
  if (!raw) return null;
  const v = raw.data || raw.result || raw.document || raw.version || raw;

  const candidate =
    v?.latest_version?.revision_number ??
    v?.latestVersion?.revisionNumber ??
    v?.latestVersion?.revision_number ??
    v?.revision_number ??
    v?.revisionNumber ??
    v?.revision;

  const n = Number(candidate);
  return Number.isFinite(n) && n > 0 ? n : null;
}

function getStoreKey(activeItemKey: string, activeItemId: string, composeMode: boolean): string {
  if (composeMode) return String(activeItemKey || activeItemId || "").trim();
  return String(activeItemId || "").trim();
}


// ── Internal email handling decision ─────────────────────────────────────────
// Single source of truth: determines whether internal-email suppression is
// active for the current email. All logic branches read from this function.
function getInternalHandlingDecision({
  internalHandlingEnabled,
  isInternal,
}: {
  internalHandlingEnabled: boolean;
  isInternal: boolean;
}): { suppressSuggestions: boolean; showInternalBanner: boolean } {
  const active = internalHandlingEnabled && isInternal;
  console.log(
    "[InternalEmail] enabled:", internalHandlingEnabled,
    "detected:", isInternal,
    "suppressSuggestions:", active
  );
  return { suppressSuggestions: active, showInternalBanner: active };
}

// MainWorkspace.tsx (part 3 of 5)

export default function MainWorkspace({ email, token, settings, onChangeSettings }: Props) {
  void email;

  const greeting = React.useMemo(() => getGreetingCz(new Date()), []);
  const userLabel = React.useMemo(() => getOutlookDisplayName(), []);
  const filedCatNorm = React.useMemo(() => normaliseCat(FILED_CATEGORY), []);
  const unfiledCatNorm = React.useMemo(() => normaliseCat(UNFILED_CATEGORY), []);
  const [activeItemId, setActiveItemId] = React.useState<string>(
    () => getCurrentItemIdSafe() || ""
  );
  const composeMode = React.useMemo(() => isComposeMode(), [activeItemId]);

  const [activeItemKey, setActiveItemKey] = React.useState<string>("");
  const [viewMode, setViewMode] = React.useState<ViewMode>("prompt");
  // Ref kept in sync every render — lets effects read viewMode without adding it to deps
  const viewModeRef = React.useRef<ViewMode>(viewMode);
  viewModeRef.current = viewMode;
  const [pickStep, setPickStep] = React.useState<PickStep>("case");



  const [prompt, setPrompt] = React.useState<PromptState>({
    itemId: "",
    kind: "none",
    text: "Select an email and I’ll show you relevant suggestions.",
  });

  const dismissedRef = React.useRef<Set<string>>(new Set());

  const chatBodyRef = React.useRef<HTMLDivElement | null>(null);
  const attachmentsRef = React.useRef<HTMLDivElement | null>(null);
  const chatEndRef = React.useRef<HTMLDivElement | null>(null);

  const [clientNamesById, setClientNamesById] = React.useState<Record<string, string>>({});
  const [cases, setCases] = React.useState<CaseOption[]>([]);
  const [isLoadingCases, setIsLoadingCases] = React.useState(false);

  const [selectedCaseId, setSelectedCaseId] = React.useState("");
  const [selectedSource, setSelectedSource] = React.useState<
    "" | "remembered" | "last_case" | "suggested" | "manual"
  >("");

  // Suggestion state: tracks system-recommended case without polluting selectedCaseId.
  // suggestedCaseId is written ONLY by suggestion/detection logic (never by user clicks).
  // selectedCaseId is written ONLY by explicit user picks or rememberLastCase restores.
  // filingTargetCaseId is set when user clicks Continue (from either source).
  const [suggestedCaseId, setSuggestedCaseId] = React.useState<string>("");
  const [suggestedConfidencePct, setSuggestedConfidencePct] = React.useState<number>(0);
  const [filingTargetCaseId, setFilingTargetCaseId] = React.useState<string>("");
  const [forceUnfiledLabel, setForceUnfiledLabel] = React.useState(false);

  // Content-based suggestions (triggered when user clicks "Vybrat jiný spis")
  const [contentBasedSuggestions, setContentBasedSuggestions] = React.useState<any[]>([]);
  const [isLoadingContentSuggestions] = React.useState(false);

  // Internal email detection state
  const [isInternalEmailDetected, setIsInternalEmailDetected] = React.useState(false);

  // Derived: single source of truth for internal-email suppression
  const { suppressSuggestions: suppressInternalSuggestions } = React.useMemo(
    () =>
      getInternalHandlingDecision({
        internalHandlingEnabled: settings.internalEmailHandling === "doNotSuggest",
        isInternal: isInternalEmailDetected,
      }),
    [settings.internalEmailHandling, isInternalEmailDetected]
  );

  const [sentPill, setSentPill] = React.useState<SentPillData | null>(null);
  const [workspaceHost, setWorkspaceHost] = React.useState<string>("");
  const [submitError, setSubmitError] = React.useState<string>("");
  const [uploadedLinksRaw, setUploadedLinksRaw] = React.useState<UploadedItem[]>([]);
  const [uploadedLinksValidated, setUploadedLinksValidated] = React.useState<UploadedItem[]>([]);

const [lockedDocAlert, setLockedDocAlert] = React.useState<string>("");

  const [isTogglingCategory, setIsTogglingCategory] = React.useState(false);
  const [filingMode, setFilingMode] = React.useState<FilingMode>("both");
  const [selectedAttachments, setSelectedAttachments] = React.useState<string[]>([]);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [isItemLoading, setIsItemLoading] = React.useState(false);

  const [replyBaseCaseId, setReplyBaseCaseId] = React.useState("");
  const [replyBaseEmailDocId, setReplyBaseEmailDocId] = React.useState("");

  const [renameOpen, setRenameOpen] = React.useState(false);
  const [renameDoc, setRenameDoc] = React.useState<UploadedItem | null>(null);
  const [renameValue, setRenameValue] = React.useState("");
  const [renameSaving, setRenameSaving] = React.useState(false);

  const [subjectText, setSubjectText] = React.useState<string>("");

  // Ephemeral tab state — initialised synchronously from the stored default so there is no
  // flash on startup. Tab clicks update this only; the stored setting (settings.caseListScope)
  // only changes when the user saves from the Settings modal.
  const [caseGroupTab, setCaseGroupTab] = React.useState<"favourites" | "all">(() =>
    settings.caseListScope === "favourites" ? "favourites" : "all"
  );

  // When the stored default changes (via Settings modal), reset the active tab to match.
  React.useEffect(() => {
    setCaseGroupTab(settings.caseListScope === "favourites" ? "favourites" : "all");
  }, [settings.caseListScope]);

  React.useEffect(() => {
    let alive = true;
    void (async () => {
      const s = await getOutlookSubjectAsync();
      if (alive) setSubjectText(s || "");
    })();
    return () => {
      alive = false;
    };
  }, [activeItemId]);

  React.useEffect(() => {
  let mounted = true;

  void (async () => {
    if (isComposeMode()) return;
    if (isTogglingCategory) return;

    // DESKTOP OUTLOOK FIX: Add delay to allow category changes to propagate
    // Desktop Outlook has a delay before category changes are reflected in the API
    await new Promise(resolve => setTimeout(resolve, 500));
    if (!mounted) return;

    const cats = await getCurrentMailCategoriesNorm(filedCatNorm, unfiledCatNorm);
    if (!mounted) return;

    if (cats.includes(unfiledCatNorm)) {
      setForceUnfiledLabel(true);
      return;
    }

    if (cats.includes(filedCatNorm)) {
      setForceUnfiledLabel(false);
      return;
    }

    setForceUnfiledLabel(false);
  })();

  return () => {
    mounted = false;
  };
}, [activeItemId, filedCatNorm, unfiledCatNorm, isTogglingCategory]);

React.useEffect(() => {
  if (composeMode) return;
  if (viewMode !== "sent") return;
  if (isTogglingCategory) return;

  void syncForceUnfiledFromOutlook(filedCatNorm, unfiledCatNorm, setForceUnfiledLabel);
}, [viewMode, composeMode, activeItemId, filedCatNorm, unfiledCatNorm, isTogglingCategory]);

const fromEmail = React.useMemo(
  () => getOutlookFromEmail(),
  [activeItemId, activeItemKey]
);

const fromName = React.useMemo(
  () => getOutlookFromName(),
  [activeItemId, activeItemKey]
);

const conversationKey = React.useMemo(
  () => getConversationKey(),
  [activeItemId, activeItemKey]
);

const attachmentsLite = React.useMemo(
  () => getOutlookAttachmentsLite(),
  [activeItemId, activeItemKey]
);

const attachmentIds = React.useMemo(
  () => (attachmentsLite || []).map((a) => String(a.id)).filter(Boolean),
  [attachmentsLite]
);

  React.useEffect(() => {
    if (!composeMode) return;
    setFilingMode("both");
  }, [composeMode, activeItemId]);

  const storeKey = React.useMemo(
    () => getStoreKey(activeItemKey, activeItemId, composeMode),
    [activeItemKey, activeItemId, composeMode]
  );

  const [suggestBodySnippet, setSuggestBodySnippet] = React.useState("");
  const [isUploadingNewVersion, setIsUploadingNewVersion] = React.useState(false);

  const filedCaseId = React.useMemo(() => {
    const id = sentPill?.caseId ? String(sentPill.caseId) : "";
    return id;
  }, [sentPill?.caseId]);


  const [autoFileUserSet, setAutoFileUserSet] = React.useState(false);

  const showFiledSummary = React.useMemo(() => {
    // Only show FiledSummary if we have actual data (sentPill or uploaded documents)
    const hasData = sentPill?.caseId || uploadedLinksValidated.length > 0;

    if (viewMode === "sent" && hasData) return true;
    if (viewMode === "prompt" && prompt.kind === "filed" && sentPill?.caseId) return true;
    return false;
  }, [viewMode, prompt.kind, sentPill?.caseId, uploadedLinksValidated.length]);

  const visibleCases = React.useMemo(() => {
    if (caseGroupTab === "all") return cases;
    return cases.filter((c) => !isClosedStatus((c as any)?.status));
  }, [cases, caseGroupTab]);

  const [composeRecipientsLive, setComposeRecipientsLive] = React.useState<string[]>([]);
  const [chatStep, setChatStep] = React.useState<ChatStep>("idle");
  const [quickActions, setQuickActions] = React.useState<QuickAction[]>([]);
  const [detectedFrequentCaseId, setDetectedFrequentCaseId] = React.useState<string>("");
  const [autoFileOnSend, setAutoFileOnSend] = React.useState<boolean>(
    () => settings.filingOnSend !== "off"
  );
  const [dismissedFrequentKey, setDismissedFrequentKey] = React.useState<string>("");

  React.useEffect(() => {
    if (!autoFileUserSet) setAutoFileOnSend(settings.filingOnSend !== "off");
  }, [settings.filingOnSend, autoFileUserSet]);

  // Restore per-draft autoFileOnSend when switching between compose drafts.
  // "off" → always false; "warn"/"always" → load saved preference or default to true.
  React.useEffect(() => {
    if (!composeMode || !storeKey) return;
    if (settings.filingOnSend === "off") {
      setAutoFileOnSend(false);
      return;
    }
    const stored = getDraftAutoFile(storeKey);
    if (stored !== null) {
      setAutoFileOnSend(stored);
      setAutoFileUserSet(true);
    } else {
      // No explicit preference for this draft → reset to default-on
      setAutoFileUserSet(false);
      setAutoFileOnSend(true);
    }
  }, [composeMode, storeKey, settings.filingOnSend]);

  // Submail detection state
  const [submailDetectedCaseId, setSubmailDetectedCaseId] = React.useState<string>("");
  const [submailDetectedCaseKey, setSubmailDetectedCaseKey] = React.useState<string>("");
  const [submailDetectedCaseName, setSubmailDetectedCaseName] = React.useState<string>("");

  // Already filed detection state (read mode, server-authoritative)
  const [filedStatusChecked, setFiledStatusChecked] = React.useState(false);
  const [alreadyFiled, setAlreadyFiled] = React.useState(false);
  const [alreadyFiledCaseId, setAlreadyFiledCaseId] = React.useState("");
  const [alreadyFiledCaseLabel, setAlreadyFiledCaseLabel] = React.useState("");
  const [alreadyFiledDocumentId, setAlreadyFiledDocumentId] = React.useState("");
  const [allowRefilingOverride, setAllowRefilingOverride] = React.useState(false);
  const [duplicateFilingWarning, setDuplicateFilingWarning] = React.useState(false);
  const [pendingFiling, setPendingFiling] = React.useState<{
    caseId: string;
    subject: string;
    conversationId: string;
    sentAt: string;
  } | null>(null);

  // Divergence: Outlook category says "SC: Filed" but server says not filed
  const [filingDivergenceDetected, setFilingDivergenceDetected] = React.useState(false);

  const suggestionEmail = React.useMemo(() => {
    if (!composeMode) return fromEmail;
    return composeRecipientsLive?.[0] || "";
  }, [composeMode, fromEmail, composeRecipientsLive]);

  async function clearComposeIntent(itemKey: string) {
    try {
      await removeStored(`sc_intent:${itemKey}`);
    } catch {
      // ignore
    }

    try {
      const realId = String(Office?.context?.mailbox?.item?.itemId || "").trim();
      if (realId && realId !== itemKey) {
        await removeStored(`sc_intent:${realId}`);
      }
    } catch {
      // ignore
    }
  }

  React.useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const hostRaw = (await getStored(STORAGE_KEYS.workspaceHost)) || "";
        const host = hostRaw
          .replace(/^https?:\/\//i, "")
          .split("/")[0]
          .trim();
        if (!cancelled) setWorkspaceHost(host);
      } catch {
        // ignore
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  React.useEffect(() => {
    // Compose mode is handled by the polling effect below; skip here to avoid a one-shot
    // fetch that would immediately go stale when the user keeps typing.
    if (composeMode) return () => {};
    let mounted = true;
    void (async () => {
      if (!activeItemId) {
        if (mounted) setSuggestBodySnippet("");
        return;
      }
      const snip = await getEmailBodySnippet(600);
      if (mounted) setSuggestBodySnippet(snip || "");
    })();
    return () => {
      mounted = false;
    };
  }, [activeItemId, composeMode]);

  // Compose mode only: poll subject + body every 400 ms so case suggestions stay
  // live while the user is typing (activeItemId does not change in compose mode).
  React.useEffect(() => {
    if (!composeMode) return () => {};

    let mounted = true;
    let lastSubj = "";
    let lastBody = "";
    let inFlight = false;

    const tick = async () => {
      if (!mounted || inFlight) return;
      inFlight = true;
      try {
        const s = String(await getOutlookSubjectAsync()).trim();
        const b = String(await getEmailBodySnippet(600)).trim();
        if (!mounted) return;
        if (s !== lastSubj) { lastSubj = s; setSubjectText(s); }
        if (b !== lastBody) { lastBody = b; setSuggestBodySnippet(b); }
      } finally {
        inFlight = false;
      }
    };

    const id = window.setInterval(() => void tick(), 400);
    void tick(); // run immediately so the first render already has values

    return () => {
      mounted = false;
      window.clearInterval(id);
    };
  }, [composeMode, activeItemKey]);

  React.useEffect(() => {
    let cancelled = false;

    void (async () => {
      if (!storeKey) {
        if (!cancelled) {
          setUploadedLinksRaw([]);
          setUploadedLinksValidated([]);
        }
        return;
      }

      try {
        const links = (await loadUploadedLinks(storeKey)) as any[];
        if (!cancelled) setUploadedLinksRaw(Array.isArray(links) ? (links as UploadedItem[]) : []);
      } catch {
        if (!cancelled) setUploadedLinksRaw([]);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [storeKey, viewMode]);

  // MainWorkspace.tsx (part 4 of 5)

  React.useEffect(() => {
    let cancelled = false;

    void (async () => {
      if (!storeKey) return;

      const base = Array.isArray(uploadedLinksRaw) ? uploadedLinksRaw : [];
      if (base.length === 0) {
        if (!cancelled) setUploadedLinksValidated([]);
        return;
      }

      const expectedCaseId = String(sentPill?.caseId || "").trim();
      const slice = base.slice(0, 10);

      const results = await allSettled(
        slice.map(async (it) => {
          const id = String(it?.id || "").trim();
          if (!id) return null;

          let meta: any = null;
          try {
            const metaRaw = await getDocumentMeta(id);
            meta = extractDocMeta(metaRaw);
          } catch {
            meta = null;
          }

          // If meta is missing, keep the item as-is (do NOT drop)
          if (!meta) return it as UploadedItem;

               const metaCaseId = String(meta.case_id || meta.caseId || "").trim();
        if (expectedCaseId && metaCaseId && metaCaseId !== expectedCaseId) return null;

        const name = String(it.name || meta.name || "").trim();

        const lock = extractLockInfo(meta);

        return {
          ...it,
          name: name || it.name,
          caseId: metaCaseId || it.caseId,
          isLocked: lock.isLocked,
          lockedBy: lock.lockedBy,
          lockedUntilIso: lock.lockedUntilIso,
        } as UploadedItem;
        })
      );

      const existing = results
        .filter((r) => r.status === "fulfilled")
        .map((r) => (r as any).value)
        .filter(Boolean) as UploadedItem[];

      const seen = new Set<string>();
      const deduped: UploadedItem[] = [];
      for (const it of existing) {
        const id = String(it.id || "");
        if (!id) continue;
        if (seen.has(id)) continue;
        seen.add(id);
        deduped.push(it);
      }

      if (cancelled) return;

      setUploadedLinksValidated(deduped);
    })();

    return () => {
      cancelled = true;
    };
  }, [storeKey, uploadedLinksRaw, sentPill?.caseId]);

  React.useEffect(() => {
    if (viewMode !== "pickCase") return;
    if (pickStep !== "attachments") return;

    setSelectedAttachments(attachmentIds);

    requestAnimationFrame(() => {
      attachmentsRef.current?.scrollIntoView({ block: "start", behavior: "smooth" });
    });
  }, [viewMode, pickStep, attachmentIds]);

  React.useEffect(() => {
    chatEndRef.current?.scrollIntoView({ block: "end", behavior: "auto" });
  }, [
    viewMode,
    pickStep,
    prompt.itemId,
    prompt.kind,
    prompt.text,
    selectedCaseId,
    isSubmitting,
    uploadedLinksValidated.length,
    sentPill?.sent,
  ]);

  React.useEffect(() => {
    console.log("[useEffect:saveIntent] Triggered", {
      composeMode,
      storeKey,
      selectedCaseId,
      autoFileOnSend,
    });

    if (!composeMode) {
      console.log("[useEffect:saveIntent] Skipping - not compose mode");
      return;
    }
    if (!storeKey) {
      console.log("[useEffect:saveIntent] Skipping - no storeKey");
      return;
    }

    const wantToFile =
      settings.filingOnSend === "always"
        ? Boolean(selectedCaseId)
        : selectedCaseId && autoFileOnSend;

    if (wantToFile) {
      const shouldVersion =
        isUploadingNewVersion &&
        Boolean(replyBaseEmailDocId) &&
        Boolean(replyBaseCaseId) &&
        String(selectedCaseId) === String(replyBaseCaseId);

      void saveComposeIntent({
        itemKey: storeKey,
        caseId: selectedCaseId,
        autoFileOnSend,
        filingOnSend: settings.filingOnSend,
        duplicates: settings.duplicates,
        baseCaseId: shouldVersion ? replyBaseCaseId : "",
        baseEmailDocId: shouldVersion ? replyBaseEmailDocId : "",
      });
    } else {
      void clearComposeIntent(storeKey);
    }
  }, [
    composeMode,
    storeKey,
    selectedCaseId,
    autoFileOnSend,
    settings.filingOnSend,
    isUploadingNewVersion,
    replyBaseCaseId,
    replyBaseEmailDocId,
  ]);

  React.useEffect(() => {
    const needCases = viewMode === "pickCase" || showFiledSummary || composeMode;
    if (!needCases) return undefined;

    let mounted = true;

    void (async () => {
      setIsLoadingCases(true);
      try {
        const [casesRes, clientsRes] = await Promise.all([
          listCases(token, settings.caseListScope),
          listClients(token),
        ]);

        if (!mounted) return;

        setCases(casesRes);

        const map: Record<string, string> = {};
        for (const c of clientsRes) map[String(c.id)] = c.name;
        setClientNamesById(map);
      } catch {
        if (!mounted) return;
        setCases([]);
        setClientNamesById({});
      } finally {
        if (mounted) setIsLoadingCases(false);
      }
    })();

    return () => {
      mounted = false;
    };
  }, [viewMode, showFiledSummary, composeMode, token, settings.caseListScope]);

  const { suggestions: caseSuggestions } = useCaseSuggestions({
    enabled: viewMode === "pickCase" && !(settings.internalEmailHandling === "doNotSuggest" && isInternalEmailDetected),
    emailItemId: activeItemId,
    conversationKey,
    subject: subjectText,
    bodySnippet: suggestBodySnippet,
    fromEmail: suggestionEmail,
    attachments: attachmentsLite,
    cases: visibleCases,
    selectedCaseId,
    selectedSource,
    onAutoSelectCaseId: (id) => {
      if (!settings.rememberLastCase) return;
      setSelectedCaseId(id);
      setSelectedSource("suggested");
    },
    topK: 3,
  });

  React.useEffect(() => {
    if (!detectedFrequentCaseId) return;

    if (!settings.rememberLastCase) {
      // OFF: surface as a suggestion, never touch selectedCaseId
      setSuggestedCaseId(detectedFrequentCaseId);
      setSuggestedConfidencePct(85); // conversation-based match is treated as high-confidence
      return;
    }

    // ON: existing auto-select behaviour
    if (selectedSource === "manual" || selectedSource === "remembered") return;
    if (!isComposeMode()) return;
    if (submailDetectedCaseId) return;

    if (selectedCaseId !== detectedFrequentCaseId) {
      setSelectedCaseId(detectedFrequentCaseId);
      setSelectedSource("suggested");
    }
  }, [detectedFrequentCaseId, settings.rememberLastCase, selectedSource, selectedCaseId, submailDetectedCaseId]);

  // Sync suggestedCaseId/suggestedConfidencePct from content-based suggestions
  // (only while the picker is open and rememberLastCase is OFF).
  React.useEffect(() => {
    if (settings.rememberLastCase) return;
    if (viewMode !== "pickCase") return;
    const top = caseSuggestions[0];
    if (top) {
      setSuggestedCaseId(top.caseId);
      setSuggestedConfidencePct(top.confidencePct);
    }
  }, [caseSuggestions, viewMode, settings.rememberLastCase]);

  React.useEffect(() => {
    let mounted = true;
    let pollId: number | null = null;

    const update = async () => {
      if (!mounted) return;

      if (!isComposeMode()) {
        setComposeRecipientsLive(prev => (prev.length === 0 ? prev : []));
        return;
      }

      const r = await getDraftRecipientsEmailsAsync();
      if (!mounted) return;
      // Stabilize: only update state when content actually changed.
      // Without this, every poll produces a new [] reference → Effect A re-runs
      // → setPrompt() with a new object reference → re-render every 350 ms → flicker.
      setComposeRecipientsLive(prev => {
        if (prev.length === r.length && prev.every((e, i) => e === r[i])) return prev;
        return r;
      });
    };

    const setup = async () => {
      try {
        if (typeof Office?.onReady === "function") await Office.onReady();

        try {
          Office.context.mailbox.addHandlerAsync(Office.EventType.RecipientsChanged, () => {
            void update();
          });
        } catch {
          // ignore
        }

        pollId = window.setInterval(() => {
          void update();
        }, 350);

        void update();
      } catch {
        // ignore
      }
    };

    void setup();

    return () => {
      mounted = false;
      if (pollId) window.clearInterval(pollId);

      try {
        Office.context.mailbox.removeHandlerAsync(
          Office.EventType.RecipientsChanged,
          update as any
        );
      } catch {
        // ignore
      }
    };
  }, [activeItemId]);

  const doSubmitRef = React.useRef<(() => Promise<void>) | null>(null);

  // Auto-detect and auto-select case from SingleCase submail in recipients (highest priority)
  React.useEffect(() => {
    if (!composeMode) {
      // Clear detection when not in compose mode
      if (submailDetectedCaseId) {
        setSubmailDetectedCaseId("");
        setSubmailDetectedCaseKey("");
        setSubmailDetectedCaseName("");
      }
      return;
    }

    const emails = composeRecipientsLive || [];
    if (emails.length === 0 || !workspaceHost) {
      // Clear detection when no recipients or workspace
      if (submailDetectedCaseId) {
        if (VERBOSE_LOGGING) console.log("[submail-detection] Clearing detection (no recipients or workspace)");
        setSubmailDetectedCaseId("");
        setSubmailDetectedCaseKey("");
        setSubmailDetectedCaseName("");
        // Clear selection if it was from submail (not manual)
        if (selectedSource === "suggested" && selectedCaseId === submailDetectedCaseId) {
          setSelectedCaseId("");
          setSelectedSource("");
        }
      }
      return;
    }

    if (VERBOSE_LOGGING) {
      console.log("[submail-detection] Checking recipients for submail", {
        recipientCount: emails.length,
        workspaceHost,
        caseCount: cases.length,
      });
    }

    // Extract case keys from submail addresses
    const caseKeys = extractSubmailCaseKeys(emails, workspaceHost);
    if (caseKeys.length === 0) {
      // Clear detection when submail removed
      if (submailDetectedCaseId) {
        if (VERBOSE_LOGGING) console.log("[submail-detection] Submail removed, clearing detection");
        setSubmailDetectedCaseId("");
        setSubmailDetectedCaseKey("");
        setSubmailDetectedCaseName("");
        // Clear selection if it was from submail (not manual)
        if (selectedSource === "suggested" && selectedCaseId === submailDetectedCaseId) {
          setSelectedCaseId("");
          setSelectedSource("");
        }
      }
      return;
    }

    if (VERBOSE_LOGGING) console.log("[submail-detection] Found case keys in recipients:", caseKeys);

    // Edge case: Multiple different case submails present
    const uniqueCaseKeys = Array.from(new Set(caseKeys));
    if (uniqueCaseKeys.length > 1) {
      if (VERBOSE_LOGGING) console.warn("[submail-detection] Multiple different case submails detected:", uniqueCaseKeys);
      // Clear detection and let user manually select
      setSubmailDetectedCaseId("");
      setSubmailDetectedCaseKey("");
      setSubmailDetectedCaseName("");
      if (selectedSource === "suggested" && selectedCaseId === submailDetectedCaseId) {
        setSelectedCaseId("");
        setSelectedSource("");
      }
      return;
    }

    // Try to resolve the first case key to a case
    const resolved = resolveSubmailCaseKey(caseKeys[0], cases);
    if (!resolved) {
      if (VERBOSE_LOGGING) console.log("[submail-detection] Could not resolve case key to case:", caseKeys[0]);
      if (submailDetectedCaseId) {
        setSubmailDetectedCaseId("");
        setSubmailDetectedCaseKey("");
        setSubmailDetectedCaseName("");
        // Clear selection if it was from submail
        if (selectedSource === "suggested" && selectedCaseId === submailDetectedCaseId) {
          setSelectedCaseId("");
          setSelectedSource("");
        }
      }
      return;
    }

    if (VERBOSE_LOGGING) console.log("[submail-detection] Auto-selecting case from submail:", resolved);

    // Auto-select the case (highest priority - overrides all other suggestions)
    setSubmailDetectedCaseId(resolved.caseId);
    setSubmailDetectedCaseKey(resolved.caseKey);
    setSubmailDetectedCaseName(resolved.caseName);

    // Auto-select if not manually selected
    if (selectedSource !== "manual") {
      setSelectedCaseId(resolved.caseId);
      setSelectedSource("suggested"); // Use "suggested" to indicate automatic selection
      if (!autoFileUserSet) setAutoFileOnSend(true);
    }
  }, [
    composeMode,
    composeRecipientsLive,
    workspaceHost,
    cases,
    selectedSource,
    autoFileUserSet,
    submailDetectedCaseId,
    selectedCaseId,
  ]);

  // Effect: Detect internal emails in COMPOSE mode.
  // Keyed only on live recipients so it never re-runs on read-mode item changes.
  React.useEffect(() => {
    if (!composeMode) return;

    const emails = composeRecipientsLive || [];

    if (emails.length === 0) {
      setIsInternalEmailDetected(false);
      return;
    }

    const senderEmail = String(Office?.context?.mailbox?.userProfile?.emailAddress || "");
    if (!senderEmail) {
      setIsInternalEmailDetected(false);
      return;
    }

    const isInternal = isInternalEmail(senderEmail, emails);
    console.log("[internal-detect] compose", {
      senderEmail,
      recipientCount: emails.length,
      isInternal,
    });
    setIsInternalEmailDetected(isInternal);
  }, [composeMode, composeRecipientsLive]);

  // Effect: Detect internal emails in READ mode.
  // Keyed on activeItemId / activeItemKey so it runs whenever a new message is opened.
  React.useEffect(() => {
    if (composeMode) return;

    if (!activeItemId) {
      setIsInternalEmailDetected(false);
      return;
    }

    const userEmail = String(Office?.context?.mailbox?.userProfile?.emailAddress || "");
    if (!userEmail) {
      setIsInternalEmailDetected(false);
      return;
    }

    const fromEmail = getOutlookFromEmail();
    const recipientEmails = getReadModeRecipientEmails();

    const allParticipants = Array.from(new Set(
      [fromEmail, ...recipientEmails].map(e => normEmail(e)).filter(Boolean)
    ));

    if (allParticipants.length === 0) {
      setIsInternalEmailDetected(false);
      return;
    }

    const isInternal = isInternalEmail(userEmail, allParticipants);
    console.log("[internal-detect] read", {
      userEmail,
      participantCount: allParticipants.length,
      isInternal,
    });
    setIsInternalEmailDetected(isInternal);

    // Immediately update a "Checking…" or generic unfiled prompt to reflect internal status
    const showBanner = isInternal && settings.internalEmailHandling === "doNotSuggest";
    setPrompt(prev => {
      if (prev.kind === "unfiled" && (prev.text === "Checking..." || prev.text.includes("isn't filed yet"))) {
        return {
          ...prev,
          text: showBanner
            ? "This looks like an internal email."
            : "This email isn't filed yet. Would you like me to file it to a case?",
        };
      }
      return prev;
    });
  }, [composeMode, activeItemId, activeItemKey, settings.internalEmailHandling]);

  // Effect: Preselect the last chosen case in read mode when "Remember last selected case" is ON.
  // Priority: below submail/frequent-case (those are compose-only) and below internal-email
  // defaultToLastCase handling. Only fires when no case has been selected yet (selectedSource === "").
  React.useEffect(() => {
    if (composeMode) return;
    if (!settings.rememberLastCase) return;
    if (selectedSource !== "") return; // don't override any existing selection

    const lastCase = loadLastCaseId();
    if (!lastCase) return;

    // Only apply if the case still exists in the loaded cases list
    if (!(cases || []).some((c: any) => String(c?.id) === String(lastCase))) return;

    setSelectedCaseId(lastCase);
    setSelectedSource("last_case");
  }, [composeMode, settings.rememberLastCase, selectedSource, cases]);

  // Effect: Check if email is already filed (read mode only, using conversationId + subject)
  React.useEffect(() => {
    // Only run in read mode
    if (composeMode || !activeItemId) {
      if (filedStatusChecked) {
        setFiledStatusChecked(false);
        setAlreadyFiled(false);
        setAlreadyFiledCaseId("");
        setAlreadyFiledCaseLabel("");
        setAlreadyFiledDocumentId("");
      }
      return;
    }

    // Skip only if viewMode is "sent" AND we've already completed a fresh check for this email.
    // Do NOT skip if filedStatusChecked is false (means a new email was selected and evaluateItem
    // may have set viewMode="sent" with stale storage data before checkIfFiled had a chance to run).
    if (viewMode === "sent" && filedStatusChecked) {
      console.log("[checkIfFiled] Skipping check - already verified for this email");
      return;
    }

    // Reset filed status check for new email
    if (filedStatusChecked) {
      setFiledStatusChecked(false);
      setAlreadyFiled(false);
      setAlreadyFiledCaseId("");
      setAlreadyFiledCaseLabel("");
      setAlreadyFiledDocumentId("");
      setDuplicateFilingWarning(false);
    }

    // Only check if authenticated
    if (!token) {
      return;
    }

    const checkIfFiled = async () => {
      try {
        console.log("[checkIfFiled] Starting CACHE-BASED filed status check");

        // Step 1: Get email identifiers from Office API
        const item = Office?.context?.mailbox?.item as any;
        const conversationId = String(item?.conversationId || "").trim();
        const subject = String(item?.subject || "").trim();

        console.log("[checkIfFiled] Email identifiers:", {
          activeItemId,
          conversationId: conversationId ? conversationId.substring(0, 30) + "..." : "MISSING",
          subject: subject || "MISSING",
        });

        // Step 2: Check local cache (conversationId-based and subject-based)
        const { getFiledEmailFromCache, findFiledEmailBySubject } = await import("../../../utils/filedCache");

        let cacheEntry = null;

        // Try conversationId lookup first
        if (conversationId) {
          cacheEntry = await getFiledEmailFromCache(conversationId);
          console.log("[checkIfFiled] Cache lookup by conversationId:", cacheEntry ? "FOUND" : "NOT FOUND");
        }

        // Fallback to subject lookup
        if (!cacheEntry && subject) {
          cacheEntry = await findFiledEmailBySubject(subject, conversationId);
          console.log("[checkIfFiled] Cache lookup by subject:", cacheEntry ? "FOUND" : "NOT FOUND");
        }

        // Step 3: Process cache result
        if (!cacheEntry) {
          console.log("[checkIfFiled] ❌ Not found in cache");
          setFiledStatusChecked(true);
          setAlreadyFiled(false);
          setAlreadyFiledCaseId("");
          setAlreadyFiledCaseLabel("");
          setAlreadyFiledDocumentId("");
          return;
        }

        // Email is filed!
        console.log("[checkIfFiled] ✅ Found in cache:", cacheEntry);

        const caseLabel = cacheEntry.caseKey && cacheEntry.caseName
          ? `${cacheEntry.caseKey} · ${cacheEntry.caseName}`
          : cacheEntry.caseName || cacheEntry.caseKey || cacheEntry.caseId || "Unknown case";

        // Step 4: Update state for FiledSummary rendering
        setAlreadyFiled(true);
        setAlreadyFiledCaseId(cacheEntry.caseId);
        setAlreadyFiledCaseLabel(caseLabel);
        setAlreadyFiledDocumentId(cacheEntry.documentId);

        // Update sentPill (for FiledSummary which expects it)
        const filedAtIso = new Date(cacheEntry.filedAt).toISOString();
        console.log("[checkIfFiled] Updating sentPill for FiledSummary rendering");
        setSentPill({
          sent: true,
          caseId: cacheEntry.caseId,
          documentId: cacheEntry.documentId,
          atIso: filedAtIso,
        });

        // Persist sentPill for future loads (optional cache)
        try {
          await saveSentPill(activeItemId, {
            sent: true,
            caseId: cacheEntry.caseId,
            documentId: cacheEntry.documentId,
            atIso: filedAtIso,
          });
          console.log("[checkIfFiled] Persisted sentPill to storage");
        } catch (e) {
          console.warn("[checkIfFiled] Failed to persist sentPill (non-critical):", e);
        }

        // Step 5: Build document list for FiledSummary (without Graph API)
        if (workspaceHost && cacheEntry.documentId) {
          console.log("[checkIfFiled] Building document list for documentId:", cacheEntry.documentId);

          // Build document URL (no API call needed)
          const docUrl = `https://${workspaceHost}/documents/${cacheEntry.documentId}`;

          // Create document item — use the stored subject as the document name
          const docItem: UploadedItem = {
            id: cacheEntry.documentId,
            name: cacheEntry.subject
              ? `${cacheEntry.subject}.eml`
              : cacheEntry.caseName
                ? `Email - ${cacheEntry.caseName}`
                : "Email",
            url: docUrl,
            kind: "email",
            atIso: filedAtIso,
          };

          console.log("[checkIfFiled] Document item created:", docItem);

          // Load existing links and merge (don't replace - user may have just filed attachments)
          try {
            const existing = await loadUploadedLinks(activeItemId);
            console.log("[checkIfFiled] Existing uploaded links:", existing.length);

            // Merge: Keep email document + any attachments
            const merged = [
              docItem,
              ...existing.filter((x: any) => String(x?.id) !== String(docItem.id) && x?.kind !== "email"),
            ].slice(0, 25);

            console.log("[checkIfFiled] Merged document list:", merged.length, "items");
            setUploadedLinksValidated(merged as any);

            await saveUploadedLinks(activeItemId, merged as any);
            console.log("[checkIfFiled] Persisted merged document links to storage");
          } catch (e) {
            console.warn("[checkIfFiled] Failed to load/merge document links:", e);
            // Fallback: just set the email
            setUploadedLinksValidated([docItem]);
          }
        }

        // Step 6: Apply category via Office.js (not Graph API)
        try {
          console.log("[checkIfFiled] Applying filed category via Office.js");
          const { applyFiledCategoryToCurrentEmailOfficeJs } = await import("../../../services/graphMail");
          await applyFiledCategoryToCurrentEmailOfficeJs();
          console.log("[checkIfFiled] Category applied successfully");
        } catch (e) {
          console.warn("[checkIfFiled] Failed to apply category (non-critical):", e);
        }

        setFiledStatusChecked(true);
      } catch (e) {
        console.error("[checkIfFiled] Unexpected error:", e);
        setFiledStatusChecked(true);
        setAlreadyFiled(false);
      }
    };

    checkIfFiled();
  }, [composeMode, activeItemId, token, workspaceHost, viewMode]);

  // Update UI when filed status is detected
  React.useEffect(() => {
    // GATE: Don't show filed UI while still loading
    if (isItemLoading) return;

    if (alreadyFiled && alreadyFiledCaseLabel && !composeMode && activeItemId) {
      console.log("[useEffect:alreadyFiled] Updating UI for filed email", {
        caseLabel: alreadyFiledCaseLabel,
        documentId: alreadyFiledDocumentId,
      });

      setViewMode("prompt");
      setPickStep("case");
      setQuickActions([
        { id: "mark_unfiled", label: "Mark as SC: Unfiled", intent: "mark_as_unfiled" },
      ]);
      setPrompt({
        itemId: activeItemId,
        kind: "filed",
        text: "",
      });
    }
  }, [alreadyFiled, alreadyFiledCaseLabel, alreadyFiledDocumentId, composeMode, activeItemId, isItemLoading]);

  // Effect: Check for divergence after filing status is determined
  React.useEffect(() => {
    // Only run in read mode after filing status check completes
    if (composeMode || !filedStatusChecked || !activeItemId) {
      return;
    }

    // DISABLED: Divergence detection was causing false positives for old filed emails
    // If an email has "SC: Filed" category, trust it (even without cache entry)
    // Cache-based approach isn't reliable due to storage limits and old filings
    /*
    void (async () => {
      try {
        const filedCatNorm = normaliseCat(FILED_CATEGORY);
        const unfiledCatNorm = normaliseCat(UNFILED_CATEGORY);
        const cats = await getCurrentMailCategoriesNorm(filedCatNorm, unfiledCatNorm);
        const hasFiledCategory = cats.includes(filedCatNorm);

        console.log("[divergence-check] Checking for divergence:", {
          hasFiledCategory,
          alreadyFiled,
          filedStatusChecked,
        });

        if (hasFiledCategory && !alreadyFiled) {
          console.warn("[divergence-check] DIVERGENCE DETECTED");
          setFilingDivergenceDetected(true);
          setViewMode("prompt");
          setPrompt({
            itemId: activeItemId,
            kind: "unfiled",
            text: "This email is marked as filed in Outlook, but no filing record was found in SingleCase.",
          });
          setQuickActions([
            { id: "refresh", label: "Refresh", intent: "refresh_filing_status" },
            { id: "mark_unfiled", label: "Mark as Unfiled", intent: "mark_as_unfiled" },
          ]);
        } else {
          setFilingDivergenceDetected(false);
        }
      } catch (e) {
        console.error("[divergence-check] Error checking divergence:", e);
      }
    })();
    */

    // Instead, just clear divergence flag
    setFilingDivergenceDetected(false);
  }, [composeMode, filedStatusChecked, alreadyFiled, activeItemId]);

  React.useEffect(() => {
    let cancelled = false;

    void (async () => {
      if (!isComposeMode()) {
        setDetectedFrequentCaseId("");
        setDismissedFrequentKey("");
        return;
      }

      // Skip frequent case detection if submail is detected (highest priority)
      if (submailDetectedCaseId) {
        setDetectedFrequentCaseId("");
        return;
      }

      const emails = composeRecipientsLive || [];
      console.log("[recipientDetection] Recipients polled:", emails);
      if (emails.length === 0) {
        setDetectedFrequentCaseId("");
        return;
      }

      const keyNow = `${activeItemId}:${emails.slice().sort().join("|")}`;
      if (dismissedFrequentKey && dismissedFrequentKey === keyNow) {
        console.log("[recipientDetection] Skipped — dismissed for this recipient set");
        setDetectedFrequentCaseId("");
        return;
      }

      const best = await findBestCaseForRecipients(emails);
      if (cancelled) return;

      if (best) {
        console.log("[recipientDetection] Match found:", { caseId: best.caseId, score: best.score });
      } else {
        console.log("[recipientDetection] No strong match in history for recipients:", emails);
      }
      setDetectedFrequentCaseId(best?.caseId || "");
    })();

    return () => {
      cancelled = true;
    };
  }, [activeItemId, composeRecipientsLive, dismissedFrequentKey, submailDetectedCaseId]);

  // Auto-select case for Reply: prefer conversation mapping, fallback to last viewed filed case
  React.useEffect(() => {
  if (!composeMode) return undefined;

  let cancelled = false;

  void (async () => {
    if (selectedSource === "manual") return;

    // Skip remembered case selection if submail is detected (highest priority)
    if (submailDetectedCaseId) {
      return;
    }

    const recips = Array.isArray(composeRecipientsLive) ? composeRecipientsLive : [];
    const hasRecipients = recips.length > 0;

    // Optional: treat a typed subject/body as a signal too
    const hasSubjectSignal = String(subjectText || "").trim().length > 0;
    const hasBodySignal = String(suggestBodySnippet || "").trim().length > 0;

    // If this is a brand new compose with no signal yet, do nothing.
    // This prevents the "random remembered case" prompt on empty drafts.
    if (!hasRecipients && !hasSubjectSignal && !hasBodySignal) {
      return;
    }

    let remembered = "";
    let rememberedSource: "thread" | "last" | "" = "";

    // 1) Prefer conversation mapping (reply like behaviour)
    if (conversationKey) {
      remembered = await loadConversationFiledCase(conversationKey);
      rememberedSource = remembered ? "thread" : "";
      if (cancelled) return;
    }

    // 2) Fallback to last selected case (if setting is on, no frequent-case suggestion yet,
    //    and there is at least one recipient to confirm this isn't a completely blank draft).
    if (!remembered && settings.rememberLastCase && hasRecipients && !detectedFrequentCaseId) {
      const lastCase = loadLastCaseId();
      if (lastCase && (cases || []).some((c: any) => String(c?.id) === String(lastCase))) {
        remembered = lastCase;
        rememberedSource = "last";
      }
      if (cancelled) return;
    }

    if (!remembered) return;

    setSelectedCaseId(remembered);
    setSelectedSource(rememberedSource === "last" ? "last_case" : "remembered");
if (!autoFileUserSet) setAutoFileOnSend(true);
    setChatStep("compose_ready");

    // Only show the "recognised from a previous email" text when it was really thread based
    const c: any = (cases || []).find((x: any) => String(x?.id) === String(remembered));
    const name = String(c?.name || c?.title || c?.label || "").trim() || `Case ${remembered}`;

    if (rememberedSource === "thread") {
      setPrompt({
        itemId: activeItemId,
        kind: "unfiled",
        text: `The case was recognised from a previous email. Case: ${name}.`,
      });
    }
  })();

  return () => {
    cancelled = true;
  };
}, [
  composeMode,
  conversationKey,
  selectedSource,
  activeItemId,
  cases,
  composeRecipientsLive,
  subjectText,
  suggestBodySnippet,
  detectedFrequentCaseId,
  settings.rememberLastCase,
]);

  const detectedFrequentCaseName = React.useMemo(() => {
    if (!detectedFrequentCaseId) return "";
    const c: any = (cases || []).find((x: any) => String(x?.id) === String(detectedFrequentCaseId));
    const name = String(c?.name || c?.title || c?.label || "").trim();
    return name || `Case ${detectedFrequentCaseId}`;
  }, [cases, detectedFrequentCaseId]);

  React.useEffect(() => {
    if (!composeMode) {
      setChatStep("idle");
      // Don't clear quick actions in read mode - they're managed by evaluateItemForRead
      // and button click handlers (e.g., dismiss, file anyway, etc.)
      return;
    }

    // Filing is disabled — show informational panel
    if (settings.filingOnSend === "off") {
      setChatStep("idle");
      setViewMode("prompt");
      setPickStep("case");
      setQuickActions([{ id: "off1", label: "Turn on 'Warn each time'", intent: "enable_filing_on_send" }]);
      setPrompt({
        itemId: activeItemId,
        kind: "unfiled",
        text: "Auto filing is off. This email will send normally and won't be filed to SingleCase.",
      });
      return;
    }

    const recips = Array.isArray(composeRecipientsLive) ? composeRecipientsLive : [];

    if (recips.length === 0) {
      // Guard: if submail is already detected, keep that state stable
      // (composeRecipientsLive can briefly be [] before the first poll resolves)
      if (submailDetectedCaseId) return;

      setChatStep("compose_wait_recipients");
      setQuickActions([]);
      console.log("[prompt:set] reason=wait_recipients");
      setPrompt({
        itemId: activeItemId,
        kind: "unfiled",
        text: "Add a recipient (To or Cc) and I will suggest a case and the next steps for filing.",
      });
      return;
    }

    // HIGHEST PRIORITY: Submail detection (100% confident)
    if (submailDetectedCaseId) {
      // auto preselect
      if (!selectedCaseId) {
        setSelectedCaseId(submailDetectedCaseId);
        setSelectedSource("suggested");
        if (storeKey)
          void saveComposeIntent({
            itemKey: storeKey,
            caseId: submailDetectedCaseId,
            autoFileOnSend: true,
          });
      }

      setChatStep("compose_ready");

      const statusText =
        settings.filingOnSend === "always"
          ? "When you hit Send, I'll file this email to SingleCase automatically."
          : "Heads up — when you hit Send, the email goes through normally. Right after, I'll ask if you want to file it.";

      console.log("[prompt:set] reason=submail", { submailDetectedCaseId, submailDetectedCaseKey });
      setPrompt({
        itemId: activeItemId,
        kind: "unfiled",
        text: `${statusText} Case: ${submailDetectedCaseKey} · ${submailDetectedCaseName}.`,
      });

      setQuickActions([
        { id: "s1", label: "Select a different case", intent: "pick_another_case" },
      ]);
      return;
    }

    // PRIORITY 2: Internal email suppression (setting ON + all recipients internal).
    if (suppressInternalSuggestions) {
      // Don't interrupt the user while they are actively picking a case.
      // viewModeRef reads the current viewMode without adding it to the deps array
      // (adding viewMode as a dep would cause the effect to fight with setViewMode calls).
      if (viewModeRef.current === "pickCase") return;

      setChatStep("compose_ready");
      setViewMode("prompt");
      setPickStep("case");

      if (selectedCaseId) {
        // Branch B: user picked a case manually → show confirmation
        const c: any = (cases || []).find((x: any) => String(x?.id) === String(selectedCaseId));
        const caseName = String(c?.title || c?.name || c?.label || "").trim() || `Case ${selectedCaseId}`;
        const actionLine = settings.filingOnSend === "always"
          ? "When you hit Send, this email will be filed automatically."
          : "Ready to file. You can send the email and file it from your Sent folder afterwards.";
        console.log("[compose-flow] Internal email — case confirmed", { selectedCaseId, caseName });
        setQuickActions([{ id: "fm2", label: "Select a different case", intent: "file_manually" }]);
        setPrompt({
          itemId: activeItemId,
          kind: "unfiled",
          text: `Case selected: ${caseName}. ${actionLine}`,
        });
      } else {
        // Branch A: no case selected yet → show initial warning
        console.log("[compose-flow] Internal email suppressed — showing info banner");
        setQuickActions([{ id: "fm1", label: "File manually", intent: "file_manually" }]);
        setPrompt({
          itemId: activeItemId,
          kind: "unfiled",
          text: "This looks like an internal email. You can still file it manually if needed.",
        });
      }
      return;
    }

    if (detectedFrequentCaseId) {
      if (settings.rememberLastCase) {
        // rememberLastCase ON: existing auto-select behaviour
        if (!selectedCaseId) {
          setSelectedCaseId(detectedFrequentCaseId);
          setSelectedSource("suggested");
          if (storeKey)
            void saveComposeIntent({
              itemKey: storeKey,
              caseId: detectedFrequentCaseId,
              autoFileOnSend: true,
            });
        }

        setChatStep("compose_offer_frequent");

        const statusText =
          settings.filingOnSend === "always"
            ? "When you hit Send, I'll file this email to SingleCase automatically."
            : "Heads up — when you hit Send, the email goes through normally. Right after, I'll ask if you want to file it.";

        setPrompt({
          itemId: activeItemId,
          kind: "unfiled",
          text: `${statusText} Prepared case: ${detectedFrequentCaseName}.`,
        });

        setQuickActions([
          { id: "a2", label: "Select a different case", intent: "pick_another_case" },
        ]);
      } else {
        // rememberLastCase OFF: show picker with the suggestion, never auto-select.
        // If the user already picked a case manually, don't override the compose_ready state.
        if (selectedCaseId) return;
        setViewMode("pickCase");
        setPickStep("case");
        setChatStep("compose_offer_frequent");
      }
      return;
    }

    // If a case is already selected (but not from submail/frequent), don't show "Select a case" message
    if (selectedCaseId) return;

    setChatStep("compose_choose_case");
    setQuickActions([
      { id: "b1", label: "Select case", intent: "pick_another_case" },
    ]);

    setPrompt({
      itemId: activeItemId,
      kind: "unfiled",
      text: settings.filingOnSend === "always"
        ? "Select a case to enable auto filing on send."
        : "Heads up — when you hit Send, the email will send normally. After that, I can file it to SingleCase, but you'll need to pick a case first.",
    });
  }, [
    composeMode,
    selectedCaseId,
    activeItemId,
    composeRecipientsLive,
    detectedFrequentCaseId,
    detectedFrequentCaseName,
    autoFileOnSend,
    submailDetectedCaseId,
    submailDetectedCaseKey,
    submailDetectedCaseName,
    storeKey,
    suppressInternalSuggestions,
    settings.filingOnSend,
    settings.rememberLastCase,
  ]);

React.useEffect(() => {
  if (!composeMode) return;
  if (chatStep !== "compose_ready") return;

  // Do not show case-ready prompt when recipients are still empty — Effect A owns
  // that state and will set chatStep to "compose_wait_recipients". Running here
  // while recipients are empty causes the two effects to fight each other and
  // produces visible flickering every polling cycle.
  const recips = Array.isArray(composeRecipientsLive) ? composeRecipientsLive : [];
  if (recips.length === 0) return;

  // Submail has highest priority — the chat-step effect already set the correct prompt
  if (submailDetectedCaseId) return;

  // Internal email suppression: compose-flow already set the correct prompt; bail here too.
  if (suppressInternalSuggestions) return;

  // Normal flow: show case selection if selectedCaseId exists
  if (!selectedCaseId) return;

  const c: any = (cases || []).find((x: any) => String(x?.id) === String(selectedCaseId));
  const name =
    String(c?.name || c?.title || c?.label || "").trim() || `Case ${selectedCaseId}`;

  setViewMode("prompt");
  setPickStep("case");

  // DO NOT clear actions here, this is where user needs them
  setQuickActions([
    { id: "c1", label: "Select a different case", intent: "pick_another_case" },
  ]);

  console.log("[prompt:set] reason=compose_ready/frequent", { selectedCaseId, chatStep });
  setPrompt({
    itemId: activeItemId,
    kind: "unfiled",
    text: settings.filingOnSend === "always"
      ? `Auto filing is on. When you hit Send, I'll file this email to SingleCase automatically. Case: ${name}.`
      : `Heads up — when you hit Send, I'll let the email go through as normal. Prepared case: ${name}.`,
  });
}, [
  composeMode,
  chatStep,
  selectedCaseId,
  cases,
  activeItemId,
  autoFileOnSend,
  suppressInternalSuggestions,
  submailDetectedCaseId,
  settings.filingOnSend,
]);

  // Check for a pending filing confirmation left by the ItemSend handler (ask mode)
  React.useEffect(() => {
    if (!composeMode) {
      // In read mode, check if there's a pending filing for the current item
      void (async () => {
        try {
          const raw = await getStored("sc_pending_filing");
          if (!raw) { setPendingFiling(null); return; }
          const parsed = JSON.parse(raw);
          if (!parsed?.caseId) { setPendingFiling(null); return; }
          setPendingFiling({
            caseId: String(parsed.caseId || ""),
            subject: String(parsed.subject || ""),
            conversationId: String(parsed.conversationId || ""),
            sentAt: String(parsed.sentAt || ""),
          });
        } catch {
          setPendingFiling(null);
        }
      })();
    } else {
      setPendingFiling(null);
    }
  }, [composeMode, activeItemId]);

  const evaluateItem = React.useCallback(async (itemKey: string) => {
    try {
      if (!isMailItem() || !itemKey) {
        setViewMode("prompt");
        setSelectedCaseId("");
        setSelectedSource("");
        setSentPill(null);
        setPickStep("case");
        setForceUnfiledLabel(false);
        setSelectedAttachments([]);
        setIsUploadingNewVersion(false);
        setPrompt({
          itemId: "",
          kind: "none",
          text: "Select an email and I’ll show you relevant suggestions.",
        });
        return;
      }

      if (isComposeMode()) {
        setSentPill(null);
        setUploadedLinksRaw([]);
        setUploadedLinksValidated([]);
        setPickStep("case");
        setSelectedAttachments([]);
        setIsUploadingNewVersion(false);

        setViewMode("prompt");
        setPrompt({
          itemId: itemKey,
          kind: "unfiled",
          text: "Select a case and then use File Now.",
        });
        return;
      }

      // Priority 1: Already filed (detected via conversationId cache)
      if (alreadyFiled && alreadyFiledCaseLabel) {
        console.log("[evaluateItem] Email already filed", {
          caseId: alreadyFiledCaseId,
          caseLabel: alreadyFiledCaseLabel,
        });

        setViewMode("prompt");
        setPickStep("case");

        setQuickActions([
          { id: "mark_unfiled", label: "Mark as SC: Unfiled", intent: "mark_as_unfiled" },
        ]);

        setPrompt({
          itemId: itemKey,
          kind: "filed",
          text: "",
        });

        return;
      }

      if (dismissedRef.current.has(itemKey)) {
        // User previously dismissed — show dismissal message with manual filing option
        setViewMode("prompt");
        setSelectedCaseId("");
        setSelectedSource("");
        setSentPill(null);
        setPickStep("case");
        setSelectedAttachments([]);
        setIsUploadingNewVersion(false);
        setQuickActions([
          { id: "fm1", label: "File manually", intent: "file_manually" },
        ]);
        setPrompt({
          itemId: itemKey,
          kind: "unfiled",
          text: isInternalEmailDetected
            ? "No problem. I'll hide suggestions for this email, but you can still file it anytime."
            : "Got it. I'll step back for this email, but you can still file it later.",
        });
        return;
      }

      const pill = await loadSentPill(itemKey);
      setSentPill(pill || null);

      // Check Outlook categories
      const filedCatNorm = normaliseCat(FILED_CATEGORY);
      const unfiledCatNorm = normaliseCat(UNFILED_CATEGORY);
      const cats = await getCurrentMailCategoriesNorm(filedCatNorm, unfiledCatNorm);
      const hasFiledCategory = cats.includes(filedCatNorm);
      const hasUnfiledCategory = cats.includes(unfiledCatNorm);

      console.log("[evaluateItem] Category check:", {
        hasFiledCategory,
        hasUnfiledCategory,
        filedStatusChecked,
        alreadyFiled,
      });

      // If explicitly marked unfiled, respect that
      if (hasUnfiledCategory) {
        // Exchange changes the item ID when a category is applied, so dismissedRef keyed by
        // item ID won't match here. Check the conversation-based key instead.
        const convIdForDismiss = String(Office?.context?.mailbox?.item?.conversationId || "").trim();
        const convDismissed = convIdForDismiss && dismissedRef.current.has(`conv:${convIdForDismiss}`);

        if (convDismissed) {
          // User already said "No" this session — show the post-dismissal state, not the prompt.
          setViewMode("prompt");
          setPickStep("case");
          setIsUploadingNewVersion(false);
          setQuickActions([{ id: "fm1", label: "File manually", intent: "file_manually" }]);
          setFilingDivergenceDetected(false);
          setPrompt({
            itemId: itemKey,
            kind: "unfiled",
            text: "Got it. I'll step back for this email, but you can still file it later.",
          });
          return;
        }

        setViewMode("prompt");
        setPickStep("case");
        setIsUploadingNewVersion(false);
        setQuickActions([]);
        setFilingDivergenceDetected(false);
        setPrompt({
          itemId: itemKey,
          kind: "unfiled",
          text: "This email isn't filed yet. Would you like me to file it to a case?",
        });
        return;
      }

      // Check local storage for filing data
      const hasLocalFilingData =
        (pill?.caseId && String(pill.caseId).trim() !== "") ||
        uploadedLinksValidated.length > 0 ||
        (alreadyFiled && alreadyFiledCaseId);

      // State A: Has local storage data - show as filed
      if (hasLocalFilingData) {
        const ok = await hasAnyRealDocuments(pill || null, itemKey);

        if (ok === null) {
          // API unavailable (auth/network error) — cannot verify the document.
          // Trust the local sentPill data (same logic as State B ok===null) and show
          // FiledSummaryCard. If the data were wrong the user can mark as unfiled manually.
          try {
            const caseId = String(pill?.caseId || "").trim();
            const emailDocId = String(pill?.documentId || "").trim();
            if (caseId && emailDocId) await saveLastFiledCtx({ caseId, emailDocId });
          } catch { /* ignore */ }
          try {
            if (pill?.caseId) await saveLastFiledCase(String(pill.caseId));
          } catch { /* ignore */ }
          setViewMode("sent");
          setPickStep("case");
          setIsUploadingNewVersion(false);
          setPrompt({ itemId: itemKey, kind: "filed", text: "This email has been filed." });
          return;
        }

        if (ok === false) {
          // Document confirmed gone (404). Clear local filing data.
          await clearLocalFiling(itemKey);
          setSentPill(null);
          setUploadedLinksRaw([]);
          setUploadedLinksValidated([]);
          setViewMode("prompt");
          setPickStep("case");
          setIsUploadingNewVersion(false);
          if (hasFiledCategory) {
            // Email is still marked as filed in Outlook but its document no longer exists.
            setPrompt({
              itemId: itemKey,
              kind: "deleted",
              text: "This email and its documents have been removed from the SingleCase platform. Do you want to file it again?",
            });
          } else {
            // No Outlook category either — standard deleted message.
            setPrompt({
              itemId: itemKey,
              kind: "deleted",
              text: "I've noticed that this email or its attachments were deleted from SingleCase. Do you want to re-file it?",
            });
          }
          return;
        }

        // ok === true: document verified present
        try {
          const caseId = String(pill?.caseId || "").trim();
          const emailDocId = String(pill?.documentId || "").trim();
          if (caseId && emailDocId) await saveLastFiledCtx({ caseId, emailDocId });
        } catch {
          // ignore
        }
        try {
          if (pill?.caseId) await saveLastFiledCase(String(pill.caseId));
        } catch {
          // ignore
        }
        setViewMode("sent");
        setPickStep("case");
        setIsUploadingNewVersion(false);
        setPrompt({ itemId: itemKey, kind: "filed", text: "This email has been filed." });
        return;
      }

      // State B: Has category but NO local data — try recovery before giving up
      if (hasFiledCategory) {
        let recoveredPill: SentPillData | null = null;
        try {
          const outlookItem = Office?.context?.mailbox?.item as any;
          const convId = String(outlookItem?.conversationId || "").trim();
          if (convId) {
            // 1. Try draft key (email may have been filed from compose)
            const draftPill = await loadSentPill(`draft:${convId}`);
            if (draftPill?.caseId && String(draftPill.caseId).trim() !== "") {
              recoveredPill = draftPill;
              console.log("[evaluateItem] State B recovery: found via draft key");
            }

            // 2. Try conv_ctx (saved at file time for the whole thread)
            if (!recoveredPill) {
              const ctxRaw = await getStored(`${CONV_CTX_KEY_PREFIX}${convId}`);
              if (ctxRaw) {
                const ctx = JSON.parse(String(ctxRaw));
                if (ctx?.caseId && ctx?.emailDocId) {
                  recoveredPill = { sent: true, caseId: ctx.caseId, documentId: ctx.emailDocId, atIso: new Date().toISOString() };
                  console.log("[evaluateItem] State B recovery: found via conv_ctx");
                }
              }
            }

            // 3. Try filedCache lookup directly (bypasses stale checkIfFiled state)
            if (!recoveredPill) {
              const { getFiledEmailFromCache, findFiledEmailBySubject } = await import("../../../utils/filedCache");
              const cacheEntry = await getFiledEmailFromCache(convId);
              if (cacheEntry?.caseId) {
                recoveredPill = { sent: true, caseId: cacheEntry.caseId, documentId: cacheEntry.documentId, atIso: new Date(cacheEntry.filedAt).toISOString() };
                console.log("[evaluateItem] State B recovery: found via filedCache by conversationId");
              }
              // 4. Subject-based cache lookup as last cache attempt
              if (!recoveredPill) {
                const subject = String(outlookItem?.subject || "").trim();
                if (subject) {
                  const subjectEntry = await findFiledEmailBySubject(subject, convId);
                  if (subjectEntry?.caseId) {
                    recoveredPill = { sent: true, caseId: subjectEntry.caseId, documentId: subjectEntry.documentId, atIso: new Date(subjectEntry.filedAt).toISOString() };
                    console.log("[evaluateItem] State B recovery: found via filedCache by subject");
                  }
                }
              }
            }

            // 5. Try conv_case key — saved at file time with caseId only (no emailDocId needed)
            if (!recoveredPill) {
              const convCaseId = await loadConversationFiledCase(convId);
              if (convCaseId) {
                recoveredPill = { sent: true, caseId: convCaseId };
                console.log("[evaluateItem] State B recovery: found via conv_case key (caseId only)");
              }
            }

            // 6. Try conv:${convId} backup key saved during filing (handles item ID change after category apply)
            if (!recoveredPill) {
              const convPill = await loadSentPill(`conv:${convId}`);
              if (convPill?.caseId && String(convPill.caseId).trim() !== "") {
                recoveredPill = convPill;
                console.log("[evaluateItem] State B recovery: found via conv: backup key");
              }
            }
          }
        } catch (e) {
          console.warn("[evaluateItem] State B recovery error:", e);
        }

        if (recoveredPill) {
          // Persist so future loads use the correct key
          await saveSentPill(itemKey, recoveredPill);
          setSentPill(recoveredPill);
          // Do NOT verify documents in State B: the SC API can return 404 for valid documents
          // when the OWA token lacks the right scope, causing false "removed from platform" messages.
          // Recovery data is sourced from data saved at filing time (conv_ctx, conv_case, conv: key)
          // and the email has the SC:Filed category — trust it unconditionally.
          try {
            const caseId = String(recoveredPill.caseId || "").trim();
            const emailDocId = String(recoveredPill.documentId || "").trim();
            if (caseId && emailDocId) await saveLastFiledCtx({ caseId, emailDocId });
          } catch { /* ignore */ }
          setViewMode("sent");
          setPickStep("case");
          setIsUploadingNewVersion(false);
          setPrompt({ itemId: itemKey, kind: "filed", text: "This email has been filed." });
          return;
        }

        // All recovery paths exhausted — filing details could not be recovered.
        // Show the standard filing UI so the user can re-file and restore the data.
        setViewMode("prompt");
        setPickStep("case");
        setIsUploadingNewVersion(false);
        setPrompt({
          itemId: itemKey,
          kind: "deleted",
          text: "This email and its documents have been removed from the SingleCase platform. Do you want to file it again?",
        });
        setQuickActions([]);
        return;
      }

      // State C: Not filed - normal filing UI

      setViewMode("prompt");
      setPickStep("case");
      setIsUploadingNewVersion(false);
      setQuickActions([]);


      if (!isComposeMode()) {
        const userEmailNow = String(Office?.context?.mailbox?.userProfile?.emailAddress || "");
        const fromEmailNow = getOutlookFromEmail();
        const recipientsNow = getReadModeRecipientEmails();
        const participantsNow = Array.from(new Set(
          [fromEmailNow, ...recipientsNow].map(e => normEmail(e)).filter(Boolean)
        ));
        const internalNow =
          userEmailNow && participantsNow.length > 0
            ? isInternalEmail(userEmailNow, participantsNow)
            : false;

        setPrompt({
          itemId: itemKey,
          kind: "unfiled",
          text: internalNow
            ? "This looks like an internal email."
            : "This email isn't filed yet. Would you like me to file it to a case?",
        });
      } else {
        setPrompt({
          itemId: itemKey,
          kind: "unfiled",
          text: "This email isn't filed yet. Would you like me to file it to a case?",
        });
      }
    } catch {
      setViewMode("prompt");
      setPickStep("case");
      setIsUploadingNewVersion(false);
      setPrompt({ itemId: "", kind: "none", text: "Select an email and I’ll show you relevant suggestions." });
    }
  }, []);

  React.useEffect(() => {
    let mounted = true;
    let intervalId: number | null = null;

    const sync = async () => {
      if (!mounted) return;
      const itemKey = await getCurrentItemKey();

      
const realItemId = getCurrentItemIdSafe();
const modeNow = isComposeMode();
const subjNow = await getOutlookSubjectAsync().catch(() => "");
const convNow = String(Office?.context?.mailbox?.item?.conversationId || "");

if (VERBOSE_LOGGING) {
  console.log("[sync] snapshot", {
    modeNow,
    realItemId,
    itemKey,
    activeItemId,
    activeItemKey,
    convNow,
    subjNow: String(subjNow || "").slice(0, 80),
  });
}
const nextActiveId = realItemId || "";
const nextActiveKey = itemKey;

// IMPORTANT: trigger when either changes
const activeChanged =
  String(nextActiveId || "") !== String(activeItemId || "") ||
  String(nextActiveKey || "") !== String(activeItemKey || "");

if (itemKey && activeChanged) {
  console.log("[active-item-change] Resetting filed state for new email", {
    activeItemKey: nextActiveKey,
    activeItemId: nextActiveId,
  });

  setIsItemLoading(true);
  setSentPill(null);
  setUploadedLinksRaw([]);
  setUploadedLinksValidated([]);
  setSubmitError("");
  setReplyBaseCaseId("");
  setReplyBaseEmailDocId("");

  // CRITICAL: Clear filed status to prevent stale UI during loading
  setFiledStatusChecked(false);
  setAlreadyFiled(false);
  setAlreadyFiledCaseId("");
  setAlreadyFiledCaseLabel("");
  setAlreadyFiledDocumentId("");
  setQuickActions([]);

  setActiveItemId(nextActiveId);
  setActiveItemKey(nextActiveKey);

  setSelectedCaseId("");
  setSelectedSource("");
  setSuggestedCaseId("");
  setSuggestedConfidencePct(0);
  setFilingTargetCaseId("");
  setPickStep("case");
  setSelectedAttachments([]);
  setIsUploadingNewVersion(false);

const nextStoreKey = getStoreKey(nextActiveKey, nextActiveId, modeNow);
await evaluateItem(nextStoreKey);
  setIsItemLoading(false);
  console.log("[active-item-change] Evaluation complete", {
    activeItemKey: nextActiveKey,
    confirmedFiled: alreadyFiled,
  });
  return;
}

      // also keep activeItemKey updated even when id didn’t change
      if (itemKey && itemKey !== activeItemKey) {
        setActiveItemKey(itemKey);
      }
    };

    const onItemChanged = () => {
      void sync();
    };

    const onFocusOrVisible = () => {
      void sync();
    };

    const setup = async () => {
      try {
        if (typeof Office?.onReady === "function") await Office.onReady();

        try {
          Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
        } catch {
          // ignore
        }

        window.addEventListener("focus", onFocusOrVisible);
        document.addEventListener("visibilitychange", onFocusOrVisible);

        intervalId = window.setInterval(() => {
          void sync();
        }, 450);

        await sync();
      } catch {
        // ignore
      }
    };

    void setup();

    return () => {
      mounted = false;
      try {
        window.removeEventListener("focus", onFocusOrVisible);
        document.removeEventListener("visibilitychange", onFocusOrVisible);
      } catch {
        // ignore
      }

      if (intervalId) window.clearInterval(intervalId);

      try {
        try {
          Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged, {
            handler: onItemChanged,
          });
        } catch {
          Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
        }
      } catch {
        // ignore
      }
    };
  }, [activeItemId, activeItemKey, evaluateItem]);

  const filedCase = React.useMemo(() => {
    if (!filedCaseId) return null;
    return (cases || []).find((c: any) => String(c?.id) === filedCaseId) || null;
  }, [cases, filedCaseId]);

  const filedCaseName = React.useMemo(() => {
    const c: any = filedCase;
    const name = String(c?.name || c?.title || c?.label || "").trim();
    if (name) return name;
    return filedCaseId ? `Case ${filedCaseId}` : "";
  }, [filedCase, filedCaseId]);

  const caseUrl = React.useMemo(() => {
    if (!filedCaseId) return "";
    const apiUrl = tryGetCaseUrlFromCaseOption(workspaceHost, filedCase);
    if (apiUrl) return apiUrl;
    return buildCaseUrl(workspaceHost, filedCaseId);
  }, [filedCaseId, filedCase, workspaceHost]);

  const documentsToShow = React.useMemo(() => {
    const out: UploadedItem[] = [];

    if (sentPill?.documentId) {
      const id = String(sentPill.documentId);
      const fromStore = Array.isArray(uploadedLinksValidated)
        ? uploadedLinksValidated.find((x) => String(x?.id) === id)
        : null;
      if (fromStore?.url) out.push(fromStore);
    }

    if (Array.isArray(uploadedLinksValidated)) {
      for (const it of uploadedLinksValidated) {
        if (!it?.id || !it?.url) continue;
        out.push(it);
      }
    }

    const seen = new Set<string>();
    const deduped = out.filter((x) => {
      if (seen.has(x.id)) return false;
      seen.add(x.id);
      return true;
    });

    return deduped.slice(0, 25);
  }, [sentPill, uploadedLinksValidated]);

  const closeRename = () => {
    setRenameOpen(false);
    setRenameDoc(null);
    setRenameValue("");
    setRenameSaving(false);
  };

  const confirmRename = async () => {
    if (!renameDoc) return;

    const newName = String(renameValue || "").trim();
    if (!newName) return;
    if (newName === renameDoc.name) {
      closeRename();
      return;
    }

    setRenameSaving(true);
    try {
      const existingRaw = await loadUploadedLinks(storeKey);
      const existing: UploadedItem[] = Array.isArray(existingRaw)
        ? (existingRaw as UploadedItem[])
        : [];

      const updated: UploadedItem[] = existing.map((x) =>
        String(x.id) === String(renameDoc.id) ? { ...x, name: newName } : x
      );

      await saveUploadedLinks(storeKey, updated as any);
      setUploadedLinksRaw(updated);
      setUploadedLinksValidated(updated);

      closeRename();
    } catch (e) {
      console.error(e);
      setRenameSaving(false);
    }
  };

const handleQuickAction = React.useCallback((intent) => {
  console.log("[handleQuickAction] enter", intent);

  if (intent === "enable_filing_on_send") {
    onChangeSettings((prev) => ({ ...prev, filingOnSend: "warn" }));
    return;
  }

  if (intent === "confirm_file_now") {
    if (!pendingFiling) return;
    const caseId = pendingFiling.caseId;
    void removeStored("sc_pending_filing");
    setPendingFiling(null);
    setSelectedCaseId(caseId);
    setSelectedSource("manual");
    // Use the ref so we don't create a circular dep with doSubmit
    void doSubmitOverrideRef.current?.({ caseId });
    return;
  }

  if (intent === "skip_pending_filing") {
    void removeStored("sc_pending_filing");
    setPendingFiling(null);
    return;
  }

      if (intent === "toggle_auto_file") {
  setAutoFileUserSet(true);

  setAutoFileOnSend((v) => {
    const next = !v;

    // Persist per-draft choice so it survives closing and reopening the draft
    if (storeKey) setDraftAutoFile(storeKey, next);

    if (composeMode && selectedCaseId && storeKey) {
      if (next) {
        void saveComposeIntent({
          itemKey: storeKey,
          caseId: selectedCaseId,
          autoFileOnSend: true,
        });
      } else {
        void clearComposeIntent(storeKey);
      }
    }

    setQuickActions((prev) =>
      (prev || []).map((a) =>
        a.intent === "toggle_auto_file"
          ? {
              ...a,
              label: next ? "Disable auto file on send" : "Enable auto file on send",
            }
          : a
      )
    );

    return next;
  });

  return;
}

      if (intent === "file_now") {
  void doSubmitOverrideRef.current?.({ caseId: selectedCaseId });
  return;
}

      if (intent === "file_manually") {
  setViewMode("pickCase");
  setPickStep("case");
  setSelectedCaseId("");
  setSelectedSource("");
  setSelectedAttachments([]);
  setIsUploadingNewVersion(false);
  return;
}

      if (intent === "show_suggestions") {
  dismissedRef.current.delete(storeKey);
  dismissedRef.current.delete(activeItemKey);
  dismissedRef.current.delete(activeItemId);

  setViewMode("pickCase");
  setPickStep("case");
  setSelectedCaseId("");
  setSelectedSource("");
  setSelectedAttachments([]);
  setIsUploadingNewVersion(false);
  return;
}

      if (intent === "cancel_compose") {
setSelectedCaseId("");
setSelectedSource("manual"); // important

  // keep user's preference, don't force it off
  // setAutoFileOnSend(false);

  if (storeKey) void clearComposeIntent(storeKey);

  // take the user straight to case picker
  setViewMode("pickCase");
  setPickStep("case");
  setChatStep("compose_choose_case");
  setQuickActions([]);
  setPrompt({
    itemId: activeItemId,
    kind: "unfiled",
    text: "Select a case for this draft email.",
  });
  return;
}

      if (intent === "accept_frequent_case") {
        if (!detectedFrequentCaseId) return;

        if (settings.rememberLastCase) {
          // ON: write to selectedCaseId (existing behaviour)
          setSelectedCaseId(detectedFrequentCaseId);
          setSelectedSource("suggested");

          if (composeMode) {
            setViewMode("prompt");
            setPickStep("case");
            setQuickActions([]);
            setChatStep("compose_ready");
            if (storeKey)
              void saveComposeIntent({
                itemKey: storeKey,
                caseId: detectedFrequentCaseId,
                autoFileOnSend,
              });

            setPrompt({
              itemId: activeItemId,
              kind: "unfiled",
              text: `OK. Click File Now. Then you can send the email normally. Case: ${detectedFrequentCaseName}.`,
            });
            return;
          }

          setChatStep("idle");
          setViewMode("pickCase");
          setPickStep(attachmentsLite.length > 0 ? "attachments" : "case");
          return;
        }

        // OFF: accept means "user confirmed suggestion" → set filingTargetCaseId, not selectedCaseId
        setFilingTargetCaseId(detectedFrequentCaseId);

        if (composeMode) {
          setViewMode("prompt");
          setPickStep("case");
          setQuickActions([]);
          setChatStep("compose_ready");
          if (storeKey)
            void saveComposeIntent({
              itemKey: storeKey,
              caseId: detectedFrequentCaseId,
              autoFileOnSend,
            });

          setPrompt({
            itemId: activeItemId,
            kind: "unfiled",
            text: `OK. Click File Now. Then you can send the email normally. Case: ${detectedFrequentCaseName}.`,
          });
          return;
        }

        setChatStep("idle");
        setViewMode("pickCase");
        setPickStep(attachmentsLite.length > 0 ? "attachments" : "case");
        return;
      }

      if (intent === "pick_another_case") {
  // Check if we should trigger content-based suggestions
  const wasAutoSelected = selectedSource === "suggested" || selectedSource === "remembered";

  console.log("[pick_another_case] 🔵 Button clicked", {
    wasAutoSelected,
    selectedSource,
    visibleCasesCount: visibleCases.length,
  });

  if (composeMode) {
    setSelectedCaseId("");
    setSelectedSource("manual");

    // Trigger content-based suggestions if case was auto-selected
    if (wasAutoSelected) {
      // Fetch FRESH content from Outlook instead of using stale state
      void (async () => {
        try {
          const freshSubject = await getOutlookSubjectAsync();
          const freshBody = await getEmailBodySnippet(600);

          const hasContent = freshSubject.trim() || freshBody.trim();

          console.log("[pick_another_case] 📥 Fetched fresh content:", {
            subjectLength: freshSubject.length,
            bodyLength: freshBody.length,
            hasContent: !!hasContent,
          });

          if (!hasContent) {
            console.log("[pick_another_case] ⏭️ Skipping analysis: no content");
            setContentBasedSuggestions([]);
            setPrompt({ itemId: activeItemId, kind: "unfiled", text: "Select case." });
            return;
          }

          console.log("[pick_another_case] ✅ Triggering content analysis");

          const result = suggestCasesByContent({
            subject: freshSubject,
            bodySnippet: freshBody,
            cases: visibleCases,
            topK: 5,
          });

          console.log("[pick_another_case] 📊 Analysis complete:", {
            foundCount: result.suggestions.length,
            suggestions: result.suggestions.map(s => ({
              caseId: s.caseId,
              pct: s.confidencePct,
              score: s.score,
              reasons: s.reasons,
            })),
          });

          setContentBasedSuggestions(result.suggestions);

          if (result.suggestions.length === 0) {
            console.log("[pick_another_case] ❌ No suggestions found");
            setPrompt({ itemId: activeItemId, kind: "unfiled", text: "No suggestions based on content. Select case manually." });
          } else {
            console.log("[pick_another_case] ✅ Set", result.suggestions.length, "suggestions in state");
            setPrompt({ itemId: activeItemId, kind: "unfiled", text: "Select case based on email content." });
          }
        } catch (error) {
          console.error("[pick_another_case] ❌ Analysis failed:", error);
          setContentBasedSuggestions([]);
          setPrompt({ itemId: activeItemId, kind: "unfiled", text: "Select case." });
        }
      })();
    } else {
      console.log("[pick_another_case] ⏭️ Skipping analysis: not auto-selected");
      setContentBasedSuggestions([]);
      setPrompt({ itemId: activeItemId, kind: "unfiled", text: "Select case." });
    }
  }

  setViewMode("pickCase");
  setPickStep("case");
  setChatStep("compose_choose_case");
  setQuickActions([]);
  return;
      }

      if (intent === "skip_attachments") {
        setFilingMode("both");
        setSelectedAttachments([]);
        setChatStep("compose_ready");
        setQuickActions([]);
        setPrompt({
          itemId: activeItemId,
          kind: "unfiled",
          text: "OK. Click File Now. I will file only the email.",
        });
        return;
      }

      if (intent === "select_attachments") {
        setFilingMode("both");
        setSelectedAttachments(attachmentIds);
        setChatStep("compose_ready");
        setQuickActions([]);
        setPrompt({
          itemId: activeItemId,
          kind: "unfiled",
          text: "OK. Click File Now. I will file the email and attachments.",
        });
        return;
      }

      if (intent === "view_in_singlecase") {
        if (alreadyFiledDocumentId && workspaceHost) {
          // Open document in SingleCase web app
          const url = `https://${workspaceHost}/documents/${alreadyFiledDocumentId}`;
          console.log("[handleQuickAction] Opening SingleCase document:", url);
          window.open(url, "_blank");
        } else if (alreadyFiledCaseId && workspaceHost) {
          // Fallback: open case page
          const url = `https://${workspaceHost}/cases/${alreadyFiledCaseId}`;
          console.log("[handleQuickAction] Opening SingleCase case:", url);
          window.open(url, "_blank");
        }
        return;
      }

      if (intent === "file_now_from_weak_signal") {
        console.log("[handleQuickAction] File now - clearing weak signal state");
        // Clear any filed signals and show normal filing UI
        setFiledStatusChecked(false);
        setAlreadyFiled(false);
        setAlreadyFiledCaseId("");
        setSentPill(null);
        setViewMode("prompt");
        setPickStep("case");
        setQuickActions([]);
        setPrompt({
          itemId: activeItemKey,
          kind: "none",
          text: "Select a case for this email.",
        });
        return;
      }

      if (intent === "refresh_filing_status") {
        console.log("[handleQuickAction] Refreshing filing status");
        // Reset filed status check so it runs again
        setFiledStatusChecked(false);
        setFilingDivergenceDetected(false);
        // Trigger re-evaluation of the item
        void evaluateItem(activeItemKey);
        return;
      }

      if (intent === "mark_as_unfiled") {
        console.log("[handleQuickAction] Marking as SC: Unfiled (category only, no re-filing)");
        setFilingDivergenceDetected(false);
        void (async () => {
          try {
            setIsTogglingCategory(true);
            await applyUnfiledCategoryToCurrentEmailOfficeJs();
            // Update label — email is still filed in SingleCase, only Outlook
            // category is changing.  Do NOT call evaluateItem() as that would
            // detect "no local filing data" and restart the filing prompt.
            setForceUnfiledLabel(true);
            setAlreadyFiled(false);
            setViewMode("sent"); // keep FiledSummaryCard visible
          } catch (e) {
            console.error("[handleQuickAction] Failed to apply SC: Unfiled category:", e);
          } finally {
            setIsTogglingCategory(false);
          }
        })();
        return;
      }
    },
    [
      activeItemId,
      attachmentIds,
      attachmentsLite.length,
      composeRecipientsLive,
      detectedFrequentCaseId,
      composeMode,
      detectedFrequentCaseName,
      detectedFrequentCaseName,
      detectedFrequentCaseId,
      selectedCaseId,
      storeKey,
      autoFileOnSend,
      alreadyFiledDocumentId,
      alreadyFiledCaseId,
      workspaceHost,
      pendingFiling,
      onChangeSettings,
    ]
  );

  const doSubmit = React.useCallback(
    async (override?: { caseId?: string }) => {
      if (isSubmitting) return;

      const caseId = String(override?.caseId || selectedCaseId || "").trim();
      if (!caseId) return;

      // ── Duplicate detection ──────────────────────────────────────────────────
      // Use a content fingerprint (subject + sender + body) so that two separate
      // emails with identical content are caught, even though their itemIds and
      // conversationIds differ (e.g. sent copy vs received copy).
      const dupFingerprint = computeEmailFingerprint(
        subjectText || "",
        fromEmail || "",
        suggestBodySnippet || ""
      );
      const dupCache = loadDuplicateCache();
      const isDuplicate =
        (dupFingerprint !== "" && hasAttached(dupCache, caseId, dupFingerprint)) ||
        (Boolean(activeItemId) && hasAttached(dupCache, caseId, activeItemId));

      if (isDuplicate && settings.duplicates !== "off" && !allowRefilingOverride) {
        if (settings.duplicates === "block") {
          console.log("[doSubmit] Duplicates=block: preventing filing");
          setSubmitError(
            "This email (or one with identical content) has already been filed to this case."
          );
          return;
        }
        // "warn"
        if (!duplicateFilingWarning) {
          console.log("[doSubmit] Duplicates=warn: showing warning on first attempt");
          setDuplicateFilingWarning(true);
          setSubmitError(
            "This email (or one with identical content) was already filed to this case. Click File again to file anyway."
          );
          setIsSubmitting(false);
          setViewMode("pickCase");
          return;
        }
        // Second click: user confirmed — proceed and reset flag
        console.log("[doSubmit] Duplicates=warn: user confirmed re-filing");
        setDuplicateFilingWarning(false);
      } else {
        setDuplicateFilingWarning(false);
      }
      // ─────────────────────────────────────────────────────────────────────────

      if (allowRefilingOverride) {
        console.log("[doSubmit] Refiling override active - bypassing duplicate guard");
      }

      if (
        pickStep === "attachments" &&
        attachmentsLite.length > 0 &&
        selectedAttachments.length === 0
      ) {
        return;
      }

      if (!storeKey) {
        setSubmitError("Chyba: nelze určit klíč emailu.");
        return;
      }

      setIsSubmitting(true);
      setViewMode("sending");
      setSubmitError("");

      try {
        const bodySnippetFull = await getEmailBodySnippet(8000);
        const bodyForEml =
          (bodySnippetFull || suggestBodySnippet || "").trim() ||
          "[No body content available from Outlook]";

        // Read the subject fresh at filing time — the subjectText state may not yet be
        // populated if the user filed quickly after selecting the email.
        const freshSubject = await getOutlookSubjectAsync().catch(() => "");
        const baseName = safeFileName(freshSubject || subjectText || "email");

        // ── Server-side filename duplicate check ────────────────────────────
        // Runs after baseName is known but before any upload.
        // `duplicateFilingWarning` being true means the user already confirmed
        // via the "Continue Filing" button — skip and proceed.
        if (settings.duplicates !== "off" && !duplicateFilingWarning && !allowRefilingOverride) {
          let serverDupName = "";
          try {
            if (filingMode !== "attachments") {
              if (await checkDuplicateFilename(caseId, `${baseName}.eml`)) {
                serverDupName = `${baseName}.eml`;
              }
            }
            if (!serverDupName) {
              for (const attId of selectedAttachments) {
                const meta = attachmentsLite.find((a) => String(a.id) === String(attId));
                if (meta?.name && await checkDuplicateFilename(caseId, meta.name)) {
                  serverDupName = meta.name;
                  break;
                }
              }
            }
          } catch (e) {
            console.warn("[doSubmit] Server duplicate check failed, proceeding:", e);
          }

          if (serverDupName) {
            console.log("[doSubmit] Server duplicate found:", serverDupName, settings.duplicates);
            setIsSubmitting(false);
            setViewMode(composeMode ? "prompt" : "pickCase");
            if (settings.duplicates === "block") {
              setSubmitError(
                `A document named "${serverDupName}" already exists in this case. Filing has been prevented.`
              );
            } else {
              // warn — show confirm/cancel UI
              setDuplicateFilingWarning(true);
              setSubmitError(
                `A document named "${serverDupName}" already exists in this case. Do you want to upload it again?`
              );
            }
            return;
          }
        }
        // ─────────────────────────────────────────────────────────────────────

        const emailText =
          `From: ${fromName} <${fromEmail}>\r\n` +
          `To: SingleCase <noreply@singlecase>\r\n` +
          `Subject: ${subjectText}\r\n` +
          `Date: ${new Date().toUTCString()}\r\n` +
          `Message-ID: <${activeItemId}@outlook>\r\n` +
          `MIME-Version: 1.0\r\n` +
          `Content-Type: text/plain; charset=UTF-8\r\n` +
          `Content-Transfer-Encoding: 8bit\r\n` +
          `\r\n` +
          `${bodyForEml}\r\n`;

        const emailBase64 = toBase64Utf8(emailText);

        const originalCaseId = String(sentPill?.caseId || replyBaseCaseId || "");
        const existingEmailDocId = String(sentPill?.documentId || replyBaseEmailDocId || "");
        console.log("[doSubmit] new version check", {
          isUploadingNewVersion,
          existingEmailDocId,
          originalCaseId,
          caseId,
          replyBaseCaseId,
          replyBaseEmailDocId,
          sentPillCaseId: sentPill?.caseId,
          sentPillDocId: sentPill?.documentId,
        });

        const uploadAsNewVersion =
          isUploadingNewVersion &&
          Boolean(existingEmailDocId) &&
          Boolean(originalCaseId) &&
          caseId === originalCaseId;

        let singlecaseRecordId: any = sentPill?.singlecaseRecordId;

        if (!uploadAsNewVersion) {
          const payload: any = {
            caseId,
            outlookItemId:
              activeItemId && !String(activeItemId).startsWith("draft:") ? activeItemId : undefined,
            subject: subjectText,
            fromEmail,
            fromName,
          };

          if (bodySnippetFull) payload.bodySnippet = bodySnippetFull;

          const res = await submitEmailToCase(token, payload);
          singlecaseRecordId = res?.singlecaseRecordId;
        }

        let emailDocId = "";
        let emailRevisionNumber = sentPill?.revisionNumber ?? 1;

        console.log("[doSubmit] About to upload email document", {
          caseId,
          filingMode,
          uploadAsNewVersion,
          workspaceHost,
          hasToken: !!token,
        });

        if (filingMode === "both") {
          if (uploadAsNewVersion) {
            console.log("[doSubmit] Uploading new version of existing document", {
              existingEmailDocId,
            });
            const v = await uploadDocumentVersion({
              documentId: existingEmailDocId,
              fileName: `${baseName}.eml`,
              mimeType: "message/rfc822",
              dataBase64: emailBase64,
            });

            let rev = extractRevisionFromVersionUploadResponse(v);

            if (!rev) {
              for (let i = 0; i < 3; i += 1) {
                const metaRaw = await getDocumentMeta(existingEmailDocId);
                const meta = extractDocMeta(metaRaw);
                rev = extractLatestRevisionNumber(meta);
                if (rev) break;
                await new Promise((r) => setTimeout(r, 300));
              }
            }

            if (!rev) {
              throw new Error(
                `Version upload succeeded but revision could not be determined. docId=${existingEmailDocId}`
              );
            }

            emailDocId = existingEmailDocId;
            emailRevisionNumber = rev;

            const url = buildDocumentUrl(workspaceHost, emailDocId);
            if (url) {
              const item: UploadedItem = {
                id: emailDocId,
                name: `${baseName}-v${emailRevisionNumber}.eml`,
                url,
                kind: "email",
                atIso: new Date().toISOString(),
                uploadedBy: userLabel,
                caseId,
              };

              const existing = await loadUploadedLinks(storeKey);
              const merged = [
                item,
                ...existing.filter((x: any) => String(x?.id) !== String(item.id)),
              ].slice(0, 25);
              await saveUploadedLinks(storeKey, merged as any);
              setUploadedLinksRaw(merged as any);
              setUploadedLinksValidated(merged as any);
            }
          } else {
            // Get conversationId for metadata (needed for server-authoritative filing detection)
            const item = Office?.context?.mailbox?.item as any;
            const conversationId = String(item?.conversationId || "").trim();

            console.log("[doSubmit] Creating new email document", {
              caseId,
              fileName: `${baseName}.eml`,
              mimeType: "message/rfc822",
              conversationId: conversationId ? conversationId.substring(0, 30) + "..." : "MISSING",
            });
            const createdEmail = await uploadDocumentToCase({
              caseId,
              fileName: `${baseName}.eml`,
              mimeType: "message/rfc822",
              dataBase64: emailBase64,
              metadata: {
                subject: subjectText,
                fromEmail,
                fromName,
                conversationId: conversationId || undefined,
              },
            });

            console.log("[doSubmit] Email document created", { response: createdEmail });
            const newDoc: any = createdEmail?.documents?.[0] || null;
            emailDocId = newDoc?.id ? String(newDoc.id) : "";
            emailRevisionNumber = newDoc?.latest_version?.revision_number ?? 1;
            console.log("[doSubmit] Extracted document ID", { emailDocId, emailRevisionNumber });

            const url = emailDocId ? buildDocumentUrl(workspaceHost, emailDocId) : "";
            if (emailDocId && url) {
              const item: UploadedItem = {
                id: emailDocId,
                name:
                  emailRevisionNumber > 1
                    ? `${baseName}-v${emailRevisionNumber}.eml`
                    : `${baseName}.eml`,
                url,
                kind: "email",
                atIso: new Date().toISOString(),
                uploadedBy: userLabel,
                caseId,
              };

              const existing = await loadUploadedLinks(storeKey);
              const merged = [
                item,
                ...existing.filter((x: any) => String(x?.id) !== String(item.id)),
              ].slice(0, 25);
              await saveUploadedLinks(storeKey, merged as any);
              setUploadedLinksRaw(merged as any);
              setUploadedLinksValidated(merged as any);
            }
          }
        }

        if (filingMode === "attachments" || filingMode === "both") {
          for (const attId of selectedAttachments) {
            const meta = attachmentsLite.find((a) => String(a.id) === String(attId));
            const content = await getAttachmentBase64(attId, meta?.name);

            const createdAtt = await uploadDocumentToCase({
              caseId,
              fileName: content.name,
              mimeType: content.mime,
              dataBase64: content.base64,
            });

            const attDoc: any = createdAtt?.documents?.[0] || null;
            const attDocId = attDoc?.id ? String(attDoc.id) : "";
            if (!attDocId) continue;

            const attUrl = buildDocumentUrl(workspaceHost, attDocId);
            if (!attUrl) continue;

            const newItem: UploadedItem = {
              id: attDocId,
              name: content.name,
              url: attUrl,
              kind: "attachment",
              atIso: new Date().toISOString(),
              uploadedBy: userLabel,
              caseId,
            };

            const existing = await loadUploadedLinks(storeKey);
            const merged = [
              newItem,
              ...existing.filter((x: any) => String(x?.id) !== String(newItem.id)),
            ].slice(0, 25);
            await saveUploadedLinks(storeKey, merged as any);
            setUploadedLinksRaw(merged as any);
            setUploadedLinksValidated(merged as any);
          }
        }

        const pill: SentPillData = {
          sent: true,
          atIso: new Date().toISOString(),
          caseId,
          singlecaseRecordId: singlecaseRecordId,
          documentId: emailDocId || (uploadAsNewVersion ? existingEmailDocId : undefined),
          revisionNumber: emailRevisionNumber,
          filedBy: userLabel as any,
        };

        await saveSentPill(storeKey, pill);
        // Also save under conv: key so State B can recover even if the Exchange
        // item ID changes after the SC:Filed category is applied.
        if (conversationKey) {
          try { await saveSentPill(`conv:${conversationKey}`, pill); } catch { /* ignore */ }
        }
        setSentPill(pill);
        markEverFiled(storeKey);
        try {
          if (conversationKey) {
            await saveConversationFiledCase(conversationKey, caseId);
          }
        } catch {
          // ignore
        }

        try {
          if (conversationKey && emailDocId) {
            await saveConversationFiledCtx(conversationKey, { caseId, emailDocId });
          }
        } catch {
          // ignore
        }
        try {
          const recips = composeMode ? composeRecipientsLive : [];
          if (recips.length > 0) await recordRecipientsFiledToCase(recips, caseId);
        } catch {
          // ignore
        }

        try {
          await applyFiledCategoryToCurrentEmailOfficeJs();
        } catch (err) {
          console.warn("Office category apply failed:", err);
        }
        setForceUnfiledLabel(false);

        // Cache filed email for "already filed" detection
        try {
          const { cacheFiledEmail, cacheFiledEmailBySubject } = await import("../../../utils/filedCache");
          const item = Office?.context?.mailbox?.item as any;
          const conversationId = String(item?.conversationId || "").trim();

          const selectedCase = cases.find((c) => String(c.id) === String(caseId));
          const caseTitle = selectedCase?.title || "";

          // Parse caseName and caseKey from title (format: "KEY · Name" or just "Name")
          const parts = caseTitle.split(" · ");
          const caseName = parts.length > 1 ? parts[1] : caseTitle;
          const caseKey = parts.length > 1 ? parts[0] : undefined;

          if (conversationId && emailDocId) {
            await cacheFiledEmail(conversationId, caseId, emailDocId, subjectText, caseName, caseKey);
            console.log("[doSubmit] Cached filed email by conversationId");
          } else if (subjectText && emailDocId) {
            await cacheFiledEmailBySubject(subjectText, caseId, emailDocId, caseName, caseKey);
            console.log("[doSubmit] Cached filed email by subject (fallback)");
          }
        } catch (err) {
          console.warn("[doSubmit] Failed to cache filed email:", err);
        }

        // Mark in DuplicateCache so future filings of this email (or an identical
        // one) can be caught regardless of itemId / conversationId.
        try {
          const freshDupCache = loadDuplicateCache();
          if (dupFingerprint) markAttached(freshDupCache, caseId, dupFingerprint);
          if (activeItemId && !activeItemId.startsWith("draft:")) {
            markAttached(freshDupCache, caseId, activeItemId);
          }
          saveDuplicateCache(freshDupCache);
          console.log("[doSubmit] Marked in duplicate cache", { fingerprint: dupFingerprint });
        } catch (err) {
          console.warn("[doSubmit] Failed to update duplicate cache:", err);
        }

        // Record in suggestion history for sender-based future suggestions
        try {
          recordSuccessfulAttach({
            caseId,
            conversationKey: conversationKey || undefined,
            senderEmail: composeMode ? undefined : fromEmail || undefined,
          });
          console.log("[doSubmit] Recorded suggest history", { caseId, senderEmail: composeMode ? undefined : fromEmail });
        } catch (err) {
          console.warn("[doSubmit] Failed to record suggest history:", err);
        }

        if (composeMode) {
          setPrompt({
            itemId: storeKey,
            kind: "filed",
            text: "Filed into SingleCase. You can now send the email normally in Outlook.",
          });
        } else {
          setPrompt({ itemId: storeKey, kind: "filed", text: "This email is already filed." });
        }

        setViewMode("sent");
        setIsUploadingNewVersion(false);
      } catch (e) {
        console.error("[doSubmit] Upload failed", e);
        console.error("[doSubmit] Error details:", {
          name: e instanceof Error ? e.name : "unknown",
          message: e instanceof Error ? e.message : String(e),
          stack: e instanceof Error ? e.stack : undefined,
        });

        // Properly extract error message
        let msg = "Unknown error";
        if (e instanceof Error) {
          msg = e.message;
        } else if (typeof e === "string") {
          msg = e;
        } else if (e && typeof e === "object") {
          // Try to extract message from object
          msg = (e as any).message || (e as any).error || JSON.stringify(e);
        }

        setSubmitError(`Error while saving: ${msg}`);
        setViewMode(composeMode ? "prompt" : "pickCase");
        setTimeout(() => {
          try {
            chatBodyRef.current?.scrollTo({ top: 0, behavior: "smooth" });
          } catch {
            // ignore
          }
        }, 50);
      } finally {
        setIsSubmitting(false);
      }
    },
    [
      isSubmitting,
      selectedCaseId,
      pickStep,
      attachmentsLite.length,
      selectedAttachments,
      subjectText,
      fromEmail,
      fromName,
      token,
      suggestBodySnippet,
      workspaceHost,
      filingMode,
      isUploadingNewVersion,
      replyBaseCaseId,
      replyBaseEmailDocId,
      sentPill?.caseId,
      sentPill?.documentId,
      sentPill?.singlecaseRecordId,
      sentPill?.revisionNumber,
      userLabel,
      composeMode,
      allowRefilingOverride,
      composeRecipientsLive,
      storeKey,
      activeItemId,
      settings.duplicates,
      settings.internalEmailHandling,
      duplicateFilingWarning,
    ]
  );

  const doSubmitOverrideRef = React.useRef<((override: { caseId: string }) => Promise<void>) | null>(null);

  React.useEffect(() => {
    doSubmitRef.current = () => doSubmit();
    doSubmitOverrideRef.current = (override) => doSubmit(override);
    return () => {
      doSubmitRef.current = null;
      doSubmitOverrideRef.current = null;
    };
  }, [doSubmit]);

  return (
    <div className="mwPage">
      <div className="mwHero">
        <h1 className="mwGreeting">
          {greeting}, {userLabel}
        </h1>
        <p className="mwQuestion">What can I do for you?</p>
      </div>

      {submitError ? (
        <div className="mwChatBubble mwChatBubbleError">
          {submitError}
        </div>
      ) : null}
      {duplicateFilingWarning ? (
        <div className="mwQuickActions">
          <button
            type="button"
            className="mwQuickAction"
            onClick={() => void doSubmit()}
          >
            Continue filing
          </button>
          <button
            type="button"
            className="mwQuickAction"
            onClick={() => {
              setDuplicateFilingWarning(false);
              setSubmitError("");
            }}
          >
            Cancel
          </button>
        </div>
      ) : null}
      {isItemLoading ? (
        <div className="mwFiledSummaryCard">
          <div className="mwChatBubble">
            <span className="mwThinking">
              Loading email
              <span className="mwDot mwDotBounce1" />
              <span className="mwDot mwDotBounce2" />
              <span className="mwDot mwDotBounce3" />
            </span>
          </div>
        </div>
      ) : showFiledSummary ? (
     <FiledSummaryCard
  caseUrl={caseUrl}
  filedCaseName={filedCaseName}
  sentPill={sentPill}
  documentsToShow={documentsToShow}
  workspaceHost={workspaceHost}
  onOpenUrl={openUrl}
  buildLiveEditUrl={buildLiveEditUrl}
  fmtCs={fmtCs}
  onLockedDocAttempt={(msg) => setLockedDocAlert(msg)}
/>
      ) : null}

      <div className="mwChatCard">
        <div className="mwChatBody" ref={chatBodyRef}>
          {isItemLoading ? (
            <div className="mwChatBubble">
              <span className="mwThinking">
                Loading
                <span className="mwDot mwDotBounce1" />
                <span className="mwDot mwDotBounce2" />
                <span className="mwDot mwDotBounce3" />
              </span>
            </div>
          ) : null}

          {pendingFiling ? (
            <div className="mwPendingFilingBanner">
              <div className="mwPendingFilingText">
                Email sent. File it to SingleCase now?
                {pendingFiling.subject ? ` "${pendingFiling.subject}"` : ""}
              </div>
              <div className="mwPendingFilingHelper">
                Nothing will be filed unless you choose File now.
              </div>
              <div className="mwPendingFilingActions">
                <button
                  className="mwPendingFilingBtn mwPendingFilingBtn--primary"
                  onClick={() => handleQuickAction("confirm_file_now")}
                >
                  File now
                </button>
                <button
                  className="mwPendingFilingBtn mwPendingFilingBtn--secondary"
                  onClick={() => handleQuickAction("skip_pending_filing")}
                >
                  Skip
                </button>
              </div>
            </div>
          ) : null}

          {viewMode === "prompt" && (prompt.kind !== "filed" || !showFiledSummary) ? (
            <PromptBubble
              text={prompt.text}
              isUnfiled={prompt.kind === "unfiled" || prompt.kind === "deleted"}
              tone={
                suppressInternalSuggestions && Boolean(selectedCaseId) ? "success" :
                composeMode && settings.filingOnSend === "warn" ? "warning" :
                composeMode && settings.filingOnSend === "always" && Boolean(selectedCaseId) ? "success" :
                "default"
              }
             actions={(quickActions || []).map((a) => ({
    id: a.id,
    label: a.label,
    onClick: () => {
      console.log("[UI] clicked", a.intent, { viewMode, chatStep });
      handleQuickAction(a.intent);
    },
  }))}
            />
          ) : null}

          {viewMode === "pickCase" ? (
            <div>
              <CaseSelector
                title="Case"
                scope={caseGroupTab}
                onScopeChange={(scope) => setCaseGroupTab(scope === "favourites" ? "favourites" : "all")}
                selectedCaseId={selectedCaseId}
                onSelectCaseId={(id) => {
                  setSelectedCaseId(id);
                  setSelectedSource("manual");
                  // Manual pick supersedes any system suggestion and resets the filing target
                  setSuggestedCaseId("");
                  setSuggestedConfidencePct(0);
                  setFilingTargetCaseId("");
                  setReplyBaseCaseId("");
                  setReplyBaseEmailDocId("");
                  setIsUploadingNewVersion(false);
                  if (settings.rememberLastCase) saveLastCaseId(id);

                  if (composeMode) {
                    const c: any = (cases || []).find((x: any) => String(x?.id) === String(id));
                    const name =
                      String(c?.name || c?.title || c?.label || "").trim() || `Case ${id}`;

                    setViewMode("prompt");
                    setPickStep("case");
                    setQuickActions([]);
                    setChatStep("compose_ready");

                    if (storeKey)
                      void saveComposeIntent({ itemKey: storeKey, caseId: id, autoFileOnSend });

                    setPrompt({
                      itemId: storeKey || activeItemId,
                      kind: "unfiled",
                      text: settings.filingOnSend === "always"
                        ? `Auto filing is on. When you hit Send, I'll file this email to SingleCase automatically. Case: ${name}.`
                        : `Heads up — when you hit Send, I'll let the email go through as normal. Prepared case: ${name}.`,
                    });
                    return;
                  }

                  // Read mode + internal email flow → show confirmation banner
                  if (suppressInternalSuggestions) {
                    const c: any = (cases || []).find((x: any) => String(x?.id) === String(id));
                    const caseName = String(c?.title || c?.name || c?.label || "").trim() || `Case ${id}`;
                    setViewMode("prompt");
                    setPickStep("case");
                    setQuickActions([
                      { id: "fm_file_now", label: "File now", intent: "file_now" },
                      { id: "fm2", label: "Select a different case", intent: "file_manually" },
                    ]);
                    setPrompt({
                      itemId: storeKey || activeItemId,
                      kind: "unfiled",
                      text: `Case selected: ${caseName}. Ready to file this email.`,
                    });
                    return;
                  }

                  if (attachmentsLite.length > 0) setPickStep("attachments");
                }}
                suggestedCaseId={suggestedCaseId}
                suggestions={caseSuggestions}
                cases={visibleCases}
                isLoadingCases={isLoadingCases}
                clientNamesById={clientNamesById}
                contentSuggestions={contentBasedSuggestions}
                isLoadingContentSuggestions={isLoadingContentSuggestions}
              />

              {pickStep === "attachments" && attachmentsLite.length > 0 ? (
                <AttachmentsStep
                  attachmentsLite={attachmentsLite}
                  attachmentIds={attachmentIds}
                  selectedAttachments={selectedAttachments}
                  onSelectionChange={setSelectedAttachments}
                  filingMode={filingMode}
                  onFilingModeChange={setFilingMode}
                  containerRef={attachmentsRef}
                />
              ) : null}
            </div>
          ) : null}

          {viewMode === "sending" ? (
            <div className="mwChatBubble">
              <span className="mwThinking">
                Sending
                <span className="mwDot mwDotBounce1" />
                <span className="mwDot mwDotBounce2" />
                <span className="mwDot mwDotBounce3" />
              </span>
            </div>
          ) : null}

          {viewMode === "sent" ? (
            <>
              {showFiledSummary && !lockedDocAlert ? (
                <div className="mwChatBubble">Documents filed successfully</div>
              ) : (
                <PromptBubble
                  text={prompt.text || "This email is already filed."}
                  isUnfiled={false}
                  tone="default"
                  actions={[]}
                />
              )}
              <div ref={chatEndRef} />
            </>
          ) : null}

          {lockedDocAlert ? (
            <div className="mwChatBubbleWarning">
              ⚠ {lockedDocAlert}
            </div>
          ) : null}

          <div ref={chatEndRef} />
        </div>

        {/* RECEIVED MODE ACTIONS (hidden when dismissed — dismissal shows its own inline actions) */}
        {viewMode === "prompt" && (prompt.kind === "unfiled" || prompt.kind === "deleted") && !composeMode && !dismissedRef.current.has(prompt.itemId) ? (
          <div className="mwActionsBar">
            <>
                {/* No / Yes, file it */}
                <button
                  className="mwGhostBtn"
                  type="button"
                  onClick={() => {
                    dismissedRef.current.add(prompt.itemId);
                    // Exchange changes the item ID after the category is applied,
                    // so also key by conversation so the dismissal survives the ID change.
                    if (conversationKey) {
                      dismissedRef.current.add(`conv:${conversationKey}`);
                    }

                    // Apply SC: Unfiled category.
                    // Office.js: updates Outlook UI immediately (best-effort).
                    // Graph: persists to Exchange — reliable on Mac Desktop and OWA.
                    void (async () => {
                      try {
                        await applyUnfiledCategoryToCurrentEmailOfficeJs();
                      } catch (e) {
                        console.warn("[No] Office.js applyUnfiledCategory failed:", e);
                      }
                      try {
                        await applyUnfiledCategoryToCurrentEmailGraph();
                      } catch (e) {
                        console.warn("[No] Graph applyUnfiledCategory failed:", e);
                      }
                      setForceUnfiledLabel(true);
                    })();

                    // Show dismissal message with manual filing option
                    const dismissMsg = "Got it. I'll step back for this email, but you can still file it later.";

                    setQuickActions([
                      { id: "fm1", label: "File manually", intent: "file_manually" },
                    ]);
                    setPrompt({
                      itemId: prompt.itemId,
                      kind: "unfiled",
                      text: dismissMsg,
                    });
                  }}
                >
                  No
                </button>

                <button
                  className="mwPrimaryBtn"
                  type="button"
                  onClick={() => {
                    console.log("[UI] Yes, file it - setting refiling override");
                    setAllowRefilingOverride(true); // Allow bypassing duplicate guard
                    setViewMode("pickCase");
                    setPickStep("case");
                    setSelectedCaseId("");
                    setSelectedSource("");
                    setSelectedAttachments([]);
                    setIsUploadingNewVersion(false);
                  }}
                >
                  Yes, file it
                </button>
            </>
          </div>
        ) : null}

        {/* RECEIVED MODE PICKCASE ACTIONS */}
        {viewMode === "pickCase" && !composeMode ? (
          <div className="mwActionsBar">
            {pickStep === "attachments" ? (
              <button
                className="mwGhostBtn"
                type="button"
                onClick={() => {
                  setPickStep("case");
                  setSelectedAttachments([]);
                }}
              >
                Back
              </button>
            ) : (
              <button
                className="mwGhostBtn"
                type="button"
                onClick={() => {
                  setSubmitError("");

                  if (isUploadingNewVersion) {
                    setIsUploadingNewVersion(false);
                    setViewMode("sent");
                    setPickStep("case");
                    setPrompt({
                      itemId: storeKey,
                      kind: "filed",
                      text: "Tento email je již zařazen.",
                    });
                    return;
                  }

                  // Clear dismiss so suggestions re-appear
                  dismissedRef.current.delete(storeKey);

                  setViewMode("prompt");
                  setPickStep("case");
                  setSelectedCaseId("");
                  setSelectedSource("");
                  setSelectedAttachments([]);
                  setPrompt({
                    itemId: storeKey,
                    kind: "unfiled",
                    text: "This email isn't filed yet. Would you like me to file it to a case?",
                  });
                }}
              >
                Close
              </button>
            )}

            {pickStep === "case" ? (
              (() => {
                // Continue is enabled when the user manually selected a case OR
                // when rememberLastCase is OFF and there is a high-confidence suggestion.
                const canContinue =
                  !!selectedCaseId ||
                  (!settings.rememberLastCase && !!suggestedCaseId && suggestedConfidencePct >= 70);
                return (
                  <button
                    className={`mwPrimaryBtn ${!canContinue ? "mwPrimaryBtnDisabled" : ""}`}
                    type="button"
                    disabled={!canContinue}
                    onClick={() => {
                      // Resolve the case to file: explicit selection wins; fall back to suggestion.
                      const target =
                        selectedCaseId ||
                        (suggestedConfidencePct >= 70 ? suggestedCaseId : "");
                      if (!target) return;
                      setFilingTargetCaseId(target);
                      if (attachmentsLite.length > 0) setPickStep("attachments");
                      else void doSubmit({ caseId: target });
                    }}
                  >
                    Continue
                  </button>
                );
              })()
            ) : (
              (() => {
                // Attachments step: use the filing target set by the previous Continue click.
                const effectiveCaseId = filingTargetCaseId || selectedCaseId;
                const attachmentsRequired = attachmentsLite.length > 0 && selectedAttachments.length === 0;
                return (
                  <button
                    className={`mwPrimaryBtn ${
                      !effectiveCaseId || isSubmitting || attachmentsRequired
                        ? "mwPrimaryBtnDisabled"
                        : ""
                    }`}
                    type="button"
                    disabled={!effectiveCaseId || isSubmitting || isItemLoading || attachmentsRequired}
                    onClick={() => void doSubmit({ caseId: effectiveCaseId })}
                  >
                    Continue
                  </button>
                );
              })()
            )}
          </div>
        ) : null}

        {viewMode === "sent" && !composeMode ? (
          <div className="mwActionsBar">
            <button
              className="mwGhostBtn"
              type="button"
              onClick={() => {
  console.log("[toggleCategory] 🔵 Button clicked!", {
    forceUnfiledLabel,
    viewMode,
    composeMode,
    activeItemId,
  });

  void (async () => {
    const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

    const removeCats = async (names: string[]) =>
      new Promise<void>((resolve, reject) => {
        try {
          const item = Office?.context?.mailbox?.item as any;
          const cats = item?.categories;
          if (!cats?.removeAsync) {
            console.log("[toggleCategory] ⚠️ removeAsync not available");
            return resolve();
          }
          console.log("[toggleCategory] Removing categories:", names);
          cats.removeAsync(names, (res: any) => {
            if (res?.status === Office.AsyncResultStatus.Failed) {
              console.error("[toggleCategory] ❌ removeAsync failed:", res?.error);
              reject(res?.error);
            } else {
              console.log("[toggleCategory] ✅ removeAsync succeeded");
              resolve();
            }
          });
        } catch (e) {
          console.error("[toggleCategory] ❌ removeAsync exception:", e);
          reject(e);
        }
      });

    try {
      setIsTogglingCategory(true);

      console.log("[toggleCategory] 📊 Starting toggle process", {
        forceUnfiledLabel,
        filed: FILED_CATEGORY,
        unfiled: UNFILED_CATEGORY,
      });

      // Read current state from Office (not Graph) as source of truth
      const beforeOffice = await getOfficeMailCategoriesNorm();
      const hasUnfiledBefore = beforeOffice.includes(unfiledCatNorm);
      const hasFiledBefore = beforeOffice.includes(filedCatNorm);

      const currentIsUnfiled = hasUnfiledBefore && !hasFiledBefore
        ? true
        : hasFiledBefore && !hasUnfiledBefore
          ? false
          : Boolean(forceUnfiledLabel); // fallback

      const targetIsUnfiled = !currentIsUnfiled;

      console.log("[toggleCategory] beforeOffice", {
        beforeOffice,
        hasUnfiledBefore,
        hasFiledBefore,
        currentIsUnfiled,
        targetIsUnfiled,
      });

      // Optimistic UI
      setForceUnfiledLabel(targetIsUnfiled);

      // Enforce exclusivity first
      await removeCats([FILED_CATEGORY, UNFILED_CATEGORY]);
      await sleep(150);

      // Apply target label
      console.log("[toggleCategory] 🎯 Applying target category:", targetIsUnfiled ? "UNFILED" : "FILED");
      if (targetIsUnfiled) {
        await applyUnfiledCategoryToCurrentEmailOfficeJs();
        console.log("[toggleCategory] ✅ Applied UNFILED category");
      } else {
        await applyFiledCategoryToCurrentEmailOfficeJs();
        console.log("[toggleCategory] ✅ Applied FILED category");
      }

      // Verify with Office reads (retry)
      for (let i = 0; i < 12; i += 1) {
        await sleep(250);

        const afterOffice = await getOfficeMailCategoriesNorm();
        const hasUnfiled = afterOffice.includes(unfiledCatNorm);
        const hasFiled = afterOffice.includes(filedCatNorm);

        console.log("[toggleCategory] verify", { i, afterOffice, hasUnfiled, hasFiled });

        // If both exist, clean and reapply once
        if (hasUnfiled && hasFiled) {
          await removeCats([FILED_CATEGORY, UNFILED_CATEGORY]);
          await sleep(150);
          if (targetIsUnfiled) await applyUnfiledCategoryToCurrentEmailOfficeJs();
          else await applyFiledCategoryToCurrentEmailOfficeJs();
          continue;
        }

        if (targetIsUnfiled && hasUnfiled && !hasFiled) {
          setForceUnfiledLabel(true);
          return;
        }

        if (!targetIsUnfiled && hasFiled && !hasUnfiled) {
          setForceUnfiledLabel(false);
          return;
        }
      }

      console.warn("[toggleCategory] ⚠️ Could not observe category change after retries, keeping optimistic UI");
    } catch (e) {
      console.error("[toggleCategory] ❌ FAILED:", e);
      console.error("[toggleCategory] Error details:", {
        message: e instanceof Error ? e.message : String(e),
        stack: e instanceof Error ? e.stack : undefined,
      });

      // revert by reading current Outlook state
      try {
        console.log("[toggleCategory] 🔄 Attempting to revert state from Outlook");
        await syncForceUnfiledFromOutlook(filedCatNorm, unfiledCatNorm, setForceUnfiledLabel);
        console.log("[toggleCategory] ✅ State reverted successfully");
      } catch (revertError) {
        console.error("[toggleCategory] ❌ Could not revert state:", revertError);
      }
    } finally {
      setIsTogglingCategory(false);
      console.log("[toggleCategory] 🏁 Toggle process completed");
    }
  })();
}}
            >
              {forceUnfiledLabel ? "Mark as Filed" : "Mark as Unfiled"}
            </button>
          </div>
        ) : null}

        {/* Filed-state quick actions (e.g. "Mark as SC: Unfiled").
            Rendered inside mwChatCard so it sits in the bottom actions strip,
            not above the FiledSummaryCard. */}
        {viewMode === "prompt" && prompt.kind === "filed" && showFiledSummary && !composeMode && !isItemLoading && (quickActions || []).length > 0 ? (
          <div className="mwActionsBar">
            {(quickActions || []).map((a) => (
              <button
                key={a.id}
                className="mwGhostBtn"
                type="button"
                onClick={() => handleQuickAction(a.intent)}
              >
                {a.label}
              </button>
            ))}
          </div>
        ) : null}
      </div>

      {renameOpen ? (
        <div
          className="mwRenameOverlay"
          onClick={() => {
            if (!renameSaving) {
              setRenameOpen(false);
              setRenameDoc(null);
            }
          }}
        >
          <div className="mwRenameDialog" onClick={(e) => e.stopPropagation()}>
            <div className="mwRenameTitle">Přejmenovat dokument</div>

            <input
              type="text"
              value={renameValue}
              onChange={(e) => setRenameValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !renameSaving && String(renameValue || "").trim()) {
                  void confirmRename();
                }
              }}
              className="mwRenameInput"
              aria-label="Název dokumentu"
              autoFocus
            />

            <div className="mwRenameActions">
              <button
                type="button"
                className="mwGhostBtn"
                disabled={renameSaving}
                onClick={() => closeRename()}
              >
                Zrušit
              </button>

              <button
                type="button"
                className="mwPrimaryBtn"
                disabled={renameSaving || !String(renameValue || "").trim()}
                onClick={() => void confirmRename()}
              >
                {renameSaving ? "Ukládám" : "Uložit"}
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
