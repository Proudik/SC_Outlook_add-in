/* global Office, OfficeRuntime */

import { getAuthRuntime, clearAuthIfExpiredRuntime } from "../services/auth";
import { getStored, setStored } from "../utils/storage";
import { STORAGE_KEYS } from "../utils/constants";
import { uploadDocumentToCase, uploadDocumentVersion, findDocumentBySubject } from "../services/singlecaseDocuments";
import { cacheFiledEmail, cacheFiledEmailBySubject } from "../utils/filedCache";
import { recordRecipientsFiledToCase } from "../utils/recipientHistory";

type SendEvent = Office.AddinCommands.Event;

const T_ITEMKEY_MS = 2000;
const T_STORAGE_MS = 1500;
const T_FETCH_MS = 10000;
const T_SUBJECT_MS = 1500;
const T_BODY_MS = 2500;

const CONV_CTX_KEY_PREFIX = "sc_conv_ctx:";
const LAST_FILED_CTX_KEY = "sc_last_filed_ctx";

function withTimeout<T>(p: Promise<T>, ms: number): Promise<T> {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error("timeout")), ms);
    p.then(
      (v) => {
        clearTimeout(t);
        resolve(v);
      },
      (e) => {
        clearTimeout(t);
        reject(e);
      }
    );
  });
}

function normalizeHost(host: string): string {
  const v = (host || "").trim().toLowerCase();
  if (!v) return "";
  return v.replace(/^https?:\/\//i, "").split("/")[0];
}

function safeFileName(value: string): string {
  const v = (value || "").trim();
  const cleaned = v
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned.slice(0, 80) || "email";
}

function toBase64Utf8(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

function getConversationIdSafe(): string {
  try {
    const item = Office.context.mailbox.item as any;
    return String(item?.conversationId || item?.conversationKey || "").trim();
  } catch {
    return "";
  }
}

async function persistFiledCtx(caseId: string, emailDocId: string) {
  const cid = String(caseId || "").trim();
  const did = String(emailDocId || "").trim();
  if (!cid || !did) return;

  const payload = JSON.stringify({ caseId: cid, emailDocId: did });

  // Always keep a last-known base
  try {
    await setStored(LAST_FILED_CTX_KEY, payload);
  } catch {
    // ignore
  }

  // And also map per conversation where possible
  try {
    const convId = getConversationIdSafe();
    if (convId) {
      await setStored(`${CONV_CTX_KEY_PREFIX}${convId}`, payload);
    }
  } catch {
    // ignore
  }
}

async function getCandidateItemKeysRuntime(): Promise<string[]> {
  const item = Office.context.mailbox.item as any;
  if (!item) {
    console.warn("[getCandidateItemKeysRuntime] No item available");
    return [];
  }

  console.log("[getCandidateItemKeysRuntime] Item properties:", {
    hasItemId: !!item.itemId,
    itemId: String(item.itemId || "").substring(0, 20),
    hasConversationId: !!item.conversationId,
    conversationId: String(item.conversationId || "").substring(0, 20),
    hasConversationKey: !!item.conversationKey,
    hasDateTimeCreated: !!item.dateTimeCreated,
    hasGetItemIdAsync: typeof item.getItemIdAsync === "function",
    itemType: item.itemType,
  });

  const keys: string[] = [];

  const direct = String(item.itemId || "").trim();
  if (direct) keys.push(direct);

  if (typeof item.getItemIdAsync === "function") {
    try {
      const asyncId: string = await new Promise((resolve) => {
        item.getItemIdAsync((res: any) => {
          if (res?.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || ""));
          else resolve("");
        });
      });
      if (asyncId) {
        console.log(
          "[getCandidateItemKeysRuntime] getItemIdAsync returned:",
          asyncId.substring(0, 20)
        );
        keys.push(asyncId);
      }
    } catch (e) {
      console.warn("[getCandidateItemKeysRuntime] getItemIdAsync failed:", e);
    }
  }

  const conv = String(item.conversationId || item.conversationKey || "").trim();
  if (conv) keys.push(`draft:${conv}`);

  const created = String(item.dateTimeCreated || "").trim();
  if (created) keys.push(`draft:${created}`);

  // Always include fallback keys for new compose emails
  keys.push("draft:current");
  keys.push("last_compose");

  console.log("[getCandidateItemKeysRuntime] Generated keys:", keys);

  return Array.from(new Set(keys.filter(Boolean)));
}

async function readIntentAny(
  itemKeys: string[]
): Promise<{
  itemKey: string;
  caseId: string;
  autoFileOnSend: boolean;
  filingOnSend: string;
  baseCaseId?: string;
  baseEmailDocId?: string;
} | null> {
  for (const k of itemKeys) {
    const key = `sc_intent:${k}`;
    console.log("[readIntentAny] Trying key:", key);

    try {
      let raw: string | null = null;

      if (typeof OfficeRuntime !== "undefined" && (OfficeRuntime as any)?.storage) {
        try {
          raw = await (OfficeRuntime as any).storage.getItem(key);
          if (raw) console.log("[readIntentAny] Found in OfficeRuntime.storage");
        } catch (e) {
          console.warn("[readIntentAny] OfficeRuntime.storage.getItem failed:", e);
        }
      }

      if (!raw && Office?.context?.roamingSettings) {
        try {
          raw = Office.context.roamingSettings.get(key);
          if (raw) console.log("[readIntentAny] Found in roamingSettings");
        } catch (e) {
          console.warn("[readIntentAny] roamingSettings.get failed:", e);
        }
      }

      if (!raw) continue;

      const obj = JSON.parse(String(raw));
      const caseId = String(obj?.caseId || "").trim();
      const autoFileOnSend = Boolean(obj?.autoFileOnSend);
      const filingOnSend = String(obj?.filingOnSend || "").trim();
      const baseCaseId = String(obj?.baseCaseId || "").trim();
      const baseEmailDocId = String(obj?.baseEmailDocId || "").trim();

      if (!caseId) continue;

      console.log("[readIntentAny] Intent found:", {
        itemKey: k,
        caseId,
        autoFileOnSend,
        filingOnSend,
        hasBase: !!(baseCaseId && baseEmailDocId),
      });

      return {
        itemKey: k,
        caseId,
        autoFileOnSend,
        filingOnSend,
        baseCaseId: baseCaseId || undefined,
        baseEmailDocId: baseEmailDocId || undefined,
      };
    } catch (e) {
      console.warn("[readIntentAny] Failed to read intent for key:", key, e);
    }
  }

  console.warn("[readIntentAny] No intent found for any key");
  return null;
}

async function getSubjectRuntime(): Promise<string> {
  const item = Office.context.mailbox.item as any;
  if (!item) {
    console.warn("[getSubjectRuntime] No item available");
    return "";
  }

  console.log("[getSubjectRuntime] Item type:", item.itemType, "Mode:", item.itemClass);

  if (typeof item.subject === "string") {
    const subj = String(item.subject || "");
    console.log("[getSubjectRuntime] Direct string subject:", subj);
    return subj;
  }

  if (item?.subject?.getAsync) {
    const v: string = await new Promise((resolve) => {
      item.subject.getAsync((res: any) => {
        if (res?.status === Office.AsyncResultStatus.Succeeded) {
          const subj = String(res.value || "");
          console.log("[getSubjectRuntime] Async subject:", subj);
          resolve(subj);
        } else {
          console.warn("[getSubjectRuntime] getAsync failed:", res?.error);
          resolve("");
        }
      });
    });
    return v || "";
  }

  console.warn("[getSubjectRuntime] No subject API available");
  return "";
}

async function getBodyTextRuntime(): Promise<string> {
  const item = Office.context.mailbox.item as any;
  if (!item?.body?.getAsync) return "";

  const text: string = await new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (res: any) => {
      if (res?.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || ""));
      else resolve("");
    });
  });

  return String(text || "");
}

async function getRecipientsRuntime(): Promise<string[]> {
  const item = Office.context.mailbox.item as any;
  if (!item) return [];

  const readField = (field: any): Promise<string[]> => {
    if (!field) return Promise.resolve([]);
    if (typeof field.getAsync === "function") {
      return new Promise((resolve) => {
        field.getAsync((res: any) => {
          if (res?.status === Office.AsyncResultStatus.Succeeded) {
            resolve(
              (res.value || [])
                .map((r: any) => String(r?.emailAddress || "").toLowerCase().trim())
                .filter(Boolean)
            );
          } else {
            resolve([]);
          }
        });
      });
    }
    if (Array.isArray(field)) {
      return Promise.resolve(
        field.map((r: any) => String(r?.emailAddress || "").toLowerCase().trim()).filter(Boolean)
      );
    }
    return Promise.resolve([]);
  };

  const [to, cc, bcc] = await Promise.all([
    readField(item.to),
    readField(item.cc),
    readField(item.bcc),
  ]);

  return Array.from(new Set([...to, ...cc, ...bcc]));
}

async function showInfo(message: string) {
  try {
    const item: any = Office.context.mailbox.item;
    if (!item?.notificationMessages?.replaceAsync) return;

    await new Promise<void>((resolve) => {
      item.notificationMessages.replaceAsync(
        "sc_send",
        {
          type: "informationalMessage",
          message,
          icon: "Icon.16x16",
          persistent: false,
        },
        () => resolve()
      );
    });
  } catch {
    // ignore
  }
}

export async function onMessageSendHandler(event: SendEvent) {
  console.log("[onMessageSendHandler] Handler fired");
  console.log("[onMessageSendHandler] Platform info", {
    hasOfficeRuntime: typeof OfficeRuntime !== "undefined",
    hasOfficeRuntimeStorage: typeof (OfficeRuntime as any)?.storage !== "undefined",
    hasRoamingSettings: !!Office?.context?.roamingSettings,
    host: Office?.context?.mailbox?.diagnostics?.hostName,
    hostVersion: Office?.context?.mailbox?.diagnostics?.hostVersion,
  });

  let done = false;

  const finish = (allowEvent: boolean, errorMessage?: string) => {
    if (done) return;
    done = true;
    console.log("[onMessageSendHandler] Finishing", { allowEvent, hasErrorMessage: !!errorMessage });
    try {
      if (errorMessage) (event.completed as any)({ allowEvent, errorMessage });
      else event.completed({ allowEvent });
    } catch (e) {
      console.error("[onMessageSendHandler] Error in event.completed:", e);
    }
  };

  try {
    console.log("[onMessageSendHandler] Clearing expired auth");
    await withTimeout(clearAuthIfExpiredRuntime(), 700);

    console.log("[onMessageSendHandler] Getting candidate item keys");
    const keys = await withTimeout(getCandidateItemKeysRuntime(), T_ITEMKEY_MS);
    console.log("[onMessageSendHandler] Item keys:", keys);

    if (keys.length === 0) {
      console.log("[onMessageSendHandler] No item keys found, skipping");
      finish(true);
      return;
    }

    console.log("[onMessageSendHandler] Reading intent from storage", {
      storageType: typeof OfficeRuntime !== "undefined" && (OfficeRuntime as any)?.storage ? "OfficeRuntime" : "roamingSettings",
      keysToTry: keys,
    });
    const intent = await withTimeout(readIntentAny(keys), T_STORAGE_MS);
    console.log("[onMessageSendHandler] Intent:", intent, {
      found: !!intent,
      foundUnderKey: intent?.itemKey,
    });

    if (!intent?.caseId) {
      console.log("[onMessageSendHandler] No case ID in intent, skipping");
      finish(true);
      return;
    }

    // Warn mode (renamed from "ask" in v2): store pending filing for post-send confirmation
    // Accept both "warn" and legacy "ask" values for backward compatibility
    if (intent.filingOnSend === "warn" || intent.filingOnSend === "ask") {
      console.log("[onMessageSendHandler] warn mode — storing pending filing intent without filing");
      try {
        const subjectForPending = await withTimeout(getSubjectRuntime(), T_SUBJECT_MS);
        const convForPending = getConversationIdSafe();
        await setStored("sc_pending_filing", JSON.stringify({
          caseId: intent.caseId,
          subject: subjectForPending,
          conversationId: convForPending,
          sentAt: new Date().toISOString(),
        }));
        console.log("[onMessageSendHandler] Pending filing stored", { caseId: intent.caseId });
      } catch (e) {
        console.warn("[onMessageSendHandler] Failed to store pending filing:", e);
      }
      await showInfo("SingleCase: otevřete panel a potvrďte zařazení.");
      finish(true);
      return;
    }

    // Off mode or no autoFileOnSend (legacy): skip filing
    const shouldFile = intent.filingOnSend === "always" || intent.autoFileOnSend;
    if (!shouldFile) {
      console.log("[onMessageSendHandler] Filing not requested (mode=off or autoFileOnSend=false), skipping");
      finish(true);
      return;
    }

    // Migrate intent from fallback key to real itemId if needed
    try {
      const isFallbackKey =
        intent.itemKey === "draft:current" || intent.itemKey === "last_compose";

      if (isFallbackKey) {
        // Find the real itemId from keys (first non-draft key)
        const realItemId = keys.find((k) => !k.startsWith("draft:") && k !== "last_compose");

        if (realItemId) {
          console.log("[onMessageSendHandler] Migrating intent from fallback", {
            from: intent.itemKey,
            to: realItemId,
          });

          const intentValue = JSON.stringify({
            caseId: intent.caseId,
            autoFileOnSend: intent.autoFileOnSend,
            baseCaseId: intent.baseCaseId || "",
            baseEmailDocId: intent.baseEmailDocId || "",
          });

          const realKey = `sc_intent:${realItemId}`;

          // Save under real itemId
          if (typeof OfficeRuntime !== "undefined" && (OfficeRuntime as any)?.storage) {
            await (OfficeRuntime as any).storage.setItem(realKey, intentValue);
            console.log("[onMessageSendHandler] Migrated to OfficeRuntime.storage");
          } else if (Office?.context?.roamingSettings) {
            Office.context.roamingSettings.set(realKey, intentValue);
            await new Promise<void>((resolve) => {
              Office.context.roamingSettings.saveAsync(() => resolve());
            });
            console.log("[onMessageSendHandler] Migrated to roamingSettings");
          }

          // Clear fallback key
          const fallbackKey = `sc_intent:${intent.itemKey}`;
          if (typeof OfficeRuntime !== "undefined" && (OfficeRuntime as any)?.storage) {
            await (OfficeRuntime as any).storage.removeItem(fallbackKey);
            console.log("[onMessageSendHandler] Cleared fallback key from OfficeRuntime.storage");
          } else if (Office?.context?.roamingSettings) {
            Office.context.roamingSettings.remove(fallbackKey);
            await new Promise<void>((resolve) => {
              Office.context.roamingSettings.saveAsync(() => resolve());
            });
            console.log("[onMessageSendHandler] Cleared fallback key from roamingSettings");
          }
        }
      }
    } catch (e) {
      console.warn("[onMessageSendHandler] Intent migration failed (non-critical):", e);
    }

    console.log("[onMessageSendHandler] Getting auth token");
    const { token } = await withTimeout(getAuthRuntime(), 900);
    if (!token) {
      console.error("[onMessageSendHandler] No auth token available");
      await showInfo("SingleCase: chybí přihlášení, nelze zařadit při odeslání.");
      finish(true);
      return;
    }
    console.log("[onMessageSendHandler] Token retrieved", { tokenPrefix: token.slice(0, 10) });

    console.log("[onMessageSendHandler] Getting workspace host");
    const hostRaw = (await getStored(STORAGE_KEYS.workspaceHost)) || "";
    const host = normalizeHost(hostRaw);
    console.log("[onMessageSendHandler] Workspace host", { hostRaw, normalized: host });

    if (!host) {
      console.error("[onMessageSendHandler] No workspace host configured");
      await showInfo("SingleCase: chybí workspace URL, nelze zařadit při odeslání.");
      finish(true);
      return;
    }

    console.log("[onMessageSendHandler] Skipping pre-flight check, proceeding to upload");

    console.log("[onMessageSendHandler] Reading email metadata");
    console.log("[onMessageSendHandler] Current item info:", {
      itemType: (Office.context.mailbox.item as any)?.itemType,
      itemClass: (Office.context.mailbox.item as any)?.itemClass,
      hasSubject: !!(Office.context.mailbox.item as any)?.subject,
    });
    const subject = await withTimeout(getSubjectRuntime(), T_SUBJECT_MS);
    const bodyText = await withTimeout(getBodyTextRuntime(), T_BODY_MS);

    // For new compose, item.from is empty - fallback to userProfile
    const itemFrom = (Office.context.mailbox.item as any)?.from;
    const fromEmail = String(
      itemFrom?.emailAddress ||
      Office.context.mailbox.userProfile?.emailAddress ||
      ""
    );
    const fromName = String(
      itemFrom?.displayName ||
      Office.context.mailbox.userProfile?.displayName ||
      ""
    );

    // Extract conversationId (cross-mailbox identifier, available at send time)
    const conversationId = getConversationIdSafe();

    console.log("[onMessageSendHandler] Email metadata", {
      subject,
      fromEmail,
      fromName,
      bodyLength: bodyText.length,
      hasConversationId: !!conversationId,
      conversationIdPreview: conversationId ? conversationId.substring(0, 30) + "..." : "(none)",
    });

    const baseName = safeFileName(subject || "email");

    const emailText =
      `From: ${fromName} <${fromEmail}>\r\n` +
      `To: SingleCase <noreply@singlecase>\r\n` +
      `Subject: ${subject}\r\n` +
      `Date: ${new Date().toUTCString()}\r\n` +
      `Message-ID: <${keys[0]}@outlook>\r\n` +
      `MIME-Version: 1.0\r\n` +
      `Content-Type: text/plain; charset=UTF-8\r\n` +
      `Content-Transfer-Encoding: 8bit\r\n` +
      `\r\n` +
      `${(bodyText || "").trim()}\r\n`;

    const emailBase64 = toBase64Utf8(emailText);
    console.log("[onMessageSendHandler] EML built", { length: emailBase64.length });

    // NEW: Subject-based versioning decision
    // Check if a document with this subject already exists in the case
    let existingDoc: Awaited<ReturnType<typeof findDocumentBySubject>> = null;

    try {
      console.log("[onMessageSendHandler] Checking for existing document with same subject");
      existingDoc = await withTimeout(
        findDocumentBySubject(intent.caseId, subject),
        T_FETCH_MS
      );

      if (existingDoc) {
        console.log("[onMessageSendHandler] Found existing document", {
          docId: existingDoc.id,
          docName: existingDoc.name,
          docSubject: existingDoc.subject,
        });
      } else {
        console.log("[onMessageSendHandler] No existing document with this subject found");
      }
    } catch (e) {
      console.warn("[onMessageSendHandler] Failed to check for existing document:", e);
      // Continue with new document creation on error
      existingDoc = null;
    }

    const shouldUploadVersion = !!existingDoc;

    console.log("[onMessageSendHandler] Version decision", {
      caseId: intent.caseId,
      subject,
      existingDocId: existingDoc?.id,
      existingDocName: existingDoc?.name,
      shouldUploadVersion,
    });

    // ── Duplicate handling (always-file mode only) ────────────────────────
    // existingDoc means a document with the same subject/filename already
    // exists in this case.  Apply the duplicates setting from the intent.
    const dupMode = String((intent as any).duplicates || "warn");
    if (shouldUploadVersion && dupMode !== "off") {
      if (dupMode === "block") {
        console.log("[onMessageSendHandler] Duplicate=block: skipping filing");
        await showInfo(
          "SingleCase: this email already exists in the case. Filing was skipped (duplicate blocked)."
        );
        finish(true); // email still sends
        return;
      }
      // warn: defer filing so the user can confirm in the add-in after sending
      if (dupMode === "warn") {
        console.log("[onMessageSendHandler] Duplicate=warn: deferring filing to user");
        try {
          const convForPending = getConversationIdSafe();
          await setStored(
            "sc_pending_filing",
            JSON.stringify({
              caseId: intent.caseId,
              subject,
              conversationId: convForPending,
              sentAt: new Date().toISOString(),
            })
          );
        } catch (e) {
          console.warn("[onMessageSendHandler] Failed to store pending filing for duplicate:", e);
        }
        await showInfo(
          "SingleCase: a duplicate was detected. Open the panel to confirm filing."
        );
        finish(true);
        return;
      }
    }
    // ─────────────────────────────────────────────────────────────────────

    if (shouldUploadVersion && existingDoc) {
      // Upload as new version of existing document
      console.log("[onMessageSendHandler] Uploading as version of existing document:", existingDoc.id);

      await withTimeout(
        uploadDocumentVersion({
          documentId: existingDoc.id,
          fileName: `${baseName}.eml`,
          mimeType: "message/rfc822",
          dataBase64: emailBase64,
        }),
        T_FETCH_MS
      );

      console.log("[onMessageSendHandler] Version uploaded successfully");

      // Record recipient history for future compose preselection
      try {
        const recipients = await getRecipientsRuntime();
        if (recipients.length > 0) {
          await recordRecipientsFiledToCase(recipients, intent.caseId);
          console.log("[onMessageSendHandler] Recipient history recorded (version)", { count: recipients.length });
        }
      } catch (e) {
        console.warn("[onMessageSendHandler] Failed to record recipient history:", e);
      }

      // Update filed context with the existing document ID
      await persistFiledCtx(intent.caseId, existingDoc.id);

      // NEW: Cache filed email for "already filed" detection
      const conversationId = getConversationIdSafe();
      if (conversationId) {
        await cacheFiledEmail(
          conversationId,
          intent.caseId,
          existingDoc.id,
          subject
        );
        console.log("[onMessageSendHandler] Cached filed email (version)", { conversationId: conversationId.substring(0, 20) + "..." });
      } else {
        // Fallback: Cache by subject when conversationId not available (new compose emails)
        await cacheFiledEmailBySubject(
          subject,
          intent.caseId,
          existingDoc.id
        );
        console.log("[onMessageSendHandler] Cached filed email by subject (version)", { subject });
      }
    } else {
      // Upload as new document
      console.log("[onMessageSendHandler] Uploading as new document");

      const created = await withTimeout(
        uploadDocumentToCase({
          caseId: intent.caseId,
          fileName: `${baseName}.eml`,
          mimeType: "message/rfc822",
          dataBase64: emailBase64,
          metadata: {
            subject,
            fromEmail,
            fromName,
            conversationId: conversationId || undefined, // Cross-mailbox identifier
          },
        }),
        T_FETCH_MS
      );

      const docs = (created as any)?.documents;
      const createdDocId = Array.isArray(docs) && docs[0]?.id ? String(docs[0].id) : "";

      console.log("[onMessageSendHandler] Created docId", { createdDocId, rawResponse: created });

      if (createdDocId) {
        // Record recipient history for future compose preselection
        try {
          const recipients = await getRecipientsRuntime();
          if (recipients.length > 0) {
            await recordRecipientsFiledToCase(recipients, intent.caseId);
            console.log("[onMessageSendHandler] Recipient history recorded (new doc)", { count: recipients.length });
          }
        } catch (e) {
          console.warn("[onMessageSendHandler] Failed to record recipient history:", e);
        }

        await persistFiledCtx(intent.caseId, createdDocId);

        // NEW: Cache filed email for "already filed" detection
        const conversationId = getConversationIdSafe();
        if (conversationId) {
          await cacheFiledEmail(
            conversationId,
            intent.caseId,
            createdDocId,
            subject
          );
          console.log("[onMessageSendHandler] Cached filed email (new doc)", { conversationId: conversationId.substring(0, 20) + "..." });
        } else {
          // Fallback: Cache by subject when conversationId not available (new compose emails)
          await cacheFiledEmailBySubject(
            subject,
            intent.caseId,
            createdDocId
          );
          console.log("[onMessageSendHandler] Cached filed email by subject (new doc)", { subject });
        }
      }
    }

    console.log("[onMessageSendHandler] Upload successful");
    await showInfo("SingleCase: email uložen při odeslání.");

    finish(true);
  } catch (e) {
    console.error("[onMessageSendHandler] Error during filing", e);

    try {
      const msg = e instanceof Error ? e.message : String(e);
      let errorHint = "";

      if (msg.includes("timeout")) errorHint = " (timeout)";
      else if (msg.toLowerCase().includes("workspace")) errorHint = " (není nastaven workspace)";
      else if (msg.toLowerCase().includes("token")) errorHint = " (přihlaste se znovu)";
      else if (msg.toLowerCase().includes("network")) errorHint = " (problém se sítí)";

      await showInfo(`SingleCase: nepodařilo se uložit${errorHint}`);
    } catch {
      // ignore
    }

    finish(true);
  }
}