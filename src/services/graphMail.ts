// src/services/graphMail.ts
declare const Office: any;
declare const OfficeRuntime: any;

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// === Category names (KEEP THESE CONSISTENT EVERYWHERE) ===
const CATEGORY_NAME = "SC: Filed";
const CATEGORY_NAME_UNFILED = "SC: Unfiled";

// Graph preset colour (for master category creation via Graph)
const CATEGORY_COLOR_FILED = "preset4"; // green

/* =========================================================
   =============== OFFICE.JS (Immediate UI) =================
   ========================================================= */

// Apply "Filed" category immediately in Outlook UI
export async function applyFiledCategoryToCurrentEmailOfficeJs(): Promise<void> {
  const item = Office?.context?.mailbox?.item as any;
  if (!item) throw new Error("No current Outlook item.");

  // 1) Ensure master category exists (best effort)
  await new Promise<void>((resolve) => {
    try {
      const mc = Office?.context?.mailbox?.masterCategories;
      if (!mc?.addAsync) return resolve();

      mc.addAsync(
        [{
          displayName: CATEGORY_NAME,
          color: Office.MailboxEnums.CategoryColor.Preset4
        }],
        () => resolve()
      );
    } catch {
      resolve();
    }
  });

  // 2) Remove "Unfiled" if present
  await new Promise<void>((resolve) => {
    try {
      const cats = item?.categories;
      if (!cats?.removeAsync) return resolve();
      cats.removeAsync([CATEGORY_NAME_UNFILED], () => resolve());
    } catch {
      resolve();
    }
  });

  // 3) Add "Filed"
  await new Promise<void>((resolve, reject) => {
    try {
      const cats = item?.categories;
      if (!cats?.addAsync) return reject(new Error("Categories API unavailable."));

      cats.addAsync([CATEGORY_NAME], (res: any) => {
        if (res?.status === Office.AsyncResultStatus.Succeeded) return resolve();
        reject(new Error(res?.error?.message || "categories.addAsync failed"));
      });
    } catch (e) {
      reject(e);
    }
  });
}

// Apply "Unfiled" category immediately in Outlook UI
export async function applyUnfiledCategoryToCurrentEmailOfficeJs(): Promise<void> {
  const item = Office?.context?.mailbox?.item as any;
  if (!item) throw new Error("No current Outlook item.");

  // 1) Ensure master category exists (best effort)
  await new Promise<void>((resolve) => {
    try {
      const mc = Office?.context?.mailbox?.masterCategories;
      if (!mc?.addAsync) return resolve();

      mc.addAsync(
        [{
          displayName: CATEGORY_NAME_UNFILED,
          color: Office.MailboxEnums.CategoryColor.Preset7
        }],
        () => resolve()
      );
    } catch {
      resolve();
    }
  });

  // 2) Remove "Filed" if present
  await new Promise<void>((resolve) => {
    try {
      const cats = item?.categories;
      if (!cats?.removeAsync) return resolve();
      cats.removeAsync([CATEGORY_NAME], () => resolve());
    } catch {
      resolve();
    }
  });

  // 3) Add "Unfiled"
  await new Promise<void>((resolve, reject) => {
    try {
      const cats = item?.categories;
      if (!cats?.addAsync) return reject(new Error("Categories API unavailable."));

      cats.addAsync([CATEGORY_NAME_UNFILED], (res: any) => {
        if (res?.status === Office.AsyncResultStatus.Succeeded) return resolve();
        reject(new Error(res?.error?.message || "categories.addAsync failed"));
      });
    } catch (e) {
      reject(e);
    }
  });
}

/* =========================================================
   ================= GRAPH API VERSION =====================
   ========================================================= */

function getRestIdForCurrentItem(): string {
  const item = Office?.context?.mailbox?.item as any;
  const itemId = String(item?.itemId || "");
  if (!itemId) return "";

  try {
    return Office.context.mailbox.convertToRestId(
      itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  } catch {
    return "";
  }
}

async function graphFetch(
  accessToken: string,
  path: string,
  init?: RequestInit
): Promise<Response> {
  return fetch(`${GRAPH_BASE}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(init?.headers || {}),
    },
  });
}

async function readJsonSafe(res: Response): Promise<any> {
  const text = await res.text().catch(() => "");
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

async function ensureMasterCategory(
  accessToken: string,
  displayName: string,
  color: string
): Promise<void> {
  const listRes = await graphFetch(
    accessToken,
    `/me/outlook/masterCategories?$filter=displayName eq '${displayName.replace(/'/g, "''")}'`
  );

  if (listRes.ok) {
    const json = await readJsonSafe(listRes);
    if (Array.isArray(json?.value) && json.value.length > 0) return;
  }

  const createRes = await graphFetch(accessToken, `/me/outlook/masterCategories`, {
    method: "POST",
    body: JSON.stringify({ displayName, color }),
  });

  if (createRes.ok || createRes.status === 409) return;

  const err = await readJsonSafe(createRes);
  throw new Error(err?.error?.message || `Failed to create category (${createRes.status})`);
}

async function getGraphAccessToken(): Promise<string> {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime?.auth?.getAccessToken) {
      const token = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        forMSGraphAccess: true,
      });
      return String(token || "");
    }
  } catch (e) {
    console.warn("[graphMail] OfficeRuntime.auth.getAccessToken failed:", e);
  }
  return "";
}

/**
 * Read current message categories via Graph (fallback for Mac / Office.js category reads).
 * Returns [] if not available or if current item cannot be resolved to a message id.
 */
export async function getCurrentEmailCategoriesGraph(): Promise<string[]> {
  try {
    const restId = getRestIdForCurrentItem();
    if (!restId) return [];

    const token = await getGraphAccessToken();
    if (!token) return [];

    const res = await graphFetch(
      token,
      `/me/messages/${encodeURIComponent(restId)}?$select=categories`
    );

    if (!res.ok) {
      const err = await readJsonSafe(res);
      console.warn("[graphMail] getCurrentEmailCategoriesGraph not ok", res.status, err);
      return [];
    }

    const json = await readJsonSafe(res);
    const cats = Array.isArray(json?.categories) ? json.categories : [];
    return cats.map((c: any) => String(c || "")).filter(Boolean);
  } catch (e) {
    console.warn("[graphMail] getCurrentEmailCategoriesGraph failed:", e);
    return [];
  }
}

// Apply "Unfiled" category via Graph (reliable on Mac / OWA — persists to Exchange)
export async function applyUnfiledCategoryToCurrentEmailGraph(): Promise<void> {
  const restId = getRestIdForCurrentItem();
  if (!restId) throw new Error("Cannot resolve REST id.");

  const token = await getGraphAccessToken();
  if (!token) throw new Error("Cannot get Graph access token.");

  await ensureMasterCategory(token, CATEGORY_NAME_UNFILED, "preset7");

  const patchRes = await graphFetch(token, `/me/messages/${encodeURIComponent(restId)}`, {
    method: "PATCH",
    body: JSON.stringify({ categories: [CATEGORY_NAME_UNFILED] }),
  });

  if (!patchRes.ok) {
    const err = await readJsonSafe(patchRes);
    throw new Error(err?.error?.message || `Failed to apply Unfiled category (${patchRes.status})`);
  }
}

// Apply Filed via Graph (caller provides accessToken)
export async function applyFiledCategoryToCurrentEmail(accessToken: string): Promise<void> {
  const restId = getRestIdForCurrentItem();
  if (!restId) throw new Error("Cannot resolve REST id.");

  await ensureMasterCategory(accessToken, CATEGORY_NAME, CATEGORY_COLOR_FILED);

  const patchRes = await graphFetch(accessToken, `/me/messages/${encodeURIComponent(restId)}`, {
    method: "PATCH",
    body: JSON.stringify({ categories: [CATEGORY_NAME] }),
  });

  if (!patchRes.ok) {
    const err = await readJsonSafe(patchRes);
    throw new Error(err?.error?.message || `Failed to apply category (${patchRes.status})`);
  }
}

export function getFiledCategoryName(): string {
  return CATEGORY_NAME;
}

export function getUnfiledCategoryName(): string {
  return CATEGORY_NAME_UNFILED;
}

/* =========================================================
   ======= Cross-Mailbox Filed Detection (internetMessageId) =======
   ========================================================= */

/**
 * Get internetMessageId for a specific message via Graph API
 * This is the stable cross-mailbox identifier (Message-ID header)
 *
 * @param messageId - Graph API message ID (REST ID format)
 * @returns internetMessageId or null if not available
 */
export async function getInternetMessageIdViaGraph(messageId: string): Promise<string | null> {
  try {
    console.log("[getInternetMessageIdViaGraph] Fetching for message:", messageId.substring(0, 50) + "...");

    const accessToken = await getGraphAccessToken();
    const url = `/me/messages/${encodeURIComponent(messageId)}?$select=internetMessageId`;

    const res = await graphFetch(accessToken, url, { method: "GET" });

    if (!res.ok) {
      console.warn("[getInternetMessageIdViaGraph] Failed:", res.status);
      return null;
    }

    const json = await res.json();
    const internetMessageId = json?.internetMessageId;

    if (internetMessageId) {
      console.log("[getInternetMessageIdViaGraph] Found:", internetMessageId);
      return String(internetMessageId);
    }

    console.log("[getInternetMessageIdViaGraph] No internetMessageId in response");
    return null;
  } catch (e) {
    console.error("[getInternetMessageIdViaGraph] Error:", e);
    return null;
  }
}

/**
 * Apply "Filed" category to a specific message by Graph API message ID
 * Used when we detect an email is already filed (e.g., sender filed it)
 *
 * @param messageId - Graph API message ID (REST ID format)
 */
export async function applyCategoryToMessageById(messageId: string): Promise<void> {
  try {
    console.log("[applyCategoryToMessageById] Applying category to:", messageId.substring(0, 50) + "...");

    const accessToken = await getGraphAccessToken();

    // Ensure master category exists
    await ensureMasterCategory(accessToken, CATEGORY_NAME, CATEGORY_COLOR_FILED);

    // Apply category to the specific message
    const patchRes = await graphFetch(accessToken, `/me/messages/${encodeURIComponent(messageId)}`, {
      method: "PATCH",
      body: JSON.stringify({ categories: [CATEGORY_NAME] }),
    });

    if (!patchRes.ok) {
      const err = await readJsonSafe(patchRes);
      console.warn("[applyCategoryToMessageById] Failed:", err?.error?.message || patchRes.status);
      throw new Error(err?.error?.message || `Failed to apply category (${patchRes.status})`);
    }

    console.log("[applyCategoryToMessageById] Category applied successfully");
  } catch (e) {
    console.error("[applyCategoryToMessageById] Error:", e);
    throw e;
  }
}

/**
 * Search Sent Items for a message by conversationId and subject
 * Returns the internetMessageId if found
 *
 * Useful fallback when internetMessageId is not available before send
 *
 * @param conversationId - Office.js conversationId
 * @param subject - Email subject
 * @param sentAfter - ISO timestamp to narrow search window
 */
export async function findInternetMessageIdInSentItems(
  conversationId: string,
  subject: string,
  sentAfter: string
): Promise<string | null> {
  try {
    console.log("[findInternetMessageIdInSentItems] Searching", {
      conversationId: conversationId.substring(0, 30) + "...",
      subject,
      sentAfter,
    });

    const accessToken = await getGraphAccessToken();

    // Build filter query
    const subjectEscaped = subject.replace(/'/g, "''");
    const filter = `subject eq '${subjectEscaped}' and sentDateTime ge ${sentAfter}`;
    const url = `/me/mailFolders/SentItems/messages?$filter=${encodeURIComponent(filter)}&$select=internetMessageId,subject,sentDateTime&$top=5`;

    const res = await graphFetch(accessToken, url, { method: "GET" });

    if (!res.ok) {
      console.warn("[findInternetMessageIdInSentItems] Failed:", res.status);
      return null;
    }

    const json = await res.json();
    const messages = json?.value || [];

    console.log("[findInternetMessageIdInSentItems] Found", messages.length, "messages in window");

    // Return first match (most recent)
    if (messages.length > 0 && messages[0].internetMessageId) {
      const internetMessageId = String(messages[0].internetMessageId);
      console.log("[findInternetMessageIdInSentItems] Match:", internetMessageId);
      return internetMessageId;
    }

    console.log("[findInternetMessageIdInSentItems] No match found");
    return null;
  } catch (e) {
    console.error("[findInternetMessageIdInSentItems] Error:", e);
    return null;
  }
}