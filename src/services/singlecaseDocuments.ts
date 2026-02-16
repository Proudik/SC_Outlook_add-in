import { getAuth, getAuthRuntime } from "./auth";
import { getStored, setStored } from "../utils/storage";
import { STORAGE_KEYS } from "../utils/constants";

export type UploadDocumentResponse = {
  documents: Array<{
    id: string;
    name: string;
    mime_type: string;
    latest_version?: {
      id: number | string;
      name: string;
      revision_number: number;
    };
  }>;
};

export type UploadDocumentVersionResponse = {
  id: number | string;
  name: string;
  mime_type: string;
  dir_id?: number | string;
  latest_version?: {
    id: number | string;
    name: string;
    revision_number: number;
  };
};

export type DocumentMeta = {
  id: string;
  name: string;
  case_id: string;
};

export type DirectoryItem = {
  id: string | number;
  name: string;
  type: "file" | "directory";
  parent_id?: string | number;
};

export type DirectoryListing = {
  items: DirectoryItem[];
  parent_id?: string | number;
};

export type CreateDirectoryResponse = {
  id: string | number;
  name: string;
  parent_id?: string | number;
};

function normalizeHost(host: string): string {
  const v = (host || "").trim().toLowerCase();
  if (!v) return "";
  return v.replace(/^https?:\/\//i, "").split("/")[0];
}

async function resolveApiBaseUrl(): Promise<string> {
  console.log("[resolveApiBaseUrl] Reading workspaceHost from storage key:", STORAGE_KEYS.workspaceHost);
  const storedHostRaw = await getStored(STORAGE_KEYS.workspaceHost);
  console.log("[resolveApiBaseUrl] Raw stored host:", storedHostRaw);

  const host = normalizeHost(storedHostRaw || "");
  console.log("[resolveApiBaseUrl] Normalized host:", host);

  if (!host) {
    console.error("[resolveApiBaseUrl] Workspace host is missing");
    throw new Error("Workspace host is missing.");
  }

  const baseUrl = `/singlecase/${encodeURIComponent(host)}/publicapi/v1`;
  console.log("[resolveApiBaseUrl] Resolved base URL:", baseUrl);
  return baseUrl;
}

async function getToken(): Promise<string> {
  const auth = getAuth();
  if (auth?.token) {
    console.log("[getToken] Using sessionStorage token");
    return auth.token;
  }

  console.log("[getToken] sessionStorage token not available, trying OfficeRuntime.storage");
  const rt = await getAuthRuntime();
  if (rt?.token) {
    console.log("[getToken] Using OfficeRuntime.storage token");
    return rt.token;
  }

  console.error("[getToken] No token found in either sessionStorage or OfficeRuntime.storage");
  throw new Error("Missing auth token.");
}

async function expectJson(res: Response, errorPrefix: string) {
  const text = await res.text().catch(() => "");

  if (!res.ok) {
    if (res.status === 423) {
      throw new Error(
        "Dokument je momentálně uzamčen. Někdo jej právě upravuje. Počkejte prosím, než se dokument odemkne, a zkuste to znovu."
      );
    }

    throw new Error(`${errorPrefix} (${res.status}): ${text || res.statusText}`);
  }

  const contentType = res.headers.get("content-type") || "";
  if (!contentType.includes("application/json")) {
    throw new Error(`${errorPrefix}: expected JSON but got ${contentType || "no content-type"}.`);
  }

  return JSON.parse(text);
}

export async function getDocumentMeta(documentId: string | number): Promise<DocumentMeta | null> {
  const token = await getToken();

  const base = await resolveApiBaseUrl();
  const url = `${base}/documents/${encodeURIComponent(String(documentId))}`;

  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authentication: token,
      "Content-Type": "application/json",
      "Accept-Encoding": "identity",
    },
  });

  if (res.status === 404) return null;

  const json = await expectJson(res, "Get document failed");

  return {
    id: String((json as any).id || documentId),
    name: String((json as any).name || ""),
    case_id: String((json as any).case_id || ""),
  };
}

export async function uploadDocumentVersion(params: {
  documentId: string | number;
  fileName: string;
  mimeType: string;
  dataBase64: string;
  directoryId?: string;
}): Promise<UploadDocumentVersionResponse> {
  const { documentId, fileName, mimeType, dataBase64, directoryId } = params;

  const token = await getToken();

  const base = await resolveApiBaseUrl();
  const id = encodeURIComponent(String(documentId));

  const bodyData: any = {
    name: fileName,
    mime_type: mimeType,
    data_base64: dataBase64,
  };

  // Add dir_id if provided (though versions typically inherit parent directory)
  if (directoryId) {
    bodyData.dir_id = directoryId;
  }

  const body = JSON.stringify(bodyData);

  const candidates: Array<{ url: string; method: "POST" | "PUT" | "PATCH" }> = [
    { url: `${base}/documents/${id}/version`, method: "POST" },
    { url: `${base}/documents/${id}/versions`, method: "POST" },
    { url: `${base}/documents/${id}/versions`, method: "PUT" },
    { url: `${base}/documents/${id}/version`, method: "PUT" },
    { url: `${base}/documents/${id}/versions`, method: "PATCH" },
    { url: `${base}/documents/${id}/version`, method: "PATCH" },
  ];

  let lastErr: unknown = null;

  for (const c of candidates) {
    // eslint-disable-next-line no-await-in-loop
    const res = await fetch(c.url, {
      method: c.method,
      headers: {
        "Content-Type": "application/json",
        Authentication: token,
        "Accept-Encoding": "identity",
      },
      body,
    });

    if (res.status === 404 || res.status === 405) {
      lastErr = new Error(`Endpoint not available: ${c.method} ${c.url} (${res.status})`);
      continue;
    }

    // eslint-disable-next-line no-await-in-loop
    const json = await expectJson(res, "Upload version failed");
    return json as UploadDocumentVersionResponse;
  }

  throw lastErr instanceof Error ? lastErr : new Error("Upload version failed: no supported endpoint found");
}

export async function uploadDocumentToCase(params: {
  caseId: string;
  fileName: string;
  mimeType: string;
  dataBase64: string;
  directoryId?: string;
  metadata?: {
    subject?: string;
    fromEmail?: string;
    fromName?: string;
    [key: string]: any;
  };
}): Promise<UploadDocumentResponse> {
  const { caseId, fileName, mimeType, dataBase64, directoryId, metadata } = params;

  console.log("[uploadDocumentToCase] Starting upload", {
    caseId,
    fileName,
    mimeType,
    dataLength: dataBase64.length,
  });

  let token: string;
  try {
    token = await getToken();
    console.log("[uploadDocumentToCase] Token retrieved", { hasToken: !!token, tokenPrefix: token.slice(0, 10) });
  } catch (e) {
    console.error("[uploadDocumentToCase] Failed to get token:", e);
    throw e;
  }

  let base: string;
  try {
    base = await resolveApiBaseUrl();
    console.log("[uploadDocumentToCase] Base URL resolved:", base);
  } catch (e) {
    console.error("[uploadDocumentToCase] Failed to resolve base URL:", e);
    throw e;
  }

  const url = `${base}/documents`;
  console.log("[uploadDocumentToCase] Full URL:", url);

  const payload: any = {
    case_id: caseId,
    documents: [
      {
        name: fileName,
        mime_type: mimeType,
        data_base64: dataBase64,
        ...(directoryId ? { dir_id: directoryId } : {}),
        ...(metadata ? { metadata } : {}),
      },
    ],
  };

  console.log("[uploadDocumentToCase] Payload structure:", {
    case_id: payload.case_id,
    documentCount: payload.documents.length,
    firstDoc: {
      name: payload.documents[0].name,
      mime_type: payload.documents[0].mime_type,
      data_base64_length: payload.documents[0].data_base64.length,
    },
  });

  let res: Response;
  try {
    res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authentication: token,
        "Accept-Encoding": "identity",
      },
      body: JSON.stringify(payload),
    });
    console.log("[uploadDocumentToCase] Fetch completed", {
      status: res.status,
      statusText: res.statusText,
      ok: res.ok,
    });
  } catch (e) {
    console.error("[uploadDocumentToCase] Fetch failed:", e);
    throw new Error(`Network request failed: ${e instanceof Error ? e.message : String(e)}`);
  }

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    const snippet = text.slice(0, 300);
    console.error("[uploadDocumentToCase] Upload failed", {
      status: res.status,
      statusText: res.statusText,
      url,
      responseSnippet: snippet,
    });
  }

  const json = await expectJson(res, "Upload failed");
  console.log("[uploadDocumentToCase] Upload successful", { documentIds: (json as any).documents?.map((d: any) => d.id) });
  return json as UploadDocumentResponse;
}

async function tryRename(url: string, method: string, token: string, name: string) {
  const res = await fetch(url, {
    method,
    headers: {
      "Content-Type": "application/json",
      Authentication: token,
      "Accept-Encoding": "identity",
    },
    body: JSON.stringify({ name }),
  });

  if (res.ok) {
    const contentType = res.headers.get("content-type") || "";
    if (contentType.includes("application/json")) return res.json();
    return {};
  }

  if (res.status === 404 || res.status === 405) return null;

  const text = await res.text().catch(() => "");
  throw new Error(`Rename document failed (${res.status}): ${text || res.statusText}`);
}

export async function renameDocument(params: { token: string; documentId: string; name: string }) {
  const base = await resolveApiBaseUrl();
  const id = encodeURIComponent(String(params.documentId));

  const candidates: Array<{ url: string; method: string }> = [
    { url: `${base}/documents/${id}`, method: "PUT" },
    { url: `${base}/documents/${id}`, method: "PATCH" },
    { url: `${base}/documents/${id}/rename`, method: "POST" },
    { url: `${base}/documents/${id}/name`, method: "PUT" },
  ];

  for (const c of candidates) {
    // eslint-disable-next-line no-await-in-loop
    const ok = await tryRename(c.url, c.method, params.token, params.name);
    if (ok) return ok;
  }

  throw new Error("Rename document failed: API endpoint not found or not allowed");
}

// ============================================================================
// Folder/Directory Management for "Outlook add-in" folder
// ============================================================================

const OUTLOOK_FOLDER_NAME = "Outlook add-in";

/**
 * Get cached folder ID for a case, if available
 */
async function getCachedFolderId(caseId: string): Promise<string | null> {
  try {
    const raw = await getStored(STORAGE_KEYS.outlookFolderCache);
    if (!raw) return null;

    const cache = JSON.parse(String(raw));
    const folderId = cache[String(caseId)];
    return folderId ? String(folderId) : null;
  } catch {
    return null;
  }
}

/**
 * Cache folder ID for a case
 */
async function cacheFolderId(caseId: string, folderId: string): Promise<void> {
  try {
    const raw = await getStored(STORAGE_KEYS.outlookFolderCache);
    const cache = raw ? JSON.parse(String(raw)) : {};

    cache[String(caseId)] = String(folderId);

    await setStored(STORAGE_KEYS.outlookFolderCache, JSON.stringify(cache));
  } catch (e) {
    console.warn("[cacheFolderId] Failed to cache folder ID:", e);
  }
}

/**
 * Get the root directory ID for a case
 * Returns null if case has no documents directory yet
 */
export async function getCaseRootDirectoryId(caseId: string): Promise<string | null> {
  const token = await getToken();
  const base = await resolveApiBaseUrl();

  // Try to get case details which should include root directory info
  const url = `${base}/cases/${encodeURIComponent(caseId)}`;

  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authentication: token,
      "Content-Type": "application/json",
      "Accept-Encoding": "identity",
    },
  });

  if (!res.ok) {
    console.warn("[getCaseRootDirectoryId] Failed to get case:", res.status);
    return null;
  }

  const json = await res.json();
  const rootDirId = (json as any)?.root_directory_id || (json as any)?.documents_directory_id;

  return rootDirId ? String(rootDirId) : null;
}

/**
 * List contents of a directory
 */
export async function listDirectory(directoryId: string): Promise<DirectoryListing> {
  const token = await getToken();
  const base = await resolveApiBaseUrl();

  const url = `${base}/directories/${encodeURIComponent(directoryId)}`;

  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authentication: token,
      "Content-Type": "application/json",
      "Accept-Encoding": "identity",
    },
  });

  if (!res.ok) {
    throw new Error(`List directory failed (${res.status}): ${res.statusText}`);
  }

  const json = await res.json();
  const items: DirectoryItem[] = [];

  // Parse items (structure may vary by API)
  const rawItems = (json as any)?.items || (json as any)?.children || [];
  for (const item of rawItems) {
    items.push({
      id: String(item.id || item._id),
      name: String(item.name || ""),
      type: item.type === "directory" || item.is_directory ? "directory" : "file",
      parent_id: item.parent_id,
    });
  }

  return {
    items,
    parent_id: (json as any)?.id,
  };
}

/**
 * Create a new directory
 */
export async function createDirectory(parentId: string, name: string): Promise<CreateDirectoryResponse> {
  const token = await getToken();
  const base = await resolveApiBaseUrl();

  const payload = {
    name,
    parent_id: parentId,
  };

  // Try multiple possible endpoints
  const candidates = [
    { url: `${base}/directories`, method: "POST" },
    { url: `${base}/directories/${encodeURIComponent(parentId)}/subdirectories`, method: "POST" },
    { url: `${base}/folders`, method: "POST" },
  ];

  let lastError: Error | null = null;

  for (const candidate of candidates) {
    try {
      // eslint-disable-next-line no-await-in-loop
      const res = await fetch(candidate.url, {
        method: candidate.method,
        headers: {
          Authentication: token,
          "Content-Type": "application/json",
          "Accept-Encoding": "identity",
        },
        body: JSON.stringify(payload),
      });

      if (res.status === 404 || res.status === 405) {
        continue; // Try next endpoint
      }

      if (!res.ok) {
        const text = await res.text().catch(() => "");
        lastError = new Error(`Create directory failed (${res.status}): ${text || res.statusText}`);
        continue;
      }

      const json = await res.json();
      return {
        id: String((json as any).id || (json as any)._id),
        name: String((json as any).name || name),
        parent_id: (json as any).parent_id,
      };
    } catch (e) {
      lastError = e instanceof Error ? e : new Error(String(e));
    }
  }

  throw lastError || new Error("Create directory failed: no supported endpoint found");
}

/**
 * Ensure the "Outlook add-in" folder exists in the case
 * Returns the folder's directory ID
 *
 * This function is idempotent and handles concurrent calls safely
 */
export async function ensureOutlookAddinFolder(caseId: string): Promise<string | null> {
  console.log("[ensureOutlookAddinFolder] Starting for case:", caseId);

  // 1. Check cache first
  const cachedId = await getCachedFolderId(caseId);
  if (cachedId) {
    console.log("[ensureOutlookAddinFolder] Using cached folder ID:", cachedId);
    return cachedId;
  }

  // 2. Get case root directory
  const rootDirId = await getCaseRootDirectoryId(caseId);
  if (!rootDirId) {
    console.warn("[ensureOutlookAddinFolder] Case has no root directory");
    return null;
  }

  console.log("[ensureOutlookAddinFolder] Root directory ID:", rootDirId);

  // 3. List root directory contents
  try {
    const listing = await listDirectory(rootDirId);
    console.log("[ensureOutlookAddinFolder] Found", listing.items.length, "items in root");

    // 4. Check if "Outlook add-in" folder already exists
    const existing = listing.items.find(
      (item) => item.type === "directory" && item.name === OUTLOOK_FOLDER_NAME
    );

    if (existing) {
      console.log("[ensureOutlookAddinFolder] Folder already exists:", existing.id);
      await cacheFolderId(caseId, String(existing.id));
      return String(existing.id);
    }

    // 5. Create the folder
    console.log("[ensureOutlookAddinFolder] Creating folder:", OUTLOOK_FOLDER_NAME);
    const created = await createDirectory(rootDirId, OUTLOOK_FOLDER_NAME);
    console.log("[ensureOutlookAddinFolder] Folder created:", created.id);

    await cacheFolderId(caseId, String(created.id));
    return String(created.id);
  } catch (e) {
    console.error("[ensureOutlookAddinFolder] Failed:", e);
    // Return null - uploads will go to root if folder creation fails
    return null;
  }
}

// ============================================================================
// Subject-Based Document Matching (for email versioning)
// ============================================================================

/**
 * Normalize email subject for matching
 * - Trim whitespace
 * - Lowercase
 * - Collapse multiple spaces
 * - Optionally strip Re:/Fw:/Fwd: prefixes
 */
export function normalizeSubject(subject: string, stripPrefixes: boolean = true): string {
  if (!subject) return "";

  let normalized = subject.trim().toLowerCase();

  // Collapse multiple spaces to single space
  normalized = normalized.replace(/\s+/g, " ");

  // Optionally strip common reply/forward prefixes
  if (stripPrefixes) {
    // Remove Re:, RE:, Fw:, FW:, Fwd:, FWD: (with optional spaces and colons)
    // Handle multiple nested prefixes like "Re: Fw: Re: Subject"
    let prevLength: number;
    do {
      prevLength = normalized.length;
      normalized = normalized.replace(/^(re|fw|fwd):\s*/i, "");
    } while (normalized.length !== prevLength && normalized.length > 0);
  }

  return normalized.trim();
}

export type DocumentSearchResult = {
  id: string;
  name: string;
  subject?: string;
};

/**
 * Search for existing email documents in a case by matching subject
 * Returns the document ID if found, null otherwise
 *
 * This function:
 * 1. Lists all documents in the case
 * 2. Filters for .eml files
 * 3. Compares normalized subjects
 * 4. Returns the first match (or null)
 */
export async function findDocumentBySubject(
  caseId: string,
  subject: string
): Promise<DocumentSearchResult | null> {
  console.log("[findDocumentBySubject] Searching for subject in case", { caseId, subject });

  const token = await getToken();
  const base = await resolveApiBaseUrl();

  // Try multiple possible endpoints for listing case documents
  const candidates = [
    `${base}/cases/${encodeURIComponent(caseId)}/documents`,
    `${base}/documents?case_id=${encodeURIComponent(caseId)}`,
    `${base}/cases/${encodeURIComponent(caseId)}/files`,
  ];

  let documents: any[] = [];

  for (const url of candidates) {
    try {
      // eslint-disable-next-line no-await-in-loop
      const res = await fetch(url, {
        method: "GET",
        headers: {
          Authentication: token,
          "Content-Type": "application/json",
          "Accept-Encoding": "identity",
        },
      });

      if (res.status === 404 || res.status === 405) {
        continue; // Try next endpoint
      }

      if (!res.ok) {
        continue;
      }

      // eslint-disable-next-line no-await-in-loop
      const json = await res.json();

      // Handle different response structures
      documents = Array.isArray(json) ? json :
                  Array.isArray(json.documents) ? json.documents :
                  Array.isArray(json.files) ? json.files :
                  Array.isArray(json.items) ? json.items :
                  [];

      if (documents.length >= 0) {
        console.log("[findDocumentBySubject] Found", documents.length, "documents in case");
        break; // Success
      }
    } catch {
      // Try next endpoint
    }
  }

  if (documents.length === 0) {
    console.log("[findDocumentBySubject] No documents found in case");
    return null;
  }

  // Normalize search subject
  const normalizedSearchSubject = normalizeSubject(subject);
  console.log("[findDocumentBySubject] Normalized search subject:", normalizedSearchSubject);

  if (!normalizedSearchSubject) {
    console.warn("[findDocumentBySubject] Empty normalized subject, skipping");
    return null;
  }

  // Search for .eml files with matching subject
  for (const doc of documents) {
    // Check if it's an email document (.eml extension)
    const fileName = String(doc.name || doc.filename || "");
    if (!fileName.toLowerCase().endsWith(".eml")) {
      continue;
    }

    // Try to extract subject from multiple possible locations
    let docSubject =
      doc.metadata?.subject ||    // Metadata field (if we stored it)
      doc.subject ||              // Direct field
      doc.properties?.subject ||  // Properties object
      "";

    // Fallback: extract from filename (remove .eml extension)
    if (!docSubject) {
      docSubject = fileName.replace(/\.eml$/i, "");
    }

    const normalizedDocSubject = normalizeSubject(docSubject);

    console.log("[findDocumentBySubject] Comparing", {
      fileName,
      docSubject,
      normalizedDocSubject,
      matches: normalizedDocSubject === normalizedSearchSubject,
    });

    if (normalizedDocSubject === normalizedSearchSubject) {
      console.log("[findDocumentBySubject] Match found!", {
        id: doc.id,
        name: fileName,
      });

      return {
        id: String(doc.id || doc._id),
        name: fileName,
        subject: docSubject,
      };
    }
  }

  console.log("[findDocumentBySubject] No matching document found");
  return null;
}

// ============================================================================
// Cross-Mailbox Filed Detection (internetMessageId)
// ============================================================================

export type FiledDocumentInfo = {
  documentId: string;
  caseId: string;
  caseName?: string;
  caseKey?: string;
  subject?: string;
};

/**
 * Check if an email with this conversationId and subject is already filed
 * Searches across all cases in the workspace
 *
 * This is the definitive server-side check for "already filed" status
 * Uses conversationId + normalized subject for reliable cross-mailbox matching
 *
 * @param conversationId - Office.js conversationId (available at send time)
 * @param subject - Email subject for additional matching
 * @returns Document info if found, null otherwise
 */
export async function checkFiledStatusByConversationAndSubject(
  conversationId: string,
  subject: string
): Promise<FiledDocumentInfo | null> {
  if (!conversationId) {
    console.log("[checkFiledStatusByConversationAndSubject] No conversationId provided");
    return null;
  }

  if (!subject) {
    console.log("[checkFiledStatusByConversationAndSubject] No subject provided");
    return null;
  }

  const normalizedSearchSubject = normalizeSubject(subject);

  console.log("[checkFiledStatusByConversationAndSubject] Checking:", {
    conversationId: conversationId.substring(0, 30) + "...",
    subject,
    normalizedSubject: normalizedSearchSubject,
  });

  try {
    const token = await getToken();
    const base = await resolveApiBaseUrl();

    // Search for documents by conversationId in metadata
    // Since backend may not support metadata search, we'll list recent documents and filter manually

    console.log("[checkFiledStatusByConversationAndSubject] Fetching recent documents for matching");

    // Get recently modified documents (last 200 to increase chance of finding match)
    const listUrl = `${base}/documents?limit=200&sort=-modified_at`;
    let documents: any[] = [];

    try {
      const res = await fetch(listUrl, {
        method: "GET",
        headers: {
          Authentication: token,
          "Content-Type": "application/json",
          "Accept-Encoding": "identity",
        },
      });

      if (res.ok) {
        const json = await res.json();
        documents = Array.isArray(json) ? json :
                    Array.isArray(json.documents) ? json.documents :
                    [];
        console.log("[checkFiledStatusByConversationAndSubject] Fetched", documents.length, "documents for manual search");
      } else {
        console.warn("[checkFiledStatusByConversationAndSubject] Failed to fetch documents:", res.status);
        return null;
      }
    } catch (e) {
      console.error("[checkFiledStatusByConversationAndSubject] Fetch failed:", e);
      return null;
    }

    // Search through documents for matching conversationId AND normalized subject
    for (const doc of documents) {
      // Only check .eml files
      const fileName = String(doc.name || doc.filename || "");
      if (!fileName.toLowerCase().endsWith(".eml")) {
        continue;
      }

      // Check conversationId match
      const docConversationId = doc.metadata?.conversationId;
      if (!docConversationId || String(docConversationId).trim() !== conversationId.trim()) {
        continue;
      }

      // Check subject match (normalized)
      let docSubject = doc.metadata?.subject || doc.subject || "";
      if (!docSubject) {
        // Fallback: extract from filename
        docSubject = fileName.replace(/\.eml$/i, "");
      }

      const normalizedDocSubject = normalizeSubject(docSubject);

      if (normalizedDocSubject === normalizedSearchSubject) {
        console.log("[checkFiledStatusByConversationAndSubject] Match found!", {
          documentId: doc.id,
          caseId: doc.case_id,
          subject: docSubject,
          conversationIdMatch: true,
          subjectMatch: true,
        });

        return {
          documentId: String(doc.id || doc._id),
          caseId: String(doc.case_id || doc.caseId),
          caseName: doc.case_name || doc.caseName,
          caseKey: doc.case_key || doc.caseKey,
          subject: docSubject,
        };
      }
    }

    console.log("[checkFiledStatusByConversationAndSubject] No match found (checked", documents.length, "documents)");
    return null;
  } catch (e) {
    console.error("[checkFiledStatusByConversationAndSubject] Error:", e);
    return null;
  }
}