import { getStored } from "../utils/storage";
import { STORAGE_KEYS } from "../utils/constants";

function normalizeHost(host: string): string {
  const v = (host || "").trim().toLowerCase();
  if (!v) return "";
  return v.replace(/^https?:\/\//i, "").split("/")[0];
}

async function resolveBaseUrl(): Promise<string> {
  const storedHostRaw = await getStored(STORAGE_KEYS.workspaceHost);
  const host = normalizeHost(storedHostRaw || "");

  if (!host) {
    throw new Error("Workspace host is missing. Please enter a workspace URL first.");
  }

  return `/singlecase/${encodeURIComponent(host)}/publicapi/v1`;
}

export type EmailFiling = {
  document_id: string;
  revision_number?: number;
  case_id: string;
  case_visible_id?: string;
  case_name?: string;
  filed_at?: string;
  filed_by?: string;
};

export type EmailFilingStatusResponse = {
  filed: boolean;
  filings: EmailFiling[];
};

/**
 * Query filing status from SingleCase backend
 * Uses internetMessageId and conversationId to find filed emails
 */
export async function getEmailFilingStatus(
  token: string,
  params: {
    internet_message_id?: string;
    conversation_id?: string;
    include_version_context?: boolean;
  }
): Promise<EmailFilingStatusResponse> {
  if (!token) {
    throw new Error("SingleCase token is missing");
  }

  const baseUrl = await resolveBaseUrl();
  if (!baseUrl) {
    throw new Error("SingleCase base url is missing");
  }

  // Build query string
  const queryParams = new URLSearchParams();
  if (params.internet_message_id) {
    queryParams.set("internet_message_id", params.internet_message_id);
  }
  if (params.conversation_id) {
    queryParams.set("conversation_id", params.conversation_id);
  }
  if (params.include_version_context !== undefined) {
    queryParams.set("include_version_context", String(params.include_version_context));
  }

  const url = `${baseUrl}/email-filings/status?${queryParams.toString()}`;

  console.log("[getEmailFilingStatus] Calling:", url);

  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authentication: token,
      "Content-Type": "application/json",
    },
  });

  const contentType = res.headers.get("content-type") || "";
  const bodyText = await res.text();

  if (!res.ok) {
    throw new Error(`Email filing status check failed (${res.status}): ${bodyText || res.statusText}`);
  }

  if (!contentType.includes("application/json")) {
    throw new Error(
      `Email filing status returned non-JSON (${contentType || "no content-type"}). Body: ${bodyText.slice(0, 200)}`
    );
  }

  return JSON.parse(bodyText) as EmailFilingStatusResponse;
}
