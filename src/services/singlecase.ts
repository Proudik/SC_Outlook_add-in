import { getStored } from "../utils/storage";
import { STORAGE_KEYS } from "../utils/constants";

export type CaseScope = "favourites" | "my" | "all";

export type CaseOption = {
  id: string;
  title: string;

  clientId?: string;
  client?: string;

  status?: string;

  isFavourite?: boolean;
  isMine?: boolean;
};

const MOCK_CASES: CaseOption[] = [
  {
    id: "case_1",
    title: "Contract review - Acme Corp",
    client: "Acme",
    status: "Open",
    isFavourite: true,
    isMine: true,
  },
  {
    id: "case_2",
    title: "Discovery Response - Johnson",
    client: "Johnson",
    status: "In progress",
    isFavourite: true,
    isMine: false,
  },
  {
    id: "case_3",
    title: "Merger - TechVision Inc.",
    client: "TechVision",
    status: "Open",
    isFavourite: false,
    isMine: true,
  },
];

// Dev fallback: calls the localhost proxy path.
// Your webpack devServer proxy currently forwards /singlecase/* to a fixed upstream.

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

function toQuery(params: Record<string, any> = {}) {
  const usp = new URLSearchParams();
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null || v === "") continue;
    if (Array.isArray(v)) usp.set(k, v.join(","));
    else usp.set(k, String(v));
  }
  const q = usp.toString();
  return q ? `?${q}` : "";
}

async function scRequest<T>(
  method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE",
  path: string,
  token: string,
  body?: any,
  params?: Record<string, any>
): Promise<T> {
  if (!token) throw new Error("SingleCase token is missing");

  const baseUrl = await resolveBaseUrl();
  if (!baseUrl) throw new Error("SingleCase base url is missing");

  const url = `${baseUrl}${path}${toQuery(params)}`;

  const res = await fetch(url, {
    method,
    headers: {
      Authentication: token,
      "Content-Type": "application/json",
    },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });

  const contentType = res.headers.get("content-type") || "";
  const bodyText = await res.text();

  if (!res.ok) {
    throw new Error(`SingleCase ${res.status}: ${bodyText || res.statusText}`);
  }

  if (!contentType.includes("application/json")) {
    if (!bodyText) return undefined as any;
    throw new Error(
      `SingleCase returned non JSON (${contentType || "no content-type"}). First bytes: ${bodyText.slice(0, 200)}`
    );
  }

  return JSON.parse(bodyText) as T;
}

async function scGet<T>(path: string, token: string, params?: Record<string, any>): Promise<T> {
  return scRequest<T>("GET", path, token, undefined, params);
}

function mapApiCaseToOption(apiCase: any): CaseOption {
  const clientId =
    apiCase.client_id != null
      ? String(apiCase.client_id)
      : apiCase.clientId != null
        ? String(apiCase.clientId)
        : apiCase.client?.id != null
          ? String(apiCase.client.id)
          : undefined;

  const visible = apiCase.case_id_visible ? String(apiCase.case_id_visible) : "";
  const name = apiCase.name ? String(apiCase.name) : "";
  const title = visible ? `${visible} · ${name || `Case ${apiCase.id}`}` : name || `Case ${apiCase.id}`;

  return {
    id: String(apiCase.id),
    title,
    clientId,
    client: apiCase.client?.name || apiCase.client_name || undefined,
    status: apiCase.status?.name || apiCase.status || undefined,
  };
}

export async function listCases(token: string, scope: CaseScope = "my"): Promise<CaseOption[]> {
  void scope; // until you implement scope filtering
  const raw = await scGet<any>("/cases", token);
  const arr = Array.isArray(raw) ? raw : raw.items ?? raw.data ?? [];
  return arr.map(mapApiCaseToOption);
}

export type ClientOption = { id: string; name: string };

function mapApiClientToOption(apiClient: any): ClientOption {
  return {
    id: String(apiClient.id),
    name: String(apiClient.name || ""),
  };
}

export async function listClients(token: string): Promise<ClientOption[]> {
  const raw = await scGet<any>("/clients", token);
  const arr = Array.isArray(raw) ? raw : raw.items ?? raw.data ?? [];
  return arr.map(mapApiClientToOption).filter((c) => c.id && c.name);
}

export type AttachmentPayload = {
  id: string;
  name: string;
  contentType?: string;
  size: number;
  contentBase64: string;
};

export type SubmitEmailToCasePayload = {
  caseId: string;
  outlookItemId: string;
  subject: string;
  fromEmail: string;
  fromName: string;
  bodySnippet?: string;
  attachments?: AttachmentPayload[];
};

// Still mocked until you implement the real endpoint
export async function submitEmailToCase(token: string, payload: SubmitEmailToCasePayload) {
  void token;
  await new Promise((r) => setTimeout(r, 500));

  return {
    ok: true,
    singlecaseRecordId: `mail_${Date.now()}`,
    payload,
  };
}



/* -------------------------------------------------------------------------- */
/* Users                                                                       */
/* -------------------------------------------------------------------------- */

export type UserOption = {
  id: string;
  firstName: string;
  lastName: string;
  username: string;
};

function mapApiUserToOption(apiUser: any): UserOption {
  return {
    id: String(apiUser.id),
    firstName: String(apiUser.first_name || ""),
    lastName: String(apiUser.last_name || ""),
    username: String(apiUser.username || ""),
  };
}

export async function listUsers(token: string): Promise<UserOption[]> {
  const raw = await scRequest<any>("GET", "/users", token);
  const arr = Array.isArray(raw) ? raw : raw.items ?? raw.data ?? [];
  return arr.map(mapApiUserToOption);
}

export async function getUser(token: string, userId: string): Promise<UserOption> {
  const raw = await scRequest<any>("GET", `/users/${encodeURIComponent(userId)}`, token);
  return mapApiUserToOption(raw);
}


/* -------------------------------------------------------------------------- */
/* Timesheets (Timers)                                                         */
/* -------------------------------------------------------------------------- */

export type TimerItem = {
  id?: string;
  user_id: string;
  project_id: string;
  date: string;
  total_time: number;
  total_billed_time: number;
  sheet_activity_id?: string;
  note?: string;
};

export type ListTimersFilter = {
  userId?: string;
  projectId?: string;
  from?: string; // YYYY-MM-DD
  to?: string; // YYYY-MM-DD
};

function toTimersQuery(filter: ListTimersFilter) {
  return {
    "filter[user_id]": filter.userId,
    "filter[project_id]": filter.projectId,
    "filter[from]": filter.from,
    "filter[to]": filter.to,
  };
}

function normaliseTimer(raw: any): TimerItem {
  if (!raw) return raw;

  // Some APIs include timer_id, others use id
  const id =
    raw.id != null
      ? String(raw.id)
      : raw.timer_id != null
        ? String(raw.timer_id)
        : raw.timerId != null
          ? String(raw.timerId)
          : undefined;

  return {
    id,
    user_id: raw.user_id != null ? String(raw.user_id) : String(raw.userId || ""),
    project_id: raw.project_id != null ? String(raw.project_id) : String(raw.projectId || ""),
    date: String(raw.date || ""),
    total_time: Number(raw.total_time || 0),
    total_billed_time: Number(raw.total_billed_time || 0),
    sheet_activity_id:
      raw.sheet_activity_id != null
        ? String(raw.sheet_activity_id)
        : raw.sheetActivityId != null
          ? String(raw.sheetActivityId)
          : undefined,
    note: raw.note != null ? String(raw.note) : "",
  };
}

export async function listTimers(token: string, filter: ListTimersFilter): Promise<TimerItem[]> {
  const raw = await scRequest<any>("GET", "/timers", token, undefined, toTimersQuery(filter));
  const arr = Array.isArray(raw) ? raw : raw.items ?? raw.data ?? [];
  return arr.map(normaliseTimer);
}

export async function getTimer(token: string, timerId: string): Promise<TimerItem> {
  const raw = await scRequest<any>("GET", `/timers/${encodeURIComponent(timerId)}`, token);
  return normaliseTimer(raw);
}

export type UpsertTimerPayload = {
  user_id: number;
  project_id: number;
  date: string; // "YYYY-MM-DD 00:00:00"
  total_time: number; // seconds
  total_billed_time: number; // seconds
  sheet_activity_id: number;
  note: string;
};

export async function createTimer(token: string, payload: UpsertTimerPayload): Promise<any> {
  return scRequest<any>("POST", "/timers", token, payload);
}

export async function updateTimer(token: string, timerId: string, payload: UpsertTimerPayload): Promise<any> {
  return scRequest<any>("PUT", `/timers/${encodeURIComponent(timerId)}`, token, payload);
}
