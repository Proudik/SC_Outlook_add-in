import { getAuth } from "./auth";
import { getStored } from "../utils/storage";
import { STORAGE_KEYS } from "../utils/constants";

function normalizeHost(host: string): string {
  const v = (host || "").trim().toLowerCase();
  if (!v) return "";
  return v.replace(/^https?:\/\//i, "").split("/")[0];
}

function toQuery(params: Record<string, any> = {}) {
  const usp = new URLSearchParams();
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null || v === "") continue;
    usp.set(k, String(v));
  }
  const q = usp.toString();
  return q ? `?${q}` : "";
}

export async function scGet<T>(
  path: string,
  params?: Record<string, any>,
  signal?: AbortSignal
): Promise<T> {
  const { token } = getAuth();
  console.log("[scGet] token first 8:", (token || "").slice(0, 8));

  if (!token) {
    throw new Error("SingleCase token is missing. Please sign in again.");
  }

  const storedHostRaw = await getStored(STORAGE_KEYS.workspaceHost);
  const host = normalizeHost(storedHostRaw || "");
  if (!host) {
    throw new Error("Workspace host is missing. Please select a workspace again.");
  }

  const baseUrl = `/singlecase/${encodeURIComponent(host)}/publicapi/v1`;
  const url = `${baseUrl}${path}${toQuery(params)}`;

  const res = await fetch(url, {
    method: "GET",
    signal,
    headers: {
      Authentication: token,
      "Content-Type": "application/json",
    },
  });

  const contentType = res.headers.get("content-type") || "";
  const bodyText = await res.text().catch(() => "");

  if (!res.ok) {
    throw new Error(`SingleCase ${res.status}: ${bodyText || res.statusText}`);
  }

  if (!contentType.includes("application/json")) {
    throw new Error(
      `SingleCase returned non JSON (${contentType || "no content-type"}). First bytes: ${bodyText.slice(0, 200)}`
    );
  }

  return JSON.parse(bodyText) as T;
}
