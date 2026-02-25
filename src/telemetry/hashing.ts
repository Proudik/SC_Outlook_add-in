// ─── PII-safe hashing utilities ──────────────────────────────────────────────
// All identifiers (user email, item IDs, case IDs) must be hashed before they
// are placed into any telemetry payload. Raw values must never reach the wire.

/**
 * SHA-256 hex string using SubtleCrypto (available in modern Outlook webview).
 * Falls back to a deterministic djb2 hash for environments without SubtleCrypto
 * (e.g. IE11 legacy mode) — still prevents raw PII from being logged.
 */
export async function sha256Hex(input: string): Promise<string> {
  try {
    const encoder = new TextEncoder();
    const data = encoder.encode(input);
    const buf = await crypto.subtle.digest("SHA-256", data);
    return Array.from(new Uint8Array(buf))
      .map((b) => b.toString(16).padStart(2, "0"))
      .join("");
  } catch {
    // Fallback: djb2 — not cryptographic but avoids raw PII in logs
    let h = 5381;
    for (let i = 0; i < input.length; i++) {
      h = ((h << 5) + h) ^ input.charCodeAt(i);
      h = h >>> 0;
    }
    return "h-" + h.toString(16).padStart(8, "0");
  }
}

/**
 * Derives a stable, anonymous user ID.
 * Never logs the raw email or workspace host.
 */
export async function deriveAnonymousUserId(
  email: string,
  workspaceHost: string
): Promise<string> {
  const normalized =
    (email || "").toLowerCase().trim() +
    ":" +
    (workspaceHost || "").toLowerCase().trim();
  return sha256Hex("user:" + normalized);
}

/** SHA-256 of workspaceHost — safe to store in CloudWatch. */
export async function hashWorkspaceId(workspaceHost: string): Promise<string> {
  return sha256Hex("ws:" + (workspaceHost || "").toLowerCase().trim());
}

/** SHA-256 of an Outlook item ID. Pass "" to get back "". */
export async function hashItemId(itemId: string): Promise<string> {
  if (!itemId) return "";
  return sha256Hex("item:" + itemId);
}

/** SHA-256 of a SingleCase case ID. Pass "" to get back "". */
export async function hashCaseId(caseId: string): Promise<string> {
  if (!caseId) return "";
  return sha256Hex("case:" + caseId);
}

/**
 * Strips obvious PII patterns from error messages before they reach telemetry:
 *  • email addresses  → [email]
 *  • URLs             → [url]
 *  • long opaque IDs  → [id]
 * Truncated to 300 chars to bound payload size.
 */
export function sanitizeErrorMessage(msg: string): string {
  return (msg || "")
    .replace(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g, "[email]")
    .replace(/https?:\/\/[^\s"'<>]+/g, "[url]")
    .replace(/[A-Za-z0-9_\-]{20,}/g, "[id]")
    .slice(0, 300);
}
