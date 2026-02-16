/* global OfficeRuntime */

const TOKEN_KEY = "singlecase_token";
const USER_KEY = "singlecase_user_email";
const ISSUED_AT_KEY = "singlecase_auth_issued_at";

// Mirror keys into OfficeRuntime.storage so Commands runtime can read them
const RT_TOKEN_KEY = "sc_token";
const RT_USER_KEY = "sc_user_email";
const RT_ISSUED_AT_KEY = "sc_auth_issued_at";

// Typical session TTL: 8 hours. Adjust as you like.
const SESSION_TTL_MS = 8 * 60 * 60 * 1000;

function normalizeEmail(email: string | null | undefined): string {
  const v = (email || "").trim().toLowerCase();
  return v.length > 0 ? v : "unknown@singlecase.local";
}

async function rtGet(key: string): Promise<string | null> {
  // Try OfficeRuntime.storage first
  if (typeof OfficeRuntime !== 'undefined' && (OfficeRuntime as any)?.storage) {
    try {
      const v = await (OfficeRuntime as any).storage.getItem(key);
      if (typeof v === "string") return v;
    } catch (e) {
      console.warn("[rtGet] OfficeRuntime.storage.getItem failed:", e);
    }
  }

  // Fallback to Office.context.roamingSettings (Outlook-specific, works cross-context)
  if ((Office as any)?.context?.roamingSettings) {
    try {
      const v = (Office as any).context.roamingSettings.get(key);
      if (typeof v === "string") return v;
    } catch (e) {
      console.warn("[rtGet] roamingSettings.get failed:", e);
    }
  }

  return null;
}

async function rtSet(key: string, value: string): Promise<void> {
  // Try OfficeRuntime.storage first
  if (typeof OfficeRuntime !== 'undefined' && (OfficeRuntime as any)?.storage) {
    try {
      await (OfficeRuntime as any).storage.setItem(key, value);
      console.log("[rtSet] Saved to OfficeRuntime.storage:", key);
      return; // Success
    } catch (e) {
      console.warn("[rtSet] OfficeRuntime.storage.setItem failed:", e);
    }
  }

  // Fallback to Office.context.roamingSettings (Outlook-specific, works cross-context)
  if ((Office as any)?.context?.roamingSettings) {
    try {
      (Office as any).context.roamingSettings.set(key, value);
      await new Promise<void>((resolve, reject) => {
        (Office as any).context.roamingSettings.saveAsync((result: any) => {
          if (result.status === (Office as any).AsyncResultStatus.Succeeded) {
            console.log("[rtSet] Saved to roamingSettings:", key);
            resolve();
          } else {
            console.error("[rtSet] roamingSettings.saveAsync failed:", result.error);
            reject(new Error(result.error?.message || "saveAsync failed"));
          }
        });
      });
    } catch (e) {
      console.error("[rtSet] roamingSettings failed:", e);
    }
  }
}

async function rtRemove(key: string): Promise<void> {
  try {
    await (OfficeRuntime as any).storage.removeItem(key);
  } catch {
    // ignore
  }
}

export function getAuth(): { token: string | null; email: string; issuedAt: number } {
  const token = sessionStorage.getItem(TOKEN_KEY);
  const emailRaw = sessionStorage.getItem(USER_KEY);
  const issuedAtStr = sessionStorage.getItem(ISSUED_AT_KEY);

  return {
    token,
    email: normalizeEmail(emailRaw),
    issuedAt: issuedAtStr ? Number(issuedAtStr) : 0,
  };
}

// Async version for runtimes that cannot access sessionStorage (eg Commands)
export async function getAuthRuntime(): Promise<{ token: string | null; email: string; issuedAt: number }> {
  const [token, emailRaw, issuedAtStr] = await Promise.all([
    rtGet(RT_TOKEN_KEY),
    rtGet(RT_USER_KEY),
    rtGet(RT_ISSUED_AT_KEY),
  ]);

  return {
    token,
    email: normalizeEmail(emailRaw),
    issuedAt: issuedAtStr ? Number(issuedAtStr) : 0,
  };
}

// Make this async so you can await the mirror write when needed.
export async function setAuth(token: string, email: string): Promise<void> {
  const emailNorm = normalizeEmail(email);
  const issuedAt = Date.now();

  sessionStorage.setItem(TOKEN_KEY, token);
  sessionStorage.setItem(USER_KEY, emailNorm);
  sessionStorage.setItem(ISSUED_AT_KEY, String(issuedAt));

  // Mirror for command runtime
  await Promise.all([
    rtSet(RT_TOKEN_KEY, token),
    rtSet(RT_USER_KEY, emailNorm),
    rtSet(RT_ISSUED_AT_KEY, String(issuedAt)),
  ]);
}

export function clearAuthIfExpired(): void {
  const { token, issuedAt } = getAuth();
  if (!token) return;

  const ageMs = Date.now() - (issuedAt || 0);
  if (!issuedAt || ageMs > SESSION_TTL_MS) {
    void clearAuth();
  }
}

export async function clearAuthIfExpiredRuntime(): Promise<void> {
  const { token, issuedAt } = await getAuthRuntime();
  if (!token) return;

  const ageMs = Date.now() - (issuedAt || 0);
  if (!issuedAt || ageMs > SESSION_TTL_MS) {
    await clearAuth();
  }
}

export async function clearAuth(): Promise<void> {
  sessionStorage.removeItem(TOKEN_KEY);
  sessionStorage.removeItem(USER_KEY);
  sessionStorage.removeItem(ISSUED_AT_KEY);

  await Promise.all([rtRemove(RT_TOKEN_KEY), rtRemove(RT_USER_KEY), rtRemove(RT_ISSUED_AT_KEY)]);
}

export function isLoggedIn(): boolean {
  const { token, issuedAt } = getAuth();
  if (!token) return false;

  const ageMs = Date.now() - (issuedAt || 0);
  return Boolean(issuedAt && ageMs <= SESSION_TTL_MS);
}

export async function isLoggedInRuntime(): Promise<boolean> {
  const { token, issuedAt } = await getAuthRuntime();
  if (!token) return false;

  const ageMs = Date.now() - (issuedAt || 0);
  return Boolean(issuedAt && ageMs <= SESSION_TTL_MS);
}