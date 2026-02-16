// src/utils/storage.ts
/* global Office, OfficeRuntime */

// Feature flag: Set to false to silence verbose logging (helps with render loops)
const VERBOSE_LOGGING = false;

// Debug log that persists across sessions
const DEBUG_LOG_KEY = "sc:debugLog";

export async function getDebugLog(): Promise<string> {
  try {
    if (hasRoamingSettings()) {
      return String(Office.context.roamingSettings.get(DEBUG_LOG_KEY) || "");
    } else if (typeof localStorage !== "undefined") {
      return localStorage.getItem(DEBUG_LOG_KEY) || "";
    }
    return "";
  } catch {
    return "";
  }
}

export async function clearDebugLog(): Promise<void> {
  try {
    if (hasRoamingSettings()) {
      Office.context.roamingSettings.remove(DEBUG_LOG_KEY);
      await saveRoamingSettings();
    } else if (typeof localStorage !== "undefined") {
      localStorage.removeItem(DEBUG_LOG_KEY);
    }
  } catch {
    // Silent fail
  }
}

function hasOfficeRuntimeStorage(): boolean {
  try {
    return typeof OfficeRuntime !== "undefined" && !!(OfficeRuntime as any)?.storage;
  } catch {
    return false;
  }
}

function hasRoamingSettings(): boolean {
  try {
    return !!Office?.context?.roamingSettings;
  } catch {
    return false;
  }
}

async function saveRoamingSettings(): Promise<void> {
  const startTime = Date.now();
  if (VERBOSE_LOGGING) console.log("[saveRoamingSettings] Starting saveAsync...");

  await new Promise<void>((resolve, reject) => {
    try {
      Office.context.roamingSettings.saveAsync((res: any) => {
        const duration = Date.now() - startTime;
        if (res?.status === Office.AsyncResultStatus.Succeeded) {
          if (VERBOSE_LOGGING) console.log(`[saveRoamingSettings] ✅ Succeeded in ${duration}ms`);
          resolve();
        } else {
          const errorMsg = res?.error?.message || "roamingSettings.saveAsync failed";
          console.error(`[saveRoamingSettings] ❌ Failed in ${duration}ms:`, errorMsg, {
            status: res?.status,
            errorCode: res?.error?.code,
            errorName: res?.error?.name,
          });
          reject(new Error(errorMsg));
        }
      });
    } catch (e) {
      const duration = Date.now() - startTime;
      console.error(`[saveRoamingSettings] ❌ Exception in ${duration}ms:`, e);
      reject(e);
    }
  });

  // CRITICAL FOR DESKTOP OUTLOOK: Add small delay to ensure operation completes
  // Desktop Outlook may close compose window immediately after send,
  // interrupting async operations. This delay ensures saveAsync completes.
  await new Promise(resolve => setTimeout(resolve, 100));
}

export async function getStored(key: string, forceFresh = false): Promise<string | null> {
  const k = String(key || "").trim();
  if (!k) return null;

  // Detect which storage backend is available
  const storageBackend = hasOfficeRuntimeStorage()
    ? "OfficeRuntime.storage"
    : hasRoamingSettings()
    ? "roamingSettings"
    : "localStorage";

  // Only log for important keys to avoid console spam
  const shouldLog = VERBOSE_LOGGING && !k.includes("recipientHistory") && !k.includes("recentCases");

  if (shouldLog) {
    console.log("[getStored] Using storage backend:", storageBackend, "for key:", k, forceFresh ? "(force fresh)" : "");
  }

  try {
    if (hasOfficeRuntimeStorage()) {
      const v = await (OfficeRuntime as any).storage.getItem(k);
      if (shouldLog) {
        console.log("[getStored] Got from OfficeRuntime.storage:", k, v ? `found (${v.length} chars)` : "not found");
      }
      return typeof v === "string" ? v : null;
    }

    if (hasRoamingSettings()) {
      // WORKAROUND: roamingSettings can be stale after another instance wrote data
      // Force a small delay to allow server sync when forceFresh is requested
      if (forceFresh && VERBOSE_LOGGING) {
        console.log("[getStored] Waiting 500ms for roamingSettings sync...");
        await new Promise(resolve => setTimeout(resolve, 500));
      }

      const v = Office.context.roamingSettings.get(k);
      if (shouldLog) {
        console.log("[getStored] Got from roamingSettings:", k, v ? `found (${String(v).length} chars)` : "not found");
      }
      return typeof v === "string" ? v : null;
    }

    const v = localStorage.getItem(k);
    if (shouldLog) {
      console.warn("[getStored] No Office storage, using localStorage:", k, v ? `found (${v.length} chars)` : "not found");
    }
    return v;
  } catch (e) {
    console.warn("[getStored] Failed, falling back to localStorage:", e);
    return localStorage.getItem(k);
  }
}

export async function setStored(key: string, value: string, retryCount = 0): Promise<void> {
  const k = String(key || "").trim();
  if (!k) return;

  const v = String(value ?? "");
  const MAX_RETRIES = 2;

  // Detect which storage backend is available
  const storageBackend = hasOfficeRuntimeStorage()
    ? "OfficeRuntime.storage"
    : hasRoamingSettings()
    ? "roamingSettings"
    : "localStorage";

  if (VERBOSE_LOGGING) {
    console.log("[setStored] Using storage backend:", storageBackend, "for key:", k, `(${v.length} chars)`, retryCount > 0 ? `[retry ${retryCount}]` : "");
  }

  try {
    if (hasOfficeRuntimeStorage()) {
      if (VERBOSE_LOGGING) console.log("[setStored] Writing to OfficeRuntime.storage...");
      await (OfficeRuntime as any).storage.setItem(k, v);
      if (VERBOSE_LOGGING) console.log("[setStored] ✅ Write to OfficeRuntime.storage completed");
      return;
    }

    if (hasRoamingSettings()) {
      if (VERBOSE_LOGGING) console.log("[setStored] Writing to roamingSettings...");
      Office.context.roamingSettings.set(k, v);
      if (VERBOSE_LOGGING) console.log("[setStored] Calling saveAsync...");

      try {
        await saveRoamingSettings();
        if (VERBOSE_LOGGING) console.log("[setStored] ✅ saveAsync completed");
        return;
      } catch (saveError) {
        console.error("[setStored] saveAsync failed:", saveError);

        // Retry on desktop Outlook if save fails
        if (retryCount < MAX_RETRIES) {
          const delay = 200 * (retryCount + 1); // Exponential backoff: 200ms, 400ms
          if (VERBOSE_LOGGING) console.log(`[setStored] Retrying in ${delay}ms...`);
          await new Promise(resolve => setTimeout(resolve, delay));
          return setStored(key, value, retryCount + 1);
        }

        throw saveError;
      }
    }

    if (VERBOSE_LOGGING) console.warn("[setStored] No Office storage, using localStorage for key:", k);
    localStorage.setItem(k, v);
  } catch (e) {
    console.warn("[setStored] ❌ Failed after retries, falling back to localStorage:", e);
    localStorage.setItem(k, v);
  }
}

export async function removeStored(key: string): Promise<void> {
  const k = String(key || "").trim();
  if (!k) return;

  try {
    if (hasOfficeRuntimeStorage()) {
      await (OfficeRuntime as any).storage.removeItem(k);
      return;
    }

    if (hasRoamingSettings()) {
      Office.context.roamingSettings.remove(k);
      await saveRoamingSettings();
      return;
    }

    localStorage.removeItem(k);
  } catch {
    localStorage.removeItem(k);
  }
}