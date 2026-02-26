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

/**
 * Emergency in-memory pruning when roamingSettings hits the 32KB limit.
 * Removes old individual sc:sent:* / sc_conv_ctx:* keys (via internal _data access),
 * prunes the sentPills blob and the filed-email cache to 3 entries each.
 * Does NOT call saveAsync — caller must do that after this returns.
 */
function emergencyPruneRoamingSettings(): void {
  // 1. Remove debug log
  try { Office.context.roamingSettings.remove(DEBUG_LOG_KEY); } catch { /* ignore */ }

  // 2. Try to enumerate all in-memory keys via internal _data and remove legacy per-email entries.
  //    roamingSettings._data is not part of the public API but is the underlying store in all
  //    known Office.js web/desktop builds. This cleans up old sc:sent:* keys accumulated before
  //    the blob-based sentPillStore was introduced.
  try {
    const settings = Office.context.roamingSettings as any;
    const dataStore: Record<string, unknown> | null =
      settings._data ?? settings._settings?.data ?? null;
    if (dataStore && typeof dataStore === "object") {
      const allKeys = Object.keys(dataStore);
      let removedCount = 0;
      for (const k of allKeys) {
        if (k.startsWith("sc:sent:") || k.startsWith("sc_conv_ctx:") || k.startsWith("sc:uploadedLinks:") || k === DEBUG_LOG_KEY) {
          try { Office.context.roamingSettings.remove(k); removedCount++; } catch { /* ignore */ }
        }
      }
      if (removedCount > 0) {
        console.warn("[emergencyPrune] Removed", removedCount, "legacy individual storage entries");
      }
    }
  } catch { /* ignore if internal API not available */ }

  // 3. Prune sc:sentPills blob to 3 most recent
  try {
    const raw = Office.context.roamingSettings.get("sc:sentPills");
    if (raw) {
      const blob: Record<string, any> = JSON.parse(String(raw));
      const entries = Object.entries(blob) as [string, any][];
      if (entries.length > 3) {
        entries.sort((a, b) => (b[1]._savedAt || 0) - (a[1]._savedAt || 0));
        const pruned: Record<string, any> = {};
        entries.slice(0, 3).forEach(([k, v]) => { pruned[k] = v; });
        Office.context.roamingSettings.set("sc:sentPills", JSON.stringify(pruned));
        console.warn("[emergencyPrune] Pruned sentPills from", entries.length, "to 3 entries");
      }
    }
  } catch { /* ignore */ }

  // 4. Prune sc:filedEmailsCache to 3 most recent
  try {
    const raw = Office.context.roamingSettings.get("sc:filedEmailsCache");
    if (raw) {
      const cache: Record<string, any> = JSON.parse(String(raw));
      const entries = Object.entries(cache) as [string, any][];
      if (entries.length > 3) {
        entries.sort((a, b) => (b[1].filedAt || 0) - (a[1].filedAt || 0));
        const pruned: Record<string, any> = {};
        entries.slice(0, 3).forEach(([k, v]) => { pruned[k] = v; });
        Office.context.roamingSettings.set("sc:filedEmailsCache", JSON.stringify(pruned));
        console.warn("[emergencyPrune] Pruned filedCache from", entries.length, "to 3 entries");
      }
    }
  } catch { /* ignore */ }
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

  const lsGet = (): string | null => {
    try { return typeof localStorage !== "undefined" ? localStorage.getItem(k) : null; } catch { return null; }
  };

  try {
    if (hasOfficeRuntimeStorage()) {
      const v = await (OfficeRuntime as any).storage.getItem(k);
      if (shouldLog) {
        console.log("[getStored] Got from OfficeRuntime.storage:", k, v ? `found (${v.length} chars)` : "not found");
      }
      if (typeof v === "string") return v;
      // setStored may have fallen back to localStorage (e.g. if OfficeRuntime.storage threw).
      // Also check roamingSettings in case another client wrote there (cross-context sync).
      if (hasRoamingSettings()) {
        const rv = Office.context.roamingSettings.get(k);
        if (typeof rv === "string") return rv;
      }
      return lsGet();
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
      if (typeof v === "string") return v;
      // setStored may have fallen back to localStorage on saveAsync failure.
      return lsGet();
    }

    const v = localStorage.getItem(k);
    if (shouldLog) {
      console.warn("[getStored] No Office storage, using localStorage:", k, v ? `found (${v.length} chars)` : "not found");
    }
    return v;
  } catch (e) {
    console.warn("[getStored] Failed, falling back to localStorage:", e);
    return lsGet();
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

  // Always mirror to localStorage as belt-and-suspenders backup.
  // getStored will fall back to localStorage if the primary backend returns null
  // (e.g. cross-context sync delay between Desktop and OWA, or silent write failure).
  const lsSet = () => { try { if (typeof localStorage !== "undefined") localStorage.setItem(k, v); } catch { /* ignore */ } };

  try {
    if (hasOfficeRuntimeStorage()) {
      if (VERBOSE_LOGGING) console.log("[setStored] Writing to OfficeRuntime.storage...");
      await (OfficeRuntime as any).storage.setItem(k, v);
      lsSet();
      if (VERBOSE_LOGGING) console.log("[setStored] ✅ Write to OfficeRuntime.storage completed");
      return;
    }

    if (hasRoamingSettings()) {
      if (VERBOSE_LOGGING) console.log("[setStored] Writing to roamingSettings...");
      Office.context.roamingSettings.set(k, v);
      if (VERBOSE_LOGGING) console.log("[setStored] Calling saveAsync...");

      try {
        await saveRoamingSettings();
        lsSet();
        if (VERBOSE_LOGGING) console.log("[setStored] ✅ saveAsync completed");
        return;
      } catch (saveError) {
        // If roamingSettings exceeded 32KB, try emergency prune once then silently fall back to localStorage.
        // This is expected in OWA where old accumulated keys fill the 32KB limit.
        // All callers already mirror to localStorage, so data is not lost.
        const isOverflow = (saveError as any)?.message?.includes("32 KB") ||
                           (saveError as any)?.message?.includes("size limit");
        if (isOverflow) {
          if (retryCount === 0) {
            emergencyPruneRoamingSettings();
            try {
              await saveRoamingSettings();
              lsSet();
              if (VERBOSE_LOGGING) console.log("[setStored] ✅ saveAsync succeeded after emergency prune");
              return;
            } catch {
              // Still full — fall through to localStorage silently
            }
          }
          lsSet();
          return;
        }

        console.error("[setStored] saveAsync failed:", saveError);

        // Retry on desktop Outlook if save fails for other (transient) reasons
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