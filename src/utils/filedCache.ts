import { getStored, setStored } from "./storage";

const FILED_CACHE_KEY = "sc:filedEmailsCache";

export type FiledEmailCache = {
  [conversationId: string]: {
    caseId: string;
    documentId: string;
    subject: string;
    caseName?: string;
    caseKey?: string;
    filedAt: number; // timestamp
  };
};

/**
 * Store filed email info by conversationId
 * This enables "already filed" detection for self-sent emails and replies
 *
 * Works for:
 * - Self-sent emails (sender opens received copy)
 * - Sent items (user reopens their own sent email)
 * - Replies in same thread (same conversationId)
 */
export async function cacheFiledEmail(
  conversationId: string,
  caseId: string,
  documentId: string,
  subject: string,
  caseName?: string,
  caseKey?: string
): Promise<void> {
  if (!conversationId) {
    console.warn("[cacheFiledEmail] No conversationId provided, skipping cache");
    return;
  }

  try {
    // Platform detection
    const platform = {
      host: (Office as any)?.context?.mailbox?.diagnostics?.hostName,
      hostVersion: (Office as any)?.context?.mailbox?.diagnostics?.hostVersion,
      platform: (Office as any)?.context?.platform,
    };
    console.log("[cacheFiledEmail] Platform info:", platform);

    const raw = await getStored(FILED_CACHE_KEY);
    const cache: FiledEmailCache = raw ? JSON.parse(String(raw)) : {};
    console.log("[cacheFiledEmail] Current cache size:", Object.keys(cache).length);

    cache[conversationId] = {
      caseId,
      documentId,
      subject,
      caseName,
      caseKey,
      filedAt: Date.now(),
    };

    // Clean old entries (keep last 20 filed emails - reduced to stay under 32KB limit)
    const entries = Object.entries(cache);
    if (entries.length > 20) {
      entries.sort((a, b) => b[1].filedAt - a[1].filedAt);
      const keep = entries.slice(0, 20);
      const newCache: FiledEmailCache = {};
      keep.forEach(([key, val]) => {
        newCache[key] = val;
      });
      await setStored(FILED_CACHE_KEY, JSON.stringify(newCache));
      console.log("[cacheFiledEmail] Cleaned cache, kept 20 most recent entries");
    } else {
      await setStored(FILED_CACHE_KEY, JSON.stringify(cache));
    }

    // Verify write succeeded
    const verification = await getStored(FILED_CACHE_KEY);
    const verifiedCache = verification ? JSON.parse(String(verification)) : {};
    const writeSuccess = !!verifiedCache[conversationId];
    console.log("[cacheFiledEmail] Write verification:", {
      success: writeSuccess,
      cacheSize: Object.keys(verifiedCache).length,
    });

    console.log("[cacheFiledEmail] Cached filed email", {
      conversationId: conversationId.substring(0, 20) + "...",
      caseId,
      documentId,
      subject,
      writeVerified: writeSuccess,
    });
  } catch (e) {
    console.warn("[cacheFiledEmail] Failed to cache:", e);
    // Non-critical, don't throw
  }
}

/**
 * Check if email with this conversationId was filed
 * Returns cached info if found, null otherwise
 */
export async function getFiledEmailFromCache(
  conversationId: string
): Promise<FiledEmailCache[string] | null> {
  if (!conversationId) {
    return null;
  }

  try {
    // Platform detection
    const platform = {
      host: (Office as any)?.context?.mailbox?.diagnostics?.hostName,
      hostVersion: (Office as any)?.context?.mailbox?.diagnostics?.hostVersion,
      platform: (Office as any)?.context?.platform,
    };
    console.log("[getFiledEmailFromCache] Platform info:", platform);

    // Force fresh read to avoid stale roamingSettings after recent write
    const raw = await getStored(FILED_CACHE_KEY, true);
    if (!raw) {
      console.log("[getFiledEmailFromCache] No cache found in storage");
      return null;
    }

    const cache: FiledEmailCache = JSON.parse(String(raw));
    const cacheKeys = Object.keys(cache);
    console.log("[getFiledEmailFromCache] Cache size:", cacheKeys.length, "keys");
    console.log("[getFiledEmailFromCache] Looking for conversationId:", conversationId.substring(0, 30) + "...");
    console.log("[getFiledEmailFromCache] Sample cache keys:", cacheKeys.slice(0, 3).map(k => k.substring(0, 30) + "..."));

    const entry = cache[conversationId];

    if (entry) {
      console.log("[getFiledEmailFromCache] ✅ Found cache entry", {
        conversationId: conversationId.substring(0, 20) + "...",
        caseId: entry.caseId,
        documentId: entry.documentId,
        filedAt: new Date(entry.filedAt).toISOString(),
        subject: entry.subject,
      });
      return entry;
    }

    console.log("[getFiledEmailFromCache] ❌ No entry for this conversationId");
    return null;
  } catch (e) {
    console.warn("[getFiledEmailFromCache] Failed to read cache:", e);
    return null;
  }
}

/**
 * Cache filed email by subject (fallback when conversationId not available at send time)
 * Used for NEW compose emails where conversationId isn't assigned until after send
 */
export async function cacheFiledEmailBySubject(
  subject: string,
  caseId: string,
  documentId: string,
  caseName?: string,
  caseKey?: string
): Promise<void> {
  if (!subject) {
    console.warn("[cacheFiledEmailBySubject] No subject provided, skipping cache");
    return;
  }

  try {
    // Platform detection
    const platform = {
      host: (Office as any)?.context?.mailbox?.diagnostics?.hostName,
      hostVersion: (Office as any)?.context?.mailbox?.diagnostics?.hostVersion,
      platform: (Office as any)?.context?.platform,
    };
    console.log("[cacheFiledEmailBySubject] Platform info:", platform);

    const raw = await getStored(FILED_CACHE_KEY);
    const cache: FiledEmailCache = raw ? JSON.parse(String(raw)) : {};
    console.log("[cacheFiledEmailBySubject] Current cache size:", Object.keys(cache).length);

    // Use subject as temporary key (prefixed with "subj:")
    const tempKey = `subj:${subject.trim().toLowerCase()}`;
    console.log("[cacheFiledEmailBySubject] Using temp key:", tempKey);

    cache[tempKey] = {
      caseId,
      documentId,
      subject,
      caseName,
      caseKey,
      filedAt: Date.now(),
    };

    // Clean old entries (keep last 20 - reduced to stay under 32KB limit)
    const entries = Object.entries(cache);
    if (entries.length > 20) {
      entries.sort((a, b) => b[1].filedAt - a[1].filedAt);
      const keep = entries.slice(0, 20);
      const newCache: FiledEmailCache = {};
      keep.forEach(([key, val]) => {
        newCache[key] = val;
      });
      await setStored(FILED_CACHE_KEY, JSON.stringify(newCache));
      console.log("[cacheFiledEmailBySubject] Cleaned cache, kept 20 most recent entries");
    } else {
      await setStored(FILED_CACHE_KEY, JSON.stringify(cache));
    }

    // Verify write succeeded
    const verification = await getStored(FILED_CACHE_KEY);
    const verifiedCache = verification ? JSON.parse(String(verification)) : {};
    const writeSuccess = !!verifiedCache[tempKey];
    console.log("[cacheFiledEmailBySubject] Write verification:", {
      success: writeSuccess,
      cacheSize: Object.keys(verifiedCache).length,
      tempKey,
    });

    console.log("[cacheFiledEmailBySubject] Cached filed email by subject", {
      subject,
      caseId,
      documentId,
      writeVerified: writeSuccess,
    });
  } catch (e) {
    console.warn("[cacheFiledEmailBySubject] Failed to cache:", e);
  }
}

/**
 * Search cache by subject (fallback when conversationId lookup fails)
 * Also upgrades the cache entry to use conversationId for future lookups
 */
export async function findFiledEmailBySubject(
  subject: string,
  conversationId?: string
): Promise<FiledEmailCache[string] | null> {
  if (!subject) {
    return null;
  }

  try {
    // Platform detection
    const platform = {
      host: (Office as any)?.context?.mailbox?.diagnostics?.hostName,
      hostVersion: (Office as any)?.context?.mailbox?.diagnostics?.hostVersion,
      platform: (Office as any)?.context?.platform,
    };
    console.log("[findFiledEmailBySubject] Platform info:", platform);

    // Force fresh read to avoid stale roamingSettings after recent write
    const raw = await getStored(FILED_CACHE_KEY, true);
    if (!raw) {
      console.log("[findFiledEmailBySubject] No cache found in storage");
      return null;
    }

    const cache: FiledEmailCache = JSON.parse(String(raw));
    const cacheKeys = Object.keys(cache);
    console.log("[findFiledEmailBySubject] Cache size:", cacheKeys.length, "keys");

    const tempKey = `subj:${subject.trim().toLowerCase()}`;
    console.log("[findFiledEmailBySubject] Looking for temp key:", tempKey);
    console.log("[findFiledEmailBySubject] Subject-based keys in cache:", cacheKeys.filter(k => k.startsWith("subj:")).length);

    const entry = cache[tempKey];

    if (entry) {
      console.log("[findFiledEmailBySubject] ✅ Found cache entry by subject", {
        subject,
        caseId: entry.caseId,
        documentId: entry.documentId,
        filedAt: new Date(entry.filedAt).toISOString(),
      });

      // Upgrade cache: If we now have conversationId, store under that key too
      if (conversationId) {
        console.log("[findFiledEmailBySubject] Upgrading cache with conversationId:", conversationId.substring(0, 30) + "...");
        cache[conversationId] = entry;
        // Keep the subject-based entry for a while (don't delete)
        await setStored(FILED_CACHE_KEY, JSON.stringify(cache));

        // Verify upgrade succeeded
        const verification = await getStored(FILED_CACHE_KEY);
        const verifiedCache = verification ? JSON.parse(String(verification)) : {};
        const upgradeSuccess = !!verifiedCache[conversationId];
        console.log("[findFiledEmailBySubject] Cache upgrade verification:", {
          success: upgradeSuccess,
          cacheSize: Object.keys(verifiedCache).length,
        });
      }

      return entry;
    }

    console.log("[findFiledEmailBySubject] ❌ No entry for this subject");
    return null;
  } catch (e) {
    console.warn("[findFiledEmailBySubject] Failed:", e);
    return null;
  }
}

/**
 * Remove filed email from cache (e.g., if document was deleted)
 */
export async function removeFiledEmailFromCache(conversationId: string): Promise<void> {
  if (!conversationId) {
    return;
  }

  try {
    const raw = await getStored(FILED_CACHE_KEY);
    if (!raw) return;

    const cache: FiledEmailCache = JSON.parse(String(raw));
    delete cache[conversationId];

    await setStored(FILED_CACHE_KEY, JSON.stringify(cache));
    console.log("[removeFiledEmailFromCache] Removed entry", {
      conversationId: conversationId.substring(0, 20) + "...",
    });
  } catch (e) {
    console.warn("[removeFiledEmailFromCache] Failed:", e);
  }
}
