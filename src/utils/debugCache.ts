/**
 * Debug utilities for filed email cache
 * Available as window commands for console debugging
 */

import { getStored, getDebugLog, clearDebugLog } from "./storage";

const FILED_CACHE_KEY = "sc:filedEmailsCache";

interface CacheEntry {
  caseId: string;
  documentId: string;
  subject: string;
  caseName?: string;
  caseKey?: string;
  filedAt: number;
}

interface FiledEmailCache {
  [key: string]: CacheEntry;
}

/**
 * Show cache contents in console
 */
export async function showCache(): Promise<void> {
  try {
    const raw = await getStored(FILED_CACHE_KEY);
    if (!raw) {
      console.log("📦 Cache is EMPTY");
      return;
    }

    const cache: FiledEmailCache = JSON.parse(String(raw));
    const entries = Object.entries(cache);

    console.log(`📦 Cache contains ${entries.length} entries:`);
    console.log("─".repeat(80));

    const tableData = entries.map(([key, entry]) => ({
      Key: key.substring(0, 40) + (key.length > 40 ? "..." : ""),
      Type: key.startsWith("subj:") ? "Subject" : "ConversationId",
      CaseId: entry.caseId,
      DocumentId: entry.documentId,
      Subject: entry.subject,
      FiledAt: new Date(entry.filedAt).toISOString(),
    }));

    console.table(tableData);

    console.log("\n📊 Statistics:");
    const convIdCount = entries.filter(([k]) => !k.startsWith("subj:")).length;
    const subjCount = entries.filter(([k]) => k.startsWith("subj:")).length;
    console.log(`  - ConversationId entries: ${convIdCount}`);
    console.log(`  - Subject entries: ${subjCount}`);
    console.log(`  - Total: ${entries.length}`);

    console.log("\n💡 Most recent entry:");
    const sorted = entries.sort((a, b) => b[1].filedAt - a[1].filedAt);
    if (sorted.length > 0) {
      const [key, entry] = sorted[0];
      console.log(`  - Key: ${key}`);
      console.log(`  - Subject: ${entry.subject}`);
      console.log(`  - Filed: ${new Date(entry.filedAt).toISOString()}`);
    }

    console.log("\n📝 Raw cache data:");
    console.log(JSON.stringify(cache, null, 2));
  } catch (e) {
    console.error("❌ Failed to read cache:", e);
  }
}

/**
 * Show platform information
 */
export function showPlatform(): void {
  try {
    const platform = {
      host: (Office as any)?.context?.mailbox?.diagnostics?.hostName,
      hostVersion: (Office as any)?.context?.mailbox?.diagnostics?.hostVersion,
      platform: (Office as any)?.context?.platform,
      officeVersion: (Office as any)?.context?.diagnostics?.version,
      hasOfficeRuntime: typeof OfficeRuntime !== "undefined",
      hasRoamingSettings: !!(Office as any)?.context?.roamingSettings,
    };

    console.log("🖥️  Platform Information:");
    console.log("─".repeat(80));
    console.table(platform);
  } catch (e) {
    console.error("❌ Failed to get platform info:", e);
  }
}

/**
 * Show debug log (persistent across sessions)
 */
export async function showDebugLog(): Promise<void> {
  try {
    const log = await getDebugLog();
    if (!log) {
      console.log("📄 Debug log is EMPTY");
      return;
    }

    console.log("📄 Debug Log (persistent):");
    console.log("─".repeat(80));
    console.log(log);
    console.log("─".repeat(80));
    console.log(`Total size: ${log.length} characters`);
  } catch (e) {
    console.error("❌ Failed to read debug log:", e);
  }
}

/**
 * Clear debug log
 */
export async function clearDebugLogCmd(): Promise<void> {
  try {
    await clearDebugLog();
    console.log("✅ Debug log cleared");
  } catch (e) {
    console.error("❌ Failed to clear debug log:", e);
  }
}

/**
 * Search cache by subject
 */
export async function searchCacheBySubject(subject: string): Promise<void> {
  try {
    const raw = await getStored(FILED_CACHE_KEY);
    if (!raw) {
      console.log("📦 Cache is empty");
      return;
    }

    const cache: FiledEmailCache = JSON.parse(String(raw));
    const tempKey = `subj:${subject.trim().toLowerCase()}`;

    console.log(`🔍 Searching for subject: "${subject}"`);
    console.log(`🔑 Looking for key: "${tempKey}"`);

    const entry = cache[tempKey];
    if (entry) {
      console.log("✅ Found entry:");
      console.table({
        CaseId: entry.caseId,
        DocumentId: entry.documentId,
        Subject: entry.subject,
        FiledAt: new Date(entry.filedAt).toISOString(),
      });
    } else {
      console.log("❌ No entry found for this subject");

      // Show similar subjects
      const subjectKeys = Object.keys(cache).filter(k => k.startsWith("subj:"));
      if (subjectKeys.length > 0) {
        console.log("\n📝 Available subject keys:");
        subjectKeys.forEach(k => {
          const subj = k.substring(5); // Remove "subj:" prefix
          console.log(`  - "${subj}"`);
        });
      }
    }
  } catch (e) {
    console.error("❌ Failed to search cache:", e);
  }
}

/**
 * Export all commands to window for console access
 */
export function installDebugCommands(): void {
  if (typeof window !== "undefined") {
    (window as any).scDebug = {
      showCache,
      showPlatform,
      showDebugLog,
      clearDebugLog: clearDebugLogCmd,
      searchBySubject: searchCacheBySubject,
      help: () => {
        console.log("🛠️  SingleCase Debug Commands:");
        console.log("─".repeat(80));
        console.log("scDebug.showCache()           - Show all cached filed emails");
        console.log("scDebug.showPlatform()        - Show platform information");
        console.log("scDebug.showDebugLog()        - Show persistent debug log");
        console.log("scDebug.clearDebugLog()       - Clear debug log");
        console.log("scDebug.searchBySubject('...')- Search cache by email subject");
        console.log("scDebug.help()                - Show this help");
        console.log("─".repeat(80));
      },
    };

    console.log("✅ Debug commands installed! Type scDebug.help() for available commands");
  }
}

declare const Office: any;
declare const OfficeRuntime: any;
