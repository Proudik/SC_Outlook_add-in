/**
 * Utilities for extracting internetMessageId (Message-ID header)
 * This is the stable cross-mailbox identifier for email messages
 */

/**
 * Extract internetMessageId from Office.js item (read mode)
 * Returns null if not available via Office API
 */
export function getInternetMessageIdFromItem(): string | null {
  try {
    const item = Office?.context?.mailbox?.item as any;

    // Try different property names
    const internetMessageId =
      item?.internetMessageId ||
      item?.itemId || // Sometimes contains the internet message ID
      null;

    if (internetMessageId && typeof internetMessageId === "string") {
      console.log("[getInternetMessageIdFromItem] Found:", internetMessageId.substring(0, 50) + "...");
      return internetMessageId;
    }

    console.log("[getInternetMessageIdFromItem] Not available via Office.js");
    return null;
  } catch (e) {
    console.error("[getInternetMessageIdFromItem] Error:", e);
    return null;
  }
}

/**
 * Normalize internetMessageId for storage/comparison
 * Removes angle brackets if present: <id@server> -> id@server
 */
export function normalizeInternetMessageId(messageId: string): string {
  if (!messageId) return "";

  let normalized = messageId.trim();

  // Remove angle brackets if present
  if (normalized.startsWith("<") && normalized.endsWith(">")) {
    normalized = normalized.substring(1, normalized.length - 1);
  }

  return normalized;
}

/**
 * Extract REST ID from Office.js item for Graph API calls
 * In read mode, this is the ID we use to fetch message details
 */
export function getRestIdFromItem(): string | null {
  try {
    const item = Office?.context?.mailbox?.item as any;
    const itemId = item?.itemId;

    if (!itemId) {
      console.warn("[getRestIdFromItem] No itemId available");
      return null;
    }

    // In some cases, we need to convert EWS ID to REST ID
    // For simplicity, we'll try using itemId directly first
    // If it's an EWS ID (starts with certain patterns), conversion may be needed

    return itemId;
  } catch (e) {
    console.error("[getRestIdFromItem] Error:", e);
    return null;
  }
}

/**
 * Convert EWS ID to REST ID format if needed
 * Office.js sometimes returns EWS IDs that need conversion for Graph API
 */
export function convertToRestId(ewsId: string): string {
  // If it's already a REST ID (base64-like), return as-is
  if (!ewsId.includes("=") && ewsId.length > 100) {
    return ewsId;
  }

  // For EWS IDs, try conversion
  // Note: Office.mailbox.convertToRestId() is the official method if available
  try {
    if (Office?.context?.mailbox?.convertToRestId) {
      return Office.context.mailbox.convertToRestId(
        ewsId,
        Office.MailboxEnums.RestVersion.v2_0
      );
    }
  } catch (e) {
    console.warn("[convertToRestId] Conversion failed:", e);
  }

  // Fallback: return as-is and hope it works
  return ewsId;
}
