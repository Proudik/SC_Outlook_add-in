/**
 * Internal Email Guardrail Utilities
 *
 * Detects if an email is being sent only to internal recipients (same company domain).
 * Used to prevent auto-filing of internal conversations.
 */

/**
 * Extract base domain from a domain string.
 * Treats subdomains as part of the same base domain.
 *
 * Examples:
 * - "singlecase.com" -> "singlecase.com"
 * - "eu.singlecase.com" -> "singlecase.com"
 * - "mail.google.com" -> "google.com"
 * - "co.uk" -> "co.uk" (edge case, but acceptable)
 *
 * @param domain - Domain string (e.g., "eu.singlecase.com")
 * @returns Base domain (e.g., "singlecase.com")
 */
export function getBaseDomain(domain: string): string {
  const normalized = domain.toLowerCase().trim();
  if (!normalized) return "";

  const parts = normalized.split(".");

  // If only one part (e.g., "localhost"), return as-is
  if (parts.length <= 1) return normalized;

  // If two parts (e.g., "singlecase.com"), return as-is
  if (parts.length === 2) return normalized;

  // If three or more parts (e.g., "eu.singlecase.com"), return last two
  // Edge case: "co.uk" would become "co.uk" which is acceptable
  return parts.slice(-2).join(".");
}

/**
 * Parse email address to extract domain.
 * Handles formats like:
 * - "user@domain.com"
 * - "Name <user@domain.com>"
 * - " user@domain.com " (with whitespace)
 *
 * @param email - Email address string
 * @returns Domain in lowercase, or empty string if invalid
 */
export function parseEmailDomain(email: string): string {
  const trimmed = email.trim();
  if (!trimmed) return "";

  // Handle "Name <email@domain.com>" format
  const angleMatch = trimmed.match(/<([^>]+)>/);
  const cleanEmail = angleMatch ? angleMatch[1] : trimmed;

  // Extract domain after @
  const atIndex = cleanEmail.lastIndexOf("@");
  if (atIndex === -1) return "";

  const domain = cleanEmail.slice(atIndex + 1).toLowerCase().trim();
  return domain;
}

/**
 * Determine if an email is internal (all recipients are from internal domains).
 *
 * Rules:
 * - If there are no recipients, returns false (not internal)
 * - If ANY recipient is external, returns false
 * - If ALL recipients are internal, returns true
 * - Subdomains are treated as internal (e.g., eu.singlecase.com matches singlecase.com)
 *
 * @param senderEmail - Sender's email address (used to determine default internal domain)
 * @param recipientEmails - Array of recipient email addresses (To + Cc)
 * @param internalDomainAllowlist - Optional array of additional internal domains
 * @returns true if all recipients are internal, false otherwise
 */
export function isInternalEmail(
  senderEmail: string,
  recipientEmails: string[],
  internalDomainAllowlist: string[] = []
): boolean {
  // If no recipients, not internal
  if (recipientEmails.length === 0) return false;

  // Build set of internal base domains
  const internalBaseDomains = new Set<string>();

  // Add sender's base domain as internal
  const senderDomain = parseEmailDomain(senderEmail);
  if (senderDomain) {
    const senderBaseDomain = getBaseDomain(senderDomain);
    if (senderBaseDomain) {
      internalBaseDomains.add(senderBaseDomain);
    }
  }

  // Add allowlist domains
  for (const domain of internalDomainAllowlist) {
    const baseDomain = getBaseDomain(domain.toLowerCase().trim());
    if (baseDomain) {
      internalBaseDomains.add(baseDomain);
    }
  }

  // If we have no internal domains to check against, treat as external
  if (internalBaseDomains.size === 0) return false;

  // Check each recipient
  for (const recipientEmail of recipientEmails) {
    const recipientDomain = parseEmailDomain(recipientEmail);
    if (!recipientDomain) {
      // Invalid email format - treat as external to be safe
      return false;
    }

    const recipientBaseDomain = getBaseDomain(recipientDomain);
    if (!recipientBaseDomain) {
      // Invalid domain - treat as external
      return false;
    }

    // Check if recipient's base domain matches any internal domain
    if (!internalBaseDomains.has(recipientBaseDomain)) {
      // External recipient found
      return false;
    }
  }

  // All recipients are internal
  return true;
}
