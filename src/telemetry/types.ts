// ─── Telemetry Types ─────────────────────────────────────────────────────────
// schemaVersion guards every event so CloudWatch / Athena queries can filter
// safely when fields are added in future releases.
// No PII is ever present in these types — all user/item/case identifiers
// are SHA-256 hashes computed in hashing.ts before reaching this layer.

export type EventName =
  | "addin.load"
  | "suggestion.shown"
  | "suggestion.accepted"
  | "suggestion.dismissed"
  | "case.selected.manual"
  | "filing.started"
  | "filing.succeeded"
  | "filing.failed"
  | "upload.started"
  | "upload.succeeded"
  | "upload.failed"
  | "settings.changed"
  | "auth.refreshed"
  | "auth.failed";

/** Stamped on every event. Built once at initTelemetry(), cloned per emit(). */
export interface EventContext {
  schemaVersion: 1;
  anonymousUserId: string; // SHA-256(email + workspaceHost)
  workspaceId: string;     // SHA-256(workspaceHost)
  addinVersion: string;    // __ADDIN_VERSION__
  buildSha: string;        // __BUILD_SHA__
  environment: "development" | "production";
  sessionId: string;       // random UUID, reset on Office.onReady
  correlationId?: string;  // UUID, per user action
  timestampIso: string;    // set at emit() time
}

export interface AnalyticsEvent<T extends Record<string, unknown> = Record<string, never>> {
  eventName: EventName;
  context: EventContext;
  payload: T;
}

// ─── Payload shapes ───────────────────────────────────────────────────────────

export interface AddinLoadPayload {
  composeMode: boolean;
  platform: string; // Office.context.diagnostics.platform or "unknown"
}

export interface SuggestionShownPayload {
  hashedItemId: string;
  count: number;
  topConfidencePct: number;
}

export interface SuggestionAcceptedPayload {
  hashedItemId: string;
  hashedCaseId: string;
  confidencePct: number;
  rank: number; // 0-based position in suggestion list
}

export interface SuggestionDismissedPayload {
  hashedItemId: string;
  count: number;
}

export interface CaseSelectedManualPayload {
  hashedItemId: string;
  hashedCaseId: string;
}

export interface FilingStartedPayload {
  hashedItemId: string;
  hashedCaseId: string;
  filingMode: "both" | "attachments";
  attachmentCount: number;
  selectionSource: "suggested" | "remembered" | "last_case" | "manual" | "";
  confidencePct: number;
  isNewVersion: boolean;
}

export interface FilingSucceededPayload extends FilingStartedPayload {
  durationMs: number;
  documentCount: number;
}

export interface FilingFailedPayload extends FilingStartedPayload {
  durationMs: number;
  errorCode: string;
  sanitizedMessage: string;
}

export interface UploadStartedPayload {
  hashedCaseId: string;
  kind: "email" | "attachment";
}

export interface UploadSucceededPayload extends UploadStartedPayload {
  durationMs: number;
}

export interface UploadFailedPayload extends UploadStartedPayload {
  durationMs: number;
  sanitizedMessage: string;
}

export interface SettingsChangedPayload {
  changedKeys: string[];
}

export interface AuthPayload {
  reason: string;
}
