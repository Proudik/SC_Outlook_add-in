// ─── Telemetry module ─────────────────────────────────────────────────────────
// Singleton. Initialised once at Office.onReady via initTelemetry().
//
// USAGE GUIDE — where to call each function:
//
//   initTelemetry()          → src/taskpane/index.tsx, inside Office.onReady(),
//                              after resolving anonymousUserId and workspaceId.
//
//   createCorrelationContext() → call at the START of any multi-step async action
//                              (doSubmit, file-on-send handler, upload loop).
//                              Pass correlationId through all emit() calls for
//                              that action and through the X-Correlation-ID header.
//
//   emit()                   → call at any event point. Failures are swallowed —
//                              telemetry must never block the user flow.
//
//   emitLatency()            → convenience wrapper for events that carry durationMs.
//                              Pass performance.now() or Date.now() as startMs.

import type { AnalyticsEvent, EventContext, EventName } from "./types";

// Build-time constants injected by webpack DefinePlugin (webpack.config.js).
// The declare block silences TypeScript; actual values come from DefinePlugin.
declare const __ADDIN_VERSION__: string;
declare const __BUILD_SHA__: string;
declare const __ENVIRONMENT__: string;
declare const __TELEMETRY_ENDPOINT__: string;

// ─── Internal state ───────────────────────────────────────────────────────────

let _baseCtx: Omit<EventContext, "correlationId" | "timestampIso"> | null = null;
let _endpoint = "";
let _queue: AnalyticsEvent<Record<string, unknown>>[] = [];
let _flushTimer: ReturnType<typeof setInterval> | null = null;

// ─── Helpers ─────────────────────────────────────────────────────────────────

function generateUUID(): string {
  if (
    typeof crypto !== "undefined" &&
    typeof (crypto as Crypto).randomUUID === "function"
  ) {
    return (crypto as Crypto).randomUUID();
  }
  // IE11 / old webview fallback
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
  });
}

function safeRead<T>(name: string, fallback: T): T {
  try {
    // Globals are replaced by DefinePlugin at build time.
    // At runtime they evaluate to string literals or "undefined".
    const map: Record<string, unknown> = {
      __ADDIN_VERSION__: typeof __ADDIN_VERSION__ !== "undefined" ? __ADDIN_VERSION__ : undefined,
      __BUILD_SHA__: typeof __BUILD_SHA__ !== "undefined" ? __BUILD_SHA__ : undefined,
      __ENVIRONMENT__: typeof __ENVIRONMENT__ !== "undefined" ? __ENVIRONMENT__ : undefined,
      __TELEMETRY_ENDPOINT__: typeof __TELEMETRY_ENDPOINT__ !== "undefined" ? __TELEMETRY_ENDPOINT__ : undefined,
    };
    const v = map[name];
    return v !== undefined ? (v as T) : fallback;
  } catch {
    return fallback;
  }
}

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Call once at startup (Office.onReady) after computing anonymousUserId and
 * workspaceId via hashing.ts. Idempotent — safe to call multiple times.
 *
 * @param opts.anonymousUserId  SHA-256(email + workspaceHost) from hashing.ts
 * @param opts.workspaceId      SHA-256(workspaceHost) from hashing.ts
 * @param opts.endpoint         Override the telemetry endpoint (useful in tests)
 */
export function initTelemetry(opts: {
  anonymousUserId: string;
  workspaceId: string;
  endpoint?: string;
}): void {
  const version = safeRead<string>("__ADDIN_VERSION__", "dev");
  const buildSha = safeRead<string>("__BUILD_SHA__", "unknown");
  const env = safeRead<string>("__ENVIRONMENT__", "development");
  const defaultEndpoint = safeRead<string>("__TELEMETRY_ENDPOINT__", "");

  _endpoint = opts.endpoint ?? defaultEndpoint;
  _baseCtx = {
    schemaVersion: 1,
    anonymousUserId: opts.anonymousUserId,
    workspaceId: opts.workspaceId,
    addinVersion: version,
    buildSha,
    environment: env === "production" ? "production" : "development",
    sessionId: generateUUID(),
  };

  if (_flushTimer) clearInterval(_flushTimer);
  _flushTimer = setInterval(() => {
    void _flush();
  }, 10_000);

  try {
    window.addEventListener("beforeunload", () => {
      void _flush();
    });
  } catch {
    // No window in some test environments
  }
}

/**
 * Creates a per-action correlation context.
 * Attach correlationId to all emit() calls in a single user action AND to the
 * X-Correlation-ID header sent with API calls so frontend + backend logs join.
 *
 * Example:
 *   const { correlationId } = createCorrelationContext();
 *   emit("filing.started", payload, correlationId);
 *   await submitEmailToCase(token, body, correlationId);
 *   emit("filing.succeeded", payload, correlationId);
 */
export function createCorrelationContext(): { correlationId: string } {
  return { correlationId: generateUUID() };
}

/**
 * Enqueues a telemetry event. Flushes automatically when queue reaches 10 events
 * or after 10 s. Swallows all errors — must never block the user flow.
 */
export function emit<T extends Record<string, unknown>>(
  eventName: EventName,
  payload: T,
  correlationId?: string
): void {
  if (!_baseCtx) return;
  if (!_endpoint) return;

  const event: AnalyticsEvent<T> = {
    eventName,
    context: {
      ..._baseCtx,
      correlationId,
      timestampIso: new Date().toISOString(),
    },
    payload,
  };

  _queue.push(event as AnalyticsEvent<Record<string, unknown>>);

  if (_queue.length >= 10) void _flush();
}

/**
 * Convenience wrapper for events that carry a durationMs field.
 * Pass Date.now() as startMs before the operation begins.
 */
export function emitLatency<T extends Record<string, unknown>>(
  eventName: EventName,
  startMs: number,
  payload: Omit<T, "durationMs">,
  correlationId?: string
): void {
  emit(
    eventName,
    { ...payload, durationMs: Date.now() - startMs } as unknown as T,
    correlationId
  );
}

// ─── Internal flush ───────────────────────────────────────────────────────────

async function _flush(): Promise<void> {
  if (_queue.length === 0) return;
  if (!_endpoint) return;

  const batch = _queue.splice(0, _queue.length);

  try {
    await fetch(_endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ events: batch }),
      keepalive: true, // survives beforeunload
    });
  } catch {
    // Silently drop — telemetry failures must never surface to the user.
    // Re-queue could cause infinite loops on a broken endpoint, so we discard.
  }
}
