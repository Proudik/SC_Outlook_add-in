// ─── Lambda / Express telemetry ingest handler ────────────────────────────────
// Accepts POST /ingest with body: { events: AnalyticsEvent[] }
// Validates, sanitises, and logs each event as a structured JSON line to
// stdout → CloudWatch Logs.
//
// Deploy as:
//   - AWS Lambda + Function URL (no API Gateway needed for telemetry)
//   - Express route in your existing Node backend
//
// Environment variables:
//   MAX_EVENTS_PER_BATCH   Max events accepted in a single request (default 50)
//   ALLOWED_SCHEMA_VERSION Reject events with a different schemaVersion (default 1)

import { log } from "./logger";

const MAX_EVENTS = Number(process.env.MAX_EVENTS_PER_BATCH || 50);
const ALLOWED_SCHEMA = Number(process.env.ALLOWED_SCHEMA_VERSION || 1);

// ─── Minimal runtime validation ──────────────────────────────────────────────

function isString(v: unknown): v is string {
  return typeof v === "string";
}

function isValidEvent(e: unknown): boolean {
  if (!e || typeof e !== "object") return false;
  const ev = e as Record<string, unknown>;
  if (!isString(ev.eventName)) return false;
  if (!ev.context || typeof ev.context !== "object") return false;
  const ctx = ev.context as Record<string, unknown>;
  if (ctx.schemaVersion !== ALLOWED_SCHEMA) return false;
  if (!isString(ctx.anonymousUserId)) return false;
  if (!isString(ctx.sessionId)) return false;
  if (!isString(ctx.timestampIso)) return false;
  return true;
}

/**
 * Strips any field that might accidentally contain PII.
 * The frontend should never send raw PII but this is a server-side safety net.
 */
function sanitiseEvent(e: Record<string, unknown>): Record<string, unknown> {
  const ctx = (e.context || {}) as Record<string, unknown>;
  const payload = (e.payload || {}) as Record<string, unknown>;

  // Redact any field that looks like a raw email address
  const emailRe = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
  const redact = (v: unknown): unknown => {
    if (typeof v === "string") return v.replace(emailRe, "[email]");
    return v;
  };

  const cleanPayload: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(payload)) {
    cleanPayload[k] = redact(v);
  }

  return { ...e, context: ctx, payload: cleanPayload };
}

// ─── Core handler (framework-agnostic) ───────────────────────────────────────

interface IngestResult {
  status: number;
  body: Record<string, unknown>;
}

export async function handleIngest(
  body: unknown,
  correlationId?: string
): Promise<IngestResult> {
  if (!body || typeof body !== "object") {
    return { status: 400, body: { error: "Missing or invalid body" } };
  }

  const { events } = body as { events?: unknown };
  if (!Array.isArray(events)) {
    return { status: 400, body: { error: "body.events must be an array" } };
  }

  if (events.length === 0) {
    return { status: 200, body: { accepted: 0, rejected: 0 } };
  }

  if (events.length > MAX_EVENTS) {
    return {
      status: 400,
      body: { error: `Batch too large (max ${MAX_EVENTS})` },
    };
  }

  let accepted = 0;
  let rejected = 0;

  for (const raw of events) {
    if (!isValidEvent(raw)) {
      rejected++;
      log("warn", "telemetry.ingest.invalid_event", { correlationId });
      continue;
    }

    const clean = sanitiseEvent(raw as Record<string, unknown>);
    const ctx = clean.context as Record<string, unknown>;

    // One structured log line per event → CloudWatch Logs Insights can query it.
    log("info", String(clean.eventName), {
      correlationId: correlationId || ctx.correlationId,
      schemaVersion: ctx.schemaVersion,
      anonymousUserId: ctx.anonymousUserId,
      workspaceId: ctx.workspaceId,
      sessionId: ctx.sessionId,
      addinVersion: ctx.addinVersion,
      buildSha: ctx.buildSha,
      environment: ctx.environment,
      eventTimestamp: ctx.timestampIso,
      payload: clean.payload,
    });

    accepted++;
  }

  return { status: 200, body: { accepted, rejected } };
}

// ─── AWS Lambda handler ───────────────────────────────────────────────────────
// Wire this as your Lambda handler:
//   exports.handler = lambdaHandler;
//
// Example serverless.yml:
//   handler: backend/src/telemetry/ingest.lambdaHandler
//   events:
//     - httpApi:
//         path: /ingest
//         method: POST

export async function lambdaHandler(event: {
  body?: string;
  headers?: Record<string, string>;
}): Promise<{ statusCode: number; body: string; headers: Record<string, string> }> {
  const correlationId = event.headers?.["x-correlation-id"] || undefined;

  let parsed: unknown;
  try {
    parsed = JSON.parse(event.body || "{}");
  } catch {
    return {
      statusCode: 400,
      headers: corsHeaders(),
      body: JSON.stringify({ error: "Invalid JSON body" }),
    };
  }

  const result = await handleIngest(parsed, correlationId);

  return {
    statusCode: result.status,
    headers: corsHeaders(),
    body: JSON.stringify(result.body),
  };
}

function corsHeaders(): Record<string, string> {
  return {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type,X-Correlation-ID",
  };
}
