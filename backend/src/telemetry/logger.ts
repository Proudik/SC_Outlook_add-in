// ─── Structured JSON logger for CloudWatch Logs ───────────────────────────────
// Lambda stdout is automatically forwarded to CloudWatch Logs.
// Each line must be a complete JSON object so CWL Insights can parse it.
// Usage:
//   import { log } from "./logger";
//   log("info", "filing.succeeded", { correlationId, hashedCaseId, durationMs });

export type LogLevel = "info" | "warn" | "error";

export interface LogEntry {
  level: LogLevel;
  service: string;
  eventName: string;
  timestamp: string;
  [key: string]: unknown;
}

const SERVICE = process.env.SERVICE_NAME || "sc-telemetry-ingest";

export function log(
  level: LogLevel,
  eventName: string,
  fields: Record<string, unknown> = {}
): void {
  const entry: LogEntry = {
    level,
    service: SERVICE,
    eventName,
    timestamp: new Date().toISOString(),
    ...fields,
  };
  // CloudWatch Logs captures everything written to stdout.
  // JSON.stringify on one line = one parseable CWL record.
  // eslint-disable-next-line no-console
  console.log(JSON.stringify(entry));
}
