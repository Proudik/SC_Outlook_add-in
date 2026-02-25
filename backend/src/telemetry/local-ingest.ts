// ─── Local development telemetry ingest server ────────────────────────────────
// Run with: npx ts-node backend/src/telemetry/local-ingest.ts
// (or add to package.json scripts: "telemetry": "ts-node backend/src/telemetry/local-ingest.ts")
//
// Listens on http://localhost:4001/ingest
// Writes all received events as newline-delimited JSON to .telemetry/events.ndjson
// Matches the __TELEMETRY_ENDPOINT__ webpack default for development mode.
//
// To inspect events:
//   cat .telemetry/events.ndjson | jq 'select(.eventName == "filing.succeeded")'

import * as http from "http";
import * as fs from "fs";
import * as path from "path";
import { handleIngest } from "./ingest";

const PORT = Number(process.env.TELEMETRY_PORT || 4001);
const OUT_DIR = path.resolve(process.cwd(), ".telemetry");
const OUT_FILE = path.join(OUT_DIR, "events.ndjson");

// Ensure output directory exists
if (!fs.existsSync(OUT_DIR)) {
  fs.mkdirSync(OUT_DIR, { recursive: true });
}

const server = http.createServer(async (req, res) => {
  // CORS preflight
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,X-Correlation-ID");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");

  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.method !== "POST" || req.url !== "/ingest") {
    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "Not found" }));
    return;
  }

  // Read body
  const chunks: Buffer[] = [];
  req.on("data", (c: Buffer) => chunks.push(c));
  req.on("end", async () => {
    const correlationId = req.headers["x-correlation-id"] as string | undefined;

    let parsed: unknown;
    try {
      parsed = JSON.parse(Buffer.concat(chunks).toString("utf8"));
    } catch {
      res.writeHead(400, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: "Invalid JSON" }));
      return;
    }

    const result = await handleIngest(parsed, correlationId);

    // Append each event as a separate line to the ndjson file
    const body = parsed as { events?: unknown[] };
    if (Array.isArray(body.events) && body.events.length > 0) {
      const lines = body.events
        .map((e) => JSON.stringify(e))
        .join("\n") + "\n";
      fs.appendFile(OUT_FILE, lines, (err) => {
        if (err) console.error("[local-ingest] Failed to write events.ndjson:", err);
      });
      console.log(
        `[local-ingest] +${body.events.length} event(s) → ${OUT_FILE}`
      );
    }

    res.writeHead(result.status, { "Content-Type": "application/json" });
    res.end(JSON.stringify(result.body));
  });
});

server.listen(PORT, () => {
  console.log(`[local-ingest] Listening on http://localhost:${PORT}/ingest`);
  console.log(`[local-ingest] Writing events to ${OUT_FILE}`);
});
