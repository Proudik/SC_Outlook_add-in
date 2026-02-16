/* global OfficeRuntime */

import { getAuth, getAuthRuntime } from "./auth";
import { getStored } from "../utils/storage";
import { STORAGE_KEYS } from "../utils/constants";
import { uploadDocumentToCase } from "./singlecaseDocuments";

type DiagnosticResult = {
  success: boolean;
  step: string;
  message: string;
  details?: any;
};

function toBase64Utf8(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

export async function runDiagnostics(): Promise<DiagnosticResult[]> {
  const results: DiagnosticResult[] = [];

  // Step 1: Check sessionStorage token
  try {
    const auth = getAuth();
    if (auth?.token) {
      results.push({
        success: true,
        step: "sessionStorage token",
        message: "✓ Token found in sessionStorage",
        details: {
          hasToken: true,
          tokenPrefix: auth.token.slice(0, 10),
          email: auth.email,
          issuedAt: new Date(auth.issuedAt).toISOString(),
        },
      });
    } else {
      results.push({
        success: false,
        step: "sessionStorage token",
        message: "✗ No token in sessionStorage",
        details: { hasToken: false },
      });
    }
  } catch (e) {
    results.push({
      success: false,
      step: "sessionStorage token",
      message: `✗ Error reading sessionStorage: ${e instanceof Error ? e.message : String(e)}`,
    });
  }

  // Step 2: Check OfficeRuntime.storage token
  try {
    const authRuntime = await getAuthRuntime();
    if (authRuntime?.token) {
      results.push({
        success: true,
        step: "OfficeRuntime.storage token",
        message: "✓ Token found in OfficeRuntime.storage",
        details: {
          hasToken: true,
          tokenPrefix: authRuntime.token.slice(0, 10),
          email: authRuntime.email,
          issuedAt: new Date(authRuntime.issuedAt).toISOString(),
        },
      });
    } else {
      results.push({
        success: false,
        step: "OfficeRuntime.storage token",
        message: "✗ No token in OfficeRuntime.storage",
        details: { hasToken: false },
      });
    }
  } catch (e) {
    results.push({
      success: false,
      step: "OfficeRuntime.storage token",
      message: `✗ Error reading OfficeRuntime.storage: ${e instanceof Error ? e.message : String(e)}`,
    });
  }

  // Step 3: Check workspaceHost
  try {
    const hostRaw = await getStored(STORAGE_KEYS.workspaceHost);
    if (hostRaw) {
      results.push({
        success: true,
        step: "workspaceHost",
        message: "✓ Workspace host found",
        details: { host: hostRaw },
      });
    } else {
      results.push({
        success: false,
        step: "workspaceHost",
        message: "✗ No workspace host configured",
        details: { host: null },
      });
    }
  } catch (e) {
    results.push({
      success: false,
      step: "workspaceHost",
      message: `✗ Error reading workspace host: ${e instanceof Error ? e.message : String(e)}`,
    });
  }

  // Step 4: Test GET /cases
  try {
    const auth = getAuth();
    const token = auth?.token || (await getAuthRuntime()).token;
    const hostRaw = await getStored(STORAGE_KEYS.workspaceHost);
    const host = (hostRaw || "").trim().replace(/^https?:\/\//i, "").split("/")[0];

    if (!token) {
      results.push({
        success: false,
        step: "GET /cases",
        message: "✗ Skipped: no auth token",
      });
    } else if (!host) {
      results.push({
        success: false,
        step: "GET /cases",
        message: "✗ Skipped: no workspace host",
      });
    } else {
      const url = `/singlecase/${encodeURIComponent(host)}/publicapi/v1/cases`;
      const res = await fetch(url, {
        method: "GET",
        headers: {
          Authentication: token,
          "Content-Type": "application/json",
        },
      });

      if (res.ok) {
        const data = await res.json();
        results.push({
          success: true,
          step: "GET /cases",
          message: "✓ Successfully fetched cases",
          details: {
            status: res.status,
            caseCount: Array.isArray(data) ? data.length : "unknown",
          },
        });
      } else {
        const text = await res.text().catch(() => "");
        results.push({
          success: false,
          step: "GET /cases",
          message: `✗ Failed with status ${res.status}`,
          details: {
            status: res.status,
            statusText: res.statusText,
            responseSnippet: text.slice(0, 200),
          },
        });
      }
    }
  } catch (e) {
    results.push({
      success: false,
      step: "GET /cases",
      message: `✗ Request failed: ${e instanceof Error ? e.message : String(e)}`,
    });
  }

  // Step 5: Test POST /documents with tiny test .eml
  try {
    const auth = getAuth();
    const token = auth?.token || (await getAuthRuntime()).token;
    const hostRaw = await getStored(STORAGE_KEYS.workspaceHost);

    if (!token) {
      results.push({
        success: false,
        step: "POST /documents test",
        message: "✗ Skipped: no auth token",
      });
    } else if (!hostRaw) {
      results.push({
        success: false,
        step: "POST /documents test",
        message: "✗ Skipped: no workspace host",
      });
    } else {
      // Ask user for a test case ID
      const testCaseId = prompt("Enter a test case ID for upload diagnostic (or cancel to skip):");
      if (!testCaseId) {
        results.push({
          success: false,
          step: "POST /documents test",
          message: "✗ Skipped: no case ID provided",
        });
      } else {
        const testEmail =
          "From: Test <test@example.com>\r\n" +
          "To: SingleCase <noreply@singlecase>\r\n" +
          "Subject: Diagnostic Test Email\r\n" +
          "Date: " +
          new Date().toUTCString() +
          "\r\n" +
          "MIME-Version: 1.0\r\n" +
          "Content-Type: text/plain; charset=UTF-8\r\n" +
          "\r\n" +
          "This is a diagnostic test email from the SingleCase Outlook add-in.\r\n";

        const testBase64 = toBase64Utf8(testEmail);

        const uploadResult = await uploadDocumentToCase({
          caseId: testCaseId,
          fileName: "diagnostic-test.eml",
          mimeType: "message/rfc822",
          dataBase64: testBase64,
        });

        const docId = uploadResult?.documents?.[0]?.id;
        results.push({
          success: true,
          step: "POST /documents test",
          message: "✓ Successfully uploaded test document",
          details: {
            documentId: docId,
            caseId: testCaseId,
            response: uploadResult,
          },
        });
      }
    }
  } catch (e) {
    results.push({
      success: false,
      step: "POST /documents test",
      message: `✗ Upload failed: ${e instanceof Error ? e.message : String(e)}`,
    });
  }

  return results;
}

export function formatDiagnosticResults(results: DiagnosticResult[]): string {
  let output = "=== SingleCase Diagnostics ===\n\n";

  for (const r of results) {
    output += `${r.message}\n`;
    if (r.details) {
      output += `  Details: ${JSON.stringify(r.details, null, 2)}\n`;
    }
    output += "\n";
  }

  const successCount = results.filter((r) => r.success).length;
  const totalCount = results.length;

  output += `\n=== Summary: ${successCount}/${totalCount} checks passed ===\n`;

  return output;
}
