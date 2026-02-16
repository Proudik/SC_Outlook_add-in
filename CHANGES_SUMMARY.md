# SingleCase Outlook Add-in: Debug Instrumentation Summary

## What Was Done

Added comprehensive diagnostic logging and a self-test utility to identify why emails are not being filed into SingleCase.

---

## Key Changes

### 1. ✅ Verbose Logging Added
- **[singlecaseDocuments.ts](src/services/singlecaseDocuments.ts)**: Added detailed logs for token retrieval, URL resolution, and upload process
- **[MainWorkspace.tsx](src/taskpane/components/MainWorkspace/MainWorkspace.tsx)**: Added logs to `doSubmit()` with request/response details
- **[onMessageSendHandler.ts](src/commands/onMessageSendHandler.ts)**: Added comprehensive logging throughout handler, no more silent failures

### 2. ✅ Self-Test Diagnostic Tool
- **[diagnostics.ts](src/services/diagnostics.ts)**: NEW - Built-in diagnostic utility
- **Tests**:
  - ✓ Token in sessionStorage
  - ✓ Token in OfficeRuntime.storage
  - ✓ Workspace host configuration
  - ✓ GET /cases API (through relay)
  - ✓ POST /documents API (with test .eml)

### 3. ✅ UI Button for Diagnostics
- Added **"🔧 Run Diagnostics"** button to taskpane (top of page)
- One-click to run all tests and see results

### 4. ✅ MIME Type Fixed
- Changed .eml uploads from `"text/plain"` → `"message/rfc822"` for proper email handling
- Now consistent between manual filing and auto-filing

---

## How to Debug

### Quick Start (3 steps)
1. Open the add-in taskpane
2. Click **"🔧 Run Diagnostics"** button
3. Check console (F12) for detailed logs

### What to Look For

#### ✅ Success Pattern
```
[uploadDocumentToCase] Starting upload ...
[getToken] Using sessionStorage token
[resolveApiBaseUrl] Resolved base URL: /singlecase/{host}/publicapi/v1
[uploadDocumentToCase] Fetch completed {status: 200, ok: true}
[uploadDocumentToCase] Upload successful
```

#### ❌ Failure Patterns

**No token:**
```
[getToken] No token found in either sessionStorage or OfficeRuntime.storage
```
→ **Fix**: Re-login through the add-in

**No workspace:**
```
[resolveApiBaseUrl] Workspace host is missing
```
→ **Fix**: Re-select workspace in settings

**Network error:**
```
[uploadDocumentToCase] Network request failed: Failed to fetch
```
→ **Fix**: Verify dev server is running, check if relay is accessible

**API error:**
```
[uploadDocumentToCase] Upload failed {status: 401, responseSnippet: "..."}
```
→ **Fix**: Check token validity, verify case ID, review error message

---

## Testing Workflow

### Test 1: Manual Filing
1. Open taskpane
2. Select a sent email
3. Choose a case
4. Click **"Zařadit teď"**
5. **Check console** for `[doSubmit]` and `[uploadDocumentToCase]` logs
6. Verify email appears in SingleCase

### Test 2: Auto-Filing (OnMessageSend)
1. Create new email (compose mode)
2. Add recipient
3. Click **"Zařadit teď"** to select case
4. **Send** the email
5. **Check console** for `[onMessageSendHandler]` logs
6. Check Outlook notification area for success/error message
7. Verify email appears in SingleCase

---

## Common Issues & Solutions

| Symptom | Likely Cause | Solution |
|---------|-------------|----------|
| "Missing auth token" | Token not in OfficeRuntime.storage | Re-login |
| "Workspace host is missing" | Host not saved | Re-select workspace |
| "Network request failed" | Relay not accessible | Check dev server at https://localhost:3000 |
| 401 error | Expired/invalid token | Re-login |
| 404 error | Invalid case ID or endpoint | Verify case exists |
| Upload succeeds but email invisible | Wrong MIME type (now fixed) | Already using message/rfc822 |

---

## Files Modified

| File | Changes | Lines |
|------|---------|-------|
| [singlecaseDocuments.ts](src/services/singlecaseDocuments.ts) | Added verbose logging to upload, token, and URL functions | 49-250 |
| [MainWorkspace.tsx](src/taskpane/components/MainWorkspace/MainWorkspace.tsx) | Added logging to doSubmit, diagnostic button, MIME type fix | Multiple |
| [onMessageSendHandler.ts](src/commands/onMessageSendHandler.ts) | Added comprehensive logging, better error reporting | 149-250 |
| [diagnostics.ts](src/services/diagnostics.ts) | **NEW** - Self-test utility | All |
| [DEBUG_PLAN.md](DEBUG_PLAN.md) | **NEW** - Complete debugging guide | All |

---

## Next Steps

1. **Run diagnostics**: Click the button and review results
2. **Check console**: Open browser dev tools (F12) and look for `[component]` prefixed logs
3. **Test manual filing**: Try "Zařadit teď" and verify success
4. **Test auto-filing**: Send a compose email and verify it files
5. **Review logs**: Every step now logs its actions and results

---

## Documentation

See **[DEBUG_PLAN.md](DEBUG_PLAN.md)** for:
- Complete step-by-step checklist
- Detailed hypothesis ranking
- Console output examples
- Troubleshooting guide
- Expected vs actual behavior

---

## Need Help?

1. Run diagnostics and check which tests fail
2. Review console logs for error messages
3. Check [DEBUG_PLAN.md](DEBUG_PLAN.md) Section F for interpretation
4. Look for the specific error pattern in Section J (Common Issues)

All logging is now comprehensive - no more silent failures!
