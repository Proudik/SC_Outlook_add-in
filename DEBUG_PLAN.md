# SingleCase Outlook Add-in Debug Plan

## Executive Summary

This document provides a systematic approach to debug why emails are not being filed into SingleCase, with comprehensive instrumentation added to identify the exact failure point.

---

## A) Step-by-Step Verification Checklist

### Phase 1: Verify Manual Filing (Baseline Test)
**Goal**: Determine if manual filing works at all

1. ✅ **Test manual "Zařadit teď" button**
   - Open taskpane in Outlook
   - Select a sent email
   - Choose a case
   - Click "Zařadit teď"
   - Open browser console (F12) and check for errors
   - **Location**: [MainWorkspace.tsx:1530-1797](src/taskpane/components/MainWorkspace/MainWorkspace.tsx#L1530)
   - **Expected**: Email appears in SingleCase, console shows success logs
   - **If this fails**: Focus on auth/host/API issues, not OnMessageSend

### Phase 2: Verify Storage & Auth
**Goal**: Confirm critical configuration is accessible

2. ✅ **Check workspaceHost in storage**
   - In browser console: `await OfficeRuntime.storage.getItem("sc:workspaceHost")`
   - **Expected**: Returns your workspace host (e.g., `"valfor-demo.singlecase.ch"`)
   - **Location**: [storage.ts:1-20](src/utils/storage.ts), [constants.ts:5](src/utils/constants.ts)

3. ✅ **Check auth token in sessionStorage**
   - In browser console: `sessionStorage.getItem("singlecase_token")`
   - **Expected**: Returns your auth token string
   - **Location**: [auth.ts:45-55](src/services/auth.ts)

4. ✅ **Check auth token in OfficeRuntime.storage**
   - In browser console: `await OfficeRuntime.storage.getItem("sc_token")`
   - **Expected**: Returns the same auth token
   - **Location**: [auth.ts:58-70](src/services/auth.ts)

### Phase 3: Verify Relay Middleware
**Goal**: Confirm the webpack dev server relay is working

5. ✅ **Test relay from browser console**
   ```javascript
   const token = sessionStorage.getItem("singlecase_token");
   const host = "valfor-demo.singlecase.ch"; // replace with your host
   fetch(`/singlecase/${host}/publicapi/v1/cases`, {
     headers: { Authentication: token }
   }).then(r => r.json()).then(console.log);
   ```
   - **Expected**: Returns list of cases
   - **Location**: [webpack.config.js:127-184](webpack.config.js#L127)

### Phase 4: Verify Upload API
**Goal**: Confirm document upload endpoint works

6. ✅ **Use built-in diagnostics**
   - Open taskpane
   - Click "🔧 Run Diagnostics" button (top of page)
   - Follow prompts and check console output
   - **Location**: [diagnostics.ts](src/services/diagnostics.ts)
   - **What it tests**:
     - ✓ Token in sessionStorage
     - ✓ Token in OfficeRuntime.storage
     - ✓ Workspace host configuration
     - ✓ GET /cases API call
     - ✓ POST /documents with test .eml file

### Phase 5: Verify OnMessageSend Handler
**Goal**: Confirm the send handler executes and files correctly

7. ✅ **Test OnMessageSend in compose mode**
   - Create a new email in Outlook
   - Add a recipient
   - Wait for case suggestion (if enabled)
   - Click "Zařadit teď" to pre-file
   - **Send** the email
   - Check browser console for `[onMessageSendHandler]` logs
   - Check Outlook notification area for success/failure message
   - **Location**: [onMessageSendHandler.ts:149-244](src/commands/onMessageSendHandler.ts)

---

## B) Most Likely Failure Points (Ranked by Probability)

### H1: Runtime context cannot reach relay (CRITICAL)
**Probability**: Very High
**Symptoms**: Network errors, CORS errors, timeout errors
**Root cause**: Commands runtime might be calling `https://{host}/publicapi/v1/...` directly instead of using the relay at `https://localhost:3000/singlecase/{host}/...`

**Evidence to check**:
- Console shows: `[uploadDocumentToCase] Network request failed`
- Or: `TypeError: Failed to fetch`
- Or: CORS error in console

**Fix**: Ensure commands runtime uses `https://localhost:3000` as base URL

---

### H2: Token missing in OfficeRuntime.storage
**Probability**: High
**Symptoms**: 401 errors, "Missing auth token" errors
**Root cause**: Auth token not mirrored to OfficeRuntime.storage, or mirroring failed

**Evidence to check**:
- Console shows: `[getToken] No token found in either sessionStorage or OfficeRuntime.storage`
- Or: Upload fails with 401 status

**Fix**: Verify [auth.ts:72-87](src/services/auth.ts) `setAuth()` is called on login and successfully mirrors to OfficeRuntime.storage

---

### H3: workspaceHost not in storage
**Probability**: Medium
**Symptoms**: "Workspace host is missing" errors
**Root cause**: Host not saved when workspace is selected

**Evidence to check**:
- Console shows: `[resolveApiBaseUrl] Workspace host is missing`
- `await OfficeRuntime.storage.getItem("sc:workspaceHost")` returns null

**Fix**: Ensure workspace selection calls `setStored(STORAGE_KEYS.workspaceHost, host)`

---

### H4: MIME type mismatch
**Probability**: Low
**Symptoms**: Upload succeeds (200 OK) but document not visible in SingleCase UI
**Root cause**: Wrong MIME type causes SingleCase to reject or hide document

**Evidence to check**:
- Console shows: `[uploadDocumentToCase] Upload successful`
- But document doesn't appear in case
- Response returns document ID

**Fix**: Changed manual filing to use `"message/rfc822"` instead of `"text/plain"` (already done in OnMessageSend handler)

---

### H5: Silent failures with no diagnostics
**Probability**: N/A (now fixed)
**Symptoms**: No errors visible anywhere
**Root cause**: All catch blocks silently swallow errors

**Fix**: ✅ Added comprehensive logging throughout the codebase

---

## C) Instrumentation Added

### 1. singlecaseDocuments.ts
**Changes**: Added verbose logging to `uploadDocumentToCase()`, `getToken()`, and `resolveApiBaseUrl()`

**New logs**:
- `[uploadDocumentToCase] Starting upload` - Entry point
- `[getToken] Using sessionStorage token` - Token source
- `[resolveApiBaseUrl] Resolved base URL` - Computed relay path
- `[uploadDocumentToCase] Fetch completed` - HTTP status
- `[uploadDocumentToCase] Upload successful` - Success case
- `[uploadDocumentToCase] Upload failed` - Error case with response snippet

**Location**: [singlecaseDocuments.ts:161-250](src/services/singlecaseDocuments.ts)

---

### 2. MainWorkspace.tsx
**Changes**: Added logging to `doSubmit()` function

**New logs**:
- `[doSubmit] About to upload email document` - Pre-upload state
- `[doSubmit] Creating new email document` - New doc path
- `[doSubmit] Email document created` - Success with response
- `[doSubmit] Upload failed` - Error with full details

**Location**: [MainWorkspace.tsx:1600-1770](src/taskpane/components/MainWorkspace/MainWorkspace.tsx)

---

### 3. onMessageSendHandler.ts
**Changes**: Added comprehensive logging throughout handler, NO silent failures

**New logs**:
- `[onMessageSendHandler] Handler fired` - Entry point
- `[onMessageSendHandler] Item keys: [...]` - Email identifiers
- `[onMessageSendHandler] Intent: {...}` - Auto-file decision
- `[onMessageSendHandler] Token retrieved` - Auth check
- `[onMessageSendHandler] Workspace host` - Config check
- `[onMessageSendHandler] Uploading document to case` - Pre-upload
- `[onMessageSendHandler] Upload successful` - Success
- `[onMessageSendHandler] Error during filing` - Failure with stack trace

**User notifications**: Handler now shows Outlook notification with error message on failure

**Location**: [onMessageSendHandler.ts:149-244](src/commands/onMessageSendHandler.ts)

---

### 4. diagnostics.ts (NEW)
**Purpose**: Self-contained diagnostic utility to test all critical paths

**What it does**:
1. ✅ Checks token in sessionStorage
2. ✅ Checks token in OfficeRuntime.storage
3. ✅ Checks workspaceHost configuration
4. ✅ Tests GET /cases API call through relay
5. ✅ Tests POST /documents with tiny test .eml

**How to run**: Click "🔧 Run Diagnostics" button in taskpane (top of page)

**Location**: [diagnostics.ts](src/services/diagnostics.ts)

**Output**: Alert dialog + detailed console logs with pass/fail for each check

---

## D) How to Use the Diagnostics

### Method 1: Built-in Diagnostic Button (Recommended)
1. Open the add-in taskpane in Outlook
2. Click the **"🔧 Run Diagnostics"** button at the top
3. Follow any prompts (e.g., enter a test case ID)
4. Check the alert dialog for summary
5. Open browser console (F12) for detailed results

### Method 2: Manual Console Testing
Open browser console (F12) and run:

```javascript
// Import and run diagnostics
import("./services/diagnostics.js").then(async (mod) => {
  const results = await mod.runDiagnostics();
  console.log(mod.formatDiagnosticResults(results));
});
```

---

## E) Interpreting Results

### All checks pass ✅
- Auth and storage are working correctly
- Relay middleware is accessible
- API endpoints accept requests
- **Next**: Focus on OnMessageSend handler execution

### Token checks fail ❌
- Problem: Auth not properly stored or expired
- **Fix**: Re-login through the add-in
- **Verify**: `setAuth()` is called on successful login

### workspaceHost check fails ❌
- Problem: Workspace not selected or not saved
- **Fix**: Re-select workspace in settings
- **Verify**: Check workspace selection code calls `setStored()`

### GET /cases fails ❌
- Problem: Relay not working, or API authentication issue
- **Check**: Webpack dev server is running on `https://localhost:3000`
- **Check**: Token is valid (not expired)
- **Check**: Workspace host is correct

### POST /documents fails ❌
- Problem: API rejects document upload
- **Common causes**:
  - Invalid case ID
  - Malformed base64 data
  - Wrong MIME type
  - Insufficient permissions
- **Check**: Console logs show response status and error message
- **Fix**: Review error message and adjust accordingly

---

## F) Quick Debugging Workflow

1. **Start here**: Click "🔧 Run Diagnostics" button
2. **If diagnostics pass**: Problem is in OnMessageSend handler execution
3. **If diagnostics fail**: Fix the failing check first
4. **Test manual filing**: Click "Zařadit teď" and check console
5. **Test auto-filing**: Send an email in compose mode and check console
6. **Review logs**: All `[component]` prefixed logs in console show the flow

---

## G) Expected Console Output (Success Case)

### Manual Filing:
```
[doSubmit] About to upload email document {caseId: "123", ...}
[uploadDocumentToCase] Starting upload {caseId: "123", fileName: "email.eml", ...}
[getToken] Using sessionStorage token
[resolveApiBaseUrl] Resolved base URL: /singlecase/valfor-demo.singlecase.ch/publicapi/v1
[uploadDocumentToCase] Fetch completed {status: 200, ok: true}
[uploadDocumentToCase] Upload successful {documentIds: ["456"]}
[doSubmit] Email document created {response: {...}}
```

### Auto-filing (OnMessageSend):
```
[onMessageSendHandler] Handler fired
[onMessageSendHandler] Item keys: ["draft:abc123"]
[onMessageSendHandler] Intent: {caseId: "123", autoFileOnSend: true}
[onMessageSendHandler] Token retrieved {tokenPrefix: "eyJhbGciOi..."}
[onMessageSendHandler] Workspace host {normalized: "valfor-demo.singlecase.ch"}
[onMessageSendHandler] Uploading document to case {caseId: "123", fileName: "email.eml"}
[uploadDocumentToCase] Starting upload ...
[uploadDocumentToCase] Upload successful {documentIds: ["456"]}
[onMessageSendHandler] Upload successful
```

---

## H) Next Steps After Debugging

Once you identify the root cause:

1. **If relay issue**: Ensure commands.html loads from `https://localhost:3000`
2. **If token issue**: Fix auth mirroring in `setAuth()`
3. **If host issue**: Fix workspace selection persistence
4. **If MIME issue**: Standardize on `"message/rfc822"` for all .eml uploads
5. **If API issue**: Check SingleCase API documentation and adjust payload

---

## I) Critical Files Reference

| File | Purpose | Key Functions |
|------|---------|---------------|
| [singlecaseDocuments.ts](src/services/singlecaseDocuments.ts) | Document upload logic | `uploadDocumentToCase()`, `getToken()`, `resolveApiBaseUrl()` |
| [MainWorkspace.tsx](src/taskpane/components/MainWorkspace/MainWorkspace.tsx) | Manual filing UI | `doSubmit()` |
| [onMessageSendHandler.ts](src/commands/onMessageSendHandler.ts) | Auto-filing on send | `onMessageSendHandler()` |
| [auth.ts](src/services/auth.ts) | Auth storage | `getAuth()`, `getAuthRuntime()`, `setAuth()` |
| [storage.ts](src/utils/storage.ts) | Storage helpers | `getStored()`, `setStored()` |
| [diagnostics.ts](src/services/diagnostics.ts) | Self-test utility | `runDiagnostics()` |
| [webpack.config.js](webpack.config.js) | Dev server relay | Middleware at line 127 |

---

## J) Common Issues and Solutions

### Issue: "Network request failed"
- **Cause**: Commands runtime not using relay
- **Fix**: Verify fetch URL starts with `/singlecase/...` not `https://...`

### Issue: "Missing auth token"
- **Cause**: Token not in OfficeRuntime.storage
- **Fix**: Re-login, verify `setAuth()` calls `rtSet()`

### Issue: "Workspace host is missing"
- **Cause**: Host not stored
- **Fix**: Re-select workspace, verify `setStored()` is called

### Issue: "Upload failed (401)"
- **Cause**: Expired or invalid token
- **Fix**: Re-login

### Issue: "Upload failed (404)"
- **Cause**: Invalid case ID or API endpoint
- **Fix**: Verify case ID exists, check SingleCase API docs

### Issue: "Upload failed (415)"
- **Cause**: Wrong Content-Type or MIME type
- **Fix**: Ensure `Content-Type: application/json` header and `mime_type: "message/rfc822"`

### Issue: Upload succeeds but email not visible
- **Cause**: Wrong MIME type or document in unexpected location
- **Fix**: Change MIME type to `"message/rfc822"`, check document appears in case document list

---

## K) Testing Checklist

- [ ] Diagnostics all pass
- [ ] Manual filing works (Zařadit teď)
- [ ] Console shows detailed logs with no errors
- [ ] Document appears in SingleCase case
- [ ] OnMessageSend handler triggers (compose mode)
- [ ] Auto-filing on send works
- [ ] Outlook notification shows success message
- [ ] Both .eml and attachments upload (if selected)

---

## Contact & Support

If issues persist after following this debug plan:
1. Collect console logs from a failed attempt
2. Note which diagnostic checks fail
3. Include response status codes and error messages
4. Share DEBUG_PLAN.md results with your team
