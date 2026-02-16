# Desktop Outlook Debug Guide

## Problem
The filed email cache works reliably in web Outlook but fails intermittently on desktop Outlook. DevTools closes when you send an email, preventing us from seeing the cache write logs.

## Changes Made

### 1. Enhanced Storage with Retry Logic
**File:** `src/utils/storage.ts`

**Changes:**
- Added 100ms delay after `saveRoamingSettings()` to ensure operation completes before window closes
- Added retry logic with exponential backoff (up to 2 retries: 200ms, 400ms delays)
- Added persistent debug logging functions: `getDebugLog()`, `clearDebugLog()`
- Enhanced logging to show storage backend, character counts, and timing

**Why:** Desktop Outlook may close the compose window before async `saveAsync` completes, causing cache writes to fail silently.

### 2. Console Debug Commands
**File:** `src/utils/debugCache.ts` (NEW)

**Available Commands:**
```javascript
// Show all cached emails
scDebug.showCache()

// Show platform information
scDebug.showPlatform()

// Search cache by email subject
scDebug.searchBySubject("test email subject")

// Show persistent debug log
scDebug.showDebugLog()

// Clear debug log
scDebug.clearDebugLog()

// Show help
scDebug.help()
```

**Why:** Since DevTools closes on send, you need a way to inspect cache contents AFTER the send operation. These commands persist across page reloads.

### 3. Platform Detection Logging
**Files:** `src/utils/filedCache.ts`, `src/taskpane/components/MainWorkspace/MainWorkspace.tsx`

**Added:** Platform detection logs showing:
- Host name (Outlook vs OutlookWebApp)
- Host version
- Platform type
- Storage backend being used

## How to Debug Desktop Outlook

### Step 1: Test File Operation
1. Open **desktop Outlook**
2. Open DevTools (F12 or right-click > Inspect)
3. Compose a new email to yourself
4. Select a case and enable auto-file
5. Write subject (e.g., "desktop test")
6. **Click Send**
   - DevTools will close - this is expected
   - The cache write happens during this period

### Step 2: Check Cache Contents
1. **Open any email** (inbox, sent items, doesn't matter)
2. Open DevTools again (F12)
3. In the Console tab, type:
   ```javascript
   scDebug.showCache()
   ```
4. Check if your email was cached:
   - Look for a row with `Type: "Subject"`
   - Subject should be `"desktop test"`
   - If present: ✅ Cache write succeeded
   - If absent: ❌ Cache write failed

### Step 3: Check Platform Info
```javascript
scDebug.showPlatform()
```

Compare with web Outlook:
- Desktop: `host: "Outlook"`, `hostVersion: "16.0.xxxxx"`
- Web: `host: "OutlookWebApp"`, `hostVersion: "..."`
- Check which storage backend is being used

### Step 4: Open Received Email
1. Open the email you sent to yourself
2. Open DevTools
3. Check console logs for:
   ```
   [findFiledEmailBySubject] Platform info: {...}
   [findFiledEmailBySubject] Cache size: X keys
   [findFiledEmailBySubject] ✅ Found cache entry by subject
   ```

### Step 5: Search by Subject
If the email isn't detected, search manually:
```javascript
scDebug.searchBySubject("desktop test")
```

This will show:
- Whether the entry exists in cache
- The exact cache key being used
- List of all available subject keys

## Expected Log Patterns

### Successful Cache Write (at send time):
```
[cacheFiledEmailBySubject] Platform info: {host: "Outlook", ...}
[setStored] Using storage backend: roamingSettings
[setStored] Writing to roamingSettings...
[saveRoamingSettings] Starting saveAsync...
[saveRoamingSettings] ✅ Succeeded in 45ms
[setStored] ✅ saveAsync completed
[cacheFiledEmailBySubject] Write verification: {success: true, cacheSize: 5}
```

### Successful Cache Read (at read time):
```
[findFiledEmailBySubject] Platform info: {host: "Outlook", ...}
[getStored] Using storage backend: roamingSettings
[getStored] Got from roamingSettings: found (2847 chars)
[findFiledEmailBySubject] Cache size: 5 keys
[findFiledEmailBySubject] ✅ Found cache entry by subject
```

### Failed Cache Write (what we're debugging):
```
[cacheFiledEmailBySubject] Platform info: {host: "Outlook", ...}
[setStored] Using storage backend: roamingSettings
[setStored] Writing to roamingSettings...
[saveRoamingSettings] Starting saveAsync...
[saveRoamingSettings] ❌ Failed in 120ms: [error message]
```

OR silence (if window closes before logging):
```
[cacheFiledEmailBySubject] Platform info: {host: "Outlook", ...}
[setStored] Using storage backend: roamingSettings
[setStored] Writing to roamingSettings...
(no more logs - window closed)
```

## Troubleshooting

### Cache is Empty
**Possible causes:**
1. **Window closed too fast** - The 100ms delay may not be enough
   - Solution: Increase delay in `saveRoamingSettings()`
2. **saveAsync is failing silently** - No error thrown but save not persisting
   - Solution: Add more aggressive retry logic
3. **Different storage backend** - Desktop using different API than web
   - Check: `scDebug.showPlatform()` and compare storage backends

### Cache Has Entry But Not Detected
**Possible causes:**
1. **Subject mismatch** - Check exact subject in cache vs current email
   - Use: `scDebug.searchBySubject("exact subject")`
2. **Case sensitivity** - Subject keys are lowercased, check normalization
3. **Whitespace differences** - Leading/trailing spaces

### Cache Entry Exists But Wrong Data
**Possible causes:**
1. **Multiple sends** - Old entry not updated
2. **Race condition** - Two writes happening simultaneously

## Comparison: Web vs Desktop

| Aspect | Web Outlook | Desktop Outlook |
|--------|-------------|-----------------|
| Host | `OutlookWebApp` | `Outlook` |
| Storage API | roamingSettings | roamingSettings |
| Window Behavior | Stays open after send | Closes immediately after send |
| saveAsync Timing | ~30-50ms | ~80-150ms (?) |
| DevTools Persistence | Stays open | Closes with compose window |

## Next Steps Based on Findings

### If cache is empty on desktop:
→ Increase delay after saveAsync (try 200ms, 500ms)
→ Add more retries
→ Consider alternative timing (post-send hook)

### If cache exists but not detected:
→ Subject normalization issue
→ Check conversationId availability timing

### If saveAsync fails:
→ Permissions issue with roaming settings on desktop
→ Try localStorage fallback immediately
→ Check desktop Outlook version/settings

## Quick Test Commands

```javascript
// Full diagnostic workflow
scDebug.showPlatform()
scDebug.showCache()
scDebug.searchBySubject("your email subject")

// If you see errors, check debug log
scDebug.showDebugLog()
```

## Success Criteria

✅ After sending email on desktop:
- `scDebug.showCache()` shows the email entry
- Cache size increases by 1
- Subject matches exactly

✅ After opening received email on desktop:
- Console logs show "✅ Found cache entry by subject"
- UI shows "Already filed in: [Case Name]"
- Category "SC: Zařazeno" is applied

---

**Status:** Ready for testing on desktop Outlook
**Date:** 2026-02-13
**Key Improvement:** Added 100ms delay + retry logic + console debug commands
