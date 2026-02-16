# Debugging Outlook Desktop Send Handler

## The Problem
- ✅ OWA (browser) works fine
- ❌ Outlook Desktop doesn't file emails

## Known Differences Between OWA and Desktop

### 1. Storage Mechanism
- **OWA**: Uses `OfficeRuntime.storage` (shared across contexts)
- **Desktop**: May use `Office.context.roamingSettings` (sync issues possible)

### 2. ItemId Availability
- **OWA**: itemId often available immediately
- **Desktop**: May need `getItemIdAsync()` (which we use)

### 3. Logging
- **OWA**: Browser console (F12)
- **Desktop**: Runtime logs or add-in errors dialog

## How to Enable Logging in Desktop

### Windows Desktop
1. Close Outlook
2. Create file: `%TEMP%\OutlookLogging.txt` with content:
   ```
   [LogSettings]
   Level=verbose
   ```
3. Restart Outlook
4. Logs appear in: `%TEMP%\Outlook Logging\`

### Mac Desktop
1. Open Terminal
2. Run:
   ```bash
   defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true
   defaults write com.microsoft.Outlook EnableLogging -bool true
   ```
3. Restart Outlook
4. Right-click taskpane → Inspect Element

## Quick Test: Check if Handler is Firing

Add this to your email body before sending:
```
TEST-$(date +%s)
```

Then check logs for `[onMessageSendHandler] Handler fired`

## Debugging Checklist

### Step 1: Verify Intent is Saved (Taskpane Side)
1. Open taskpane in Desktop
2. Select a case
3. Enable auto-file
4. Open browser debugger (if Mac) or check logs
5. Look for:
   ```
   [saveComposeIntent] Saving intent
   [saveComposeIntent] Saved to roamingSettings (Desktop uses this, not OfficeRuntime)
   [saveComposeIntent] Saved fallback to roamingSettings
   ```

### Step 2: Verify Handler Fires (Send Side)
1. Compose email with test subject
2. Click Send
3. Check logs for:
   ```
   [onMessageSendHandler] Handler fired
   [onMessageSendHandler] Platform info
   ```

### Step 3: Verify Intent is Found
Look for:
```
[onMessageSendHandler] Reading intent from storage
[readIntentAny] Trying key: sc_intent:draft:current
[readIntentAny] Found in roamingSettings
[onMessageSendHandler] Intent: {caseId: "...", found: true}
```

### Step 4: Verify Version Decision
Look for:
```
[onMessageSendHandler] Version decision {
  hasConversation: false,
  shouldUploadVersion: false
}
```

### Step 5: Verify Upload
Look for:
```
[onMessageSendHandler] Upload successful
```

## Common Desktop Issues

### Issue 1: roamingSettings Not Synchronized
**Symptom**: Intent saved but not found in handler

**Possible Cause**: Desktop roamingSettings sync delay

**Fix**: Add explicit sync wait in saveComposeIntent (see below)

### Issue 2: Cross-Context Storage Access
**Symptom**: Different storage contexts between taskpane and runtime

**Possible Cause**: Desktop isolates contexts more strictly

**Fix**: Use both roamingSettings AND localStorage as fallback

### Issue 3: Handler Not Firing
**Symptom**: No logs at all when sending

**Possible Cause**: Manifest not properly configured for Desktop

**Fix**: Verify manifest.xml has correct runtime configuration

## What to Share for Debugging

Please share:

1. **Platform Info** (from first log):
   ```
   hasOfficeRuntime: true/false
   hasOfficeRuntimeStorage: true/false
   hasRoamingSettings: true/false
   host: "Outlook"
   hostVersion: "16.x.x"
   ```

2. **Save Logs** (when selecting case):
   ```
   [saveComposeIntent] Saved to: (which storage?)
   ```

3. **Handler Logs** (when clicking send):
   ```
   [onMessageSendHandler] Item keys: [...]
   [onMessageSendHandler] Intent: {...}
   [onMessageSendHandler] Version decision: {...}
   ```

4. **Any Errors**:
   ```
   [onMessageSendHandler] Error during filing: ...
   ```

## Potential Fixes to Try

### If Intent Not Found in Desktop:

1. **Add Storage Sync Delay** (in saveComposeIntent after roamingSettings.saveAsync):
   ```typescript
   await new Promise(resolve => setTimeout(resolve, 500));
   ```

2. **Try SharedFolders** (alternative to roamingSettings):
   Desktop has `Office.context.mailbox.item.itemId` available earlier

3. **Add Multiple Storage Layers**:
   - Primary: roamingSettings
   - Fallback: localStorage
   - Last resort: Write to item customProperties

### If Handler Not Firing:

Check manifest.xml `<Runtime>` element is configured for Desktop
