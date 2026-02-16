# Troubleshooting - Quick Checklist

Use this when the add-in works on one machine but not another.

## 🔥 Critical: The Three Most Common Issues

### 1. Outlook Cache Not Cleared

**Problem**: Old manifest/runtime cached
**Fix**:
```bash
# Windows:
del /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*"
del /s /q "%TEMP%\Outlook Logging\*"

# macOS:
rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/*
rm -rf ~/Library/Caches/com.microsoft.Outlook/*
```

Then:
1. **Close Outlook completely**
2. **Remove add-in**: My Add-ins → ... → Remove
3. **Restart Outlook**
4. **Re-sideload manifest**

### 2. Dev Certificates Not Trusted

**Problem**: HTTPS certificate not trusted
**Fix**:
```bash
npx office-addin-dev-certs install
npx office-addin-dev-certs verify
```

Should say: ✅ `The developer certificates are installed and trusted.`

If not, run as Administrator (Windows) or enter password (macOS).

### 3. Wrong Outlook Version

**Problem**: Outlook version too old
**Check**: Help → About Outlook
**Required**:
- Windows: 16.0.14326 or newer
- macOS: 16.54 or newer

Upgrade to **Microsoft 365** if version is lower.

---

## 🔴 On Send Event Not Firing

### Quick Fixes (Try in order):

#### 1. Verify Runtime is Registered
Open `manifest.xml` and confirm:
```xml
<Runtimes>
  <Runtime resid="Taskpane.Url" lifetime="short">
    <Override type="javascript" resid="Commands.Url" />
  </Runtime>
</Runtimes>

<LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" />
```

#### 2. Check Commands.html Loads
Open in browser: `https://localhost:3000/commands.html`

Should see **blank page** with console showing:
```
Office.js loaded
```

If you see errors or 404, webpack isn't building commands entry.

#### 3. Verify Function Registration
In `src/commands/commands.ts`:
```typescript
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
```

Function name **must match** manifest exactly.

#### 4. Test Mailbox API Support
Open taskpane → Browser console:
```javascript
Office.context.requirements.isSetSupported('Mailbox', '1.13')
// Should return: true
```

If `false`: **Outlook version too old** → Upgrade to Microsoft 365

#### 5. Nuclear Option: Full Reset
```bash
# 1. Stop npm start
# 2. Clear all caches (see above)
# 3. Reinstall:
npm install
npx office-addin-dev-certs install
npm start
# 4. Close Outlook
# 5. Restart Outlook
# 6. Re-sideload manifest
```

---

## 🔴 Categories Not Working

### Quick Fixes (Try in order):

#### 1. Check Categories API Available
Taskpane console:
```javascript
Office.context.mailbox.item.categories.getAsync((result) => {
  console.log(result.status, result.value);
});
```

Expected: `succeeded` + array of categories

If **error** or `undefined`: Categories API not supported → Use Graph fallback (automatic)

#### 2. Verify Master Categories Exist
In Outlook:
1. **Home** → **Categorize** → **All Categories**
2. Check for: `SC: Zařazeno` (green) and `SC: Nezařazeno` (gray)
3. If missing, click **New** and create them manually

#### 3. Check Graph API Works
Taskpane console:
```javascript
OfficeRuntime.auth.getAccessToken({
  allowSignInPrompt: true,
  forMSGraphAccess: true
})
.then(token => console.log('Graph token:', token.substring(0, 20) + '...'))
.catch(err => console.error('Graph error:', err))
```

Expected: Token string starting with `eyJ...`

If **error**: Graph permissions not configured → User must consent

#### 4. Check Execution
File an email and check console for:
```
[graphMail] applyFiledCategoryToCurrentEmailOfficeJs
```

If missing: Category function not called → Check filing flow

---

## 🔴 BCC Not Detected

### Quick Fixes (Try in order):

#### 1. Check BCC API Support
Taskpane console (in compose mode):
```javascript
Office.context.mailbox.item.bcc.getAsync((result) => {
  console.log('BCC status:', result.status);
  console.log('BCC recipients:', result.value);
});
```

Expected: `succeeded` + array of recipients

If **error**: BCC API not supported → Requires Mailbox 1.8+ → Upgrade Outlook

#### 2. Verify Recipients Polling Works
Add BCC, then check console for **repeated logs**:
```
[submail-detection] Checking recipients for submail
recipientCount: 1
workspaceHost: "valfor-demo.singlecase.ch"
```

Should appear every ~350ms. If missing: Polling not running.

#### 3. Check Workspace Host
Console should show:
```
workspaceHost: "valfor-demo.singlecase.ch"
```

If **empty or wrong**:
1. Sign out of add-in
2. Clear browser data
3. Sign in again
4. Check workspace URL in settings

#### 4. Verify Case Name Format
Add BCC: `2023-0006@valfor-demo.singlecase.ch`

Console should show:
```
[extractSubmailCaseKeys] Found submail case key: "2023-0006"
[resolveSubmailCaseKey] Resolved case: {caseKey: "2023-0006", caseName: "..."}
```

If you see:
```
[resolveSubmailCaseKey] No case found for key: 2023-0006
```

**Problem**: Case name format doesn't match.
**Expected format**: `"Case Name (2023-0006)"`
**Fix**: Ensure case names in SingleCase end with `(case-key)`

#### 5. Check Dash/Period Normalization
For case keys like `2023-0005-001` vs `2023-0005.001`:

Console should show:
```
[normalizeCaseKey] Normalized: "2023.0005.001"
```

Both formats should match after normalization.

---

## 🔴 Event Runtime Not Loading

### Quick Fixes:

#### 1. Check Webpack Entry Points
Verify `webpack.config.js` has:
```javascript
entry: {
  taskpane: './src/taskpane/index.tsx',
  commands: './src/commands/commands.ts',  // ← CRITICAL
},
```

If missing `commands` entry: Add it and restart `npm start`

#### 2. Verify Commands Build Output
After `npm start`, check:
```
http://localhost:3000/commands.html   ← Should exist (200 OK)
http://localhost:3000/commands.js     ← Should exist (200 OK)
```

If **404**: Webpack config wrong or build failed.

#### 3. Check Manifest URLs
In `manifest.xml`:
```xml
<bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
```

Must point to **exact URL** where commands.html is served.

#### 4. Test Runtime Directly
Open: `https://localhost:3000/commands.html`

Open browser console. Should see:
```
Office.initialize called
Commands runtime ready
```

Any errors here mean runtime can't load properly.

---

## 📋 Pre-Flight Checklist

Before asking for help, verify:

### ✅ Environment Setup
- [ ] Node.js 18+ installed (`node --version`)
- [ ] Dependencies installed (`npm install` completed without errors)
- [ ] Dev certificates trusted (`npx office-addin-dev-certs verify` = ✅)
- [ ] Dev server running (`npm start` shows "Compiled successfully")
- [ ] Server accessible (`https://localhost:3000` loads in browser)

### ✅ Outlook Configuration
- [ ] Outlook version 16.0.14326+ (Help → About)
- [ ] Signed into **Microsoft 365** account (not local)
- [ ] Manifest sideloaded (SingleCase button visible in ribbon)
- [ ] Cache cleared (deleted Wef folder)
- [ ] Outlook restarted after cache clear

### ✅ Runtime Verification
- [ ] Taskpane opens (`https://localhost:3000/taskpane.html`)
- [ ] Commands runtime loads (`https://localhost:3000/commands.html`)
- [ ] Mailbox 1.13 supported (console: `Office.context.requirements.isSetSupported('Mailbox', '1.13')` = `true`)
- [ ] No certificate warnings in browser console

### ✅ Feature-Specific
**For On Send:**
- [ ] `<Runtimes>` section exists in manifest
- [ ] `<LaunchEvent>` registered with correct function name
- [ ] Console shows: `[onMessageSendHandler] Handler fired` when clicking Send

**For Categories:**
- [ ] Master categories exist in Outlook (`SC: Zařazeno`, `SC: Nezařazeno`)
- [ ] Graph token obtainable (console test above)
- [ ] Console shows category application logs

**For BCC:**
- [ ] BCC API available (console test above)
- [ ] Workspace host configured correctly
- [ ] Console shows recipient polling logs
- [ ] Case names match format: `Name (case-key)`

---

## 🔍 Diagnostic Commands

Run these in browser console (with taskpane open):

```javascript
// 1. Check Office.js version
console.log('Office version:', Office.context.diagnostics.version);
console.log('Host:', Office.context.diagnostics.host);

// 2. Check requirements support
console.log('Mailbox 1.13:', Office.context.requirements.isSetSupported('Mailbox', '1.13'));
console.log('Mailbox 1.8:', Office.context.requirements.isSetSupported('Mailbox', '1.8'));

// 3. Check current item
console.log('Item type:', Office.context.mailbox.item.itemType);
console.log('Item class:', Office.context.mailbox.item.itemClass);

// 4. Check categories API
Office.context.mailbox.item.categories.getAsync((result) => {
  console.log('Categories:', result.status, result.value);
});

// 5. Check BCC API (compose only)
Office.context.mailbox.item.bcc.getAsync((result) => {
  console.log('BCC:', result.status, result.value);
});

// 6. Check Graph token
OfficeRuntime.auth.getAccessToken({
  allowSignInPrompt: true,
  forMSGraphAccess: true
})
.then(token => console.log('Graph token obtained:', token.substring(0, 30) + '...'))
.catch(err => console.error('Graph error:', err));
```

---

## 🆘 Still Not Working?

### Compare Working vs Non-Working Machine

Create comparison checklist:

| Check | Working Machine | Non-Working Machine |
|-------|----------------|---------------------|
| Node.js version | _____________ | _____________ |
| npm version | _____________ | _____________ |
| Outlook version | _____________ | _____________ |
| OS version | _____________ | _____________ |
| Certificate status | ✅ / ❌ | ✅ / ❌ |
| Cache cleared | ✅ / ❌ | ✅ / ❌ |
| Mailbox 1.13 support | ✅ / ❌ | ✅ / ❌ |
| Commands.html loads | ✅ / ❌ | ✅ / ❌ |
| Graph token works | ✅ / ❌ | ✅ / ❌ |

### Collect Debug Info

If issue persists, collect:

1. **Console logs** (Full output from browser DevTools)
2. **Network tab** (Check for 404s or cert errors)
3. **Outlook version** (Help → About → Full version number)
4. **Manifest version** (Check `<Version>` in manifest.xml)
5. **Webpack output** (From terminal running `npm start`)

### Known Platform Issues

**Windows Specific:**
- Windows Defender may block localhost:3000 → Add firewall exception
- Antivirus (Avast, Norton) may block Office.js → Whitelist Outlook.exe
- Corporate proxy may block certificate trust → Contact IT

**macOS Specific:**
- macOS Gatekeeper may block add-in → System Preferences → Security → Allow
- Outlook must be downloaded from Microsoft (not App Store version)
- FileVault encryption may slow cache clearing → Restart required

---

## Quick Win: Nuclear Reset

If all else fails, complete reset:

```bash
# 1. Stop everything
# Kill npm start (Ctrl+C)
# Close Outlook completely

# 2. Clean project
rm -rf node_modules
rm -rf dist
npm install

# 3. Clear Outlook cache
# (Use commands from top of this file)

# 4. Reinstall certificates
npx office-addin-dev-certs install --force

# 5. Restart
npm start
# Open Outlook
# Re-sideload manifest

# 6. Test
# Compose email → Check if SingleCase button appears
```

This should resolve 90% of environment-related issues.
