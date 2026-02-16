# Quick Start Guide - SingleCase Outlook Add-in

This guide helps you set up the development environment for the SingleCase Outlook Add-in on your local machine.

## Prerequisites

- **Node.js**: Version 18.x or higher ([Download](https://nodejs.org/))
- **npm**: Version 9.x or higher (comes with Node.js)
- **Outlook Desktop**: Microsoft 365 or Outlook 2019/2021 (Windows or Mac)
- **Microsoft 365 Account**: Required for development and testing

## Initial Setup

### 1. Install Dependencies

```bash
npm install
```

This installs all required packages including:
- Office.js runtime
- React and TypeScript
- Webpack dev server
- Development certificates

### 2. Trust Development Certificates

The add-in uses HTTPS for local development. You **must** trust the development certificate.

#### On macOS:
```bash
npx office-addin-dev-certs install
```

If prompted, enter your password to trust the certificate.

#### On Windows:
```bash
npx office-addin-dev-certs install
```

Run PowerShell **as Administrator** if certificate installation fails.

**Verify Certificate Installation:**
```bash
npx office-addin-dev-certs verify
```

You should see: ✅ `The developer certificates are installed and trusted.`

### 3. Start the Development Server

```bash
npm start
```

This starts the webpack dev server on `https://localhost:3000`.

**Expected output:**
```
Compiled successfully!

You can now view the add-in in Outlook.
Webpack dev server running at https://localhost:3000
```

**⚠️ Keep this terminal window open** - the server must run continuously.

### 4. Sideload the Manifest

The manifest tells Outlook where to find your add-in and what permissions it needs.

#### On Windows (Outlook Desktop):

1. **Open Outlook Desktop**
2. Click **File** → **Get Add-ins**
3. Click **My add-ins** (left sidebar)
4. Scroll down to **Custom add-ins**
5. Click **+ Add a custom add-in** → **Add from File...**
6. Browse to: `[project-root]/manifest.xml`
7. Click **Install**
8. Confirm all permission prompts

#### On macOS (Outlook Desktop):

1. **Open Outlook Desktop**
2. Click **Tools** → **Get Add-ins**
3. Click **My add-ins** tab
4. Scroll to **Custom Add-ins**
5. Click **+ Add custom add-in** → **Add from file...**
6. Select: `[project-root]/manifest.xml`
7. Click **Install**

#### Verify Sideload Success:

1. Compose a new email
2. Look for **SingleCase** button in the ribbon
3. Click it - the taskpane should open showing the add-in UI

---

## Critical: Verify Event-Based Runtime

The add-in uses **event-based activation** for On Send and categories. This requires special runtime setup.

### Check Runtime Registration

#### On Windows:
1. Open **Task Manager** (Ctrl+Shift+Esc)
2. Go to **Details** tab
3. While Outlook is running, look for: `OfficeClickToRun.exe` or `AppVShNotify.exe`

#### On macOS:
1. Open **Activity Monitor**
2. Search for process: `Microsoft Outlook` with multiple instances

### Verify Manifest Runtime Configuration

Open `manifest.xml` and verify these sections exist:

```xml
<Runtimes>
  <Runtime resid="Taskpane.Url" lifetime="short">
    <Override type="javascript" resid="Commands.Url" />
  </Runtime>
</Runtimes>

<LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" />
```

If missing or incorrect, the On Send event will **never fire**.

---

## Verify Office.js Requirement Sets

The add-in requires specific Office.js capabilities:

### Check Your Outlook Version:

1. **Help** → **About Outlook**
2. Note the version number (e.g., `16.0.16731.20170`)

### Required Capability Sets:

From `manifest.xml`:
```xml
<Requirements>
  <Sets DefaultMinVersion="1.13">
    <Set Name="Mailbox" MinVersion="1.13" />
  </Sets>
</Requirements>
```

**Minimum versions:**
- **Windows**: Outlook 2021 or Microsoft 365 (Version 16.0.14326+)
- **macOS**: Outlook 2021 or Microsoft 365 (Version 16.54+)

If your Outlook version is older, upgrade to Microsoft 365.

---

## Verify Graph API Permissions

The add-in uses Microsoft Graph for:
- Reading categories
- Managing email properties

### Check Graph Token:

1. Open browser DevTools (F12) when taskpane is open
2. Run in console:
   ```javascript
   OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true })
   ```
3. Should return a JWT token (long string starting with `eyJ...`)

### Common Graph Issues:

**Error: `consent_required`**
- **Fix**: User must consent to Graph permissions
- Go to: https://portal.azure.com → Azure AD → Enterprise Applications → Find your add-in → Permissions

**Error: `AADSTS50011`**
- **Fix**: Reply URL mismatch
- Ensure `https://localhost:3000` is registered in Azure AD

---

## Clearing Outlook Cache (CRITICAL)

Outlook **aggressively caches** add-in manifests and runtimes. If you make changes, you **must** clear cache.

### On Windows:

1. **Close Outlook completely**
2. Open **Run** (Win+R)
3. Type: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
4. **Delete everything** in this folder
5. Also delete: `%TEMP%\Outlook Logging\*`
6. Restart Outlook

### On macOS:

1. **Quit Outlook** (Cmd+Q)
2. Open **Terminal**
3. Run:
   ```bash
   rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/*
   rm -rf ~/Library/Caches/com.microsoft.Outlook/*
   ```
4. Restart Outlook

### Force Manifest Reload:

After clearing cache, you **must** re-sideload the manifest:
1. Remove the add-in: **My Add-ins** → Click **...** → **Remove**
2. Re-sideload following steps in **Section 4**

---

## Testing the Add-in

### 1. Test Taskpane

1. Compose new email
2. Click **SingleCase** button
3. Taskpane should open
4. Check browser console (F12) for errors

### 2. Test BCC Reading

1. Compose new email
2. Add recipient to **BCC**: `2023-0006@valfor-demo.singlecase.ch`
3. Open taskpane
4. **Expected**: Add-in detects case from BCC
5. Check console logs: `[submail-detection] Found case keys in recipients`

### 3. Test Categories

1. Open any received email
2. Open taskpane
3. File the email to a case
4. **Expected**: Email gets category "SC: Zařazeno"
5. Verify in Outlook: Email should have colored category

### 4. Test On Send Event

1. Compose new email
2. Select a case in taskpane
3. Enable "Auto-file on send"
4. Click **Send**
5. **Expected**:
   - Send is delayed briefly
   - Toast notification: "SingleCase: email uložen při odeslání."
   - Email files to SingleCase automatically

---

## Troubleshooting

### 🔴 Issue: On Send Event Not Firing

**Symptoms:**
- Click Send → email sends immediately
- No delay, no toast notification
- Email doesn't file automatically

**Diagnosis:**
1. Check if event runtime is registered:
   ```bash
   # On Windows PowerShell:
   Get-Process | Where-Object {$_.ProcessName -like "*Office*"}
   ```
2. Look for `OfficeClickToRun.exe` or runtime process

**Fixes:**

#### Fix 1: Clear Outlook Cache
```bash
# Windows:
del /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*"

# macOS:
rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/*
```

#### Fix 2: Verify Manifest Runtime
Open `manifest.xml` and ensure:
```xml
<Runtimes>
  <Runtime resid="Taskpane.Url" lifetime="short">
    <Override type="javascript" resid="Commands.Url" />
  </Runtime>
</Runtimes>
```

The `Commands.Url` must point to: `https://localhost:3000/commands.html`

#### Fix 3: Check Commands HTML
Verify `public/commands.html` exists and contains:
```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
<script src="commands.js"></script>
```

#### Fix 4: Reinstall Manifest
1. Remove add-in completely
2. Close Outlook
3. Clear cache (see above)
4. Restart Outlook
5. Re-sideload manifest

#### Fix 5: Check Outlook Version
On Send requires **Mailbox 1.13** minimum.

Run in browser console (with taskpane open):
```javascript
console.log(Office.context.requirements.isSetSupported('Mailbox', '1.13'))
```

Should return `true`. If `false`, upgrade Outlook.

---

### 🔴 Issue: Categories Not Working

**Symptoms:**
- Email files successfully
- But no category appears in Outlook
- Or category appears but is wrong color

**Diagnosis:**
1. Open email
2. Open browser console (F12)
3. Run:
   ```javascript
   Office.context.mailbox.item.categories.getAsync((result) => {
     console.log('Categories API:', result.status);
     console.log('Current categories:', result.value);
   });
   ```

**Fixes:**

#### Fix 1: Check Categories API Support
```javascript
// In browser console:
console.log(Office.context.mailbox.item.categories ? 'Categories API available' : 'Categories API NOT available');
```

If `NOT available`, your Outlook version doesn't support it.

#### Fix 2: Create Master Categories Manually
1. In Outlook, go to **Home** → **Categorize** → **All Categories**
2. Click **New**
3. Create category: `SC: Zařazeno` with green color
4. Create category: `SC: Nezařazeno` with gray color
5. Try filing again

#### Fix 3: Use Graph Fallback
The add-in should fallback to Graph API if Office.js categories fail.

Check console for:
```
[graphMail] Applying category via Graph API
```

If missing, check Graph permissions (see section above).

#### Fix 4: Verify Category Names
The add-in uses these exact names:
```typescript
const CATEGORY_NAME = "SC: Zařazeno";      // Filed
const CATEGORY_NAME_UNFILED = "SC: Nezařazeno";  // Unfiled
```

If your Outlook has different category names, they won't match.

---

### 🔴 Issue: BCC Not Detected

**Symptoms:**
- Add BCC with case submail (e.g., `2023-0006@valfor-demo.singlecase.ch`)
- Taskpane doesn't show "Spis detekován z BCC"
- Wrong case is selected instead

**Diagnosis:**
Open browser console and check:
```javascript
Office.context.mailbox.item.bcc.getAsync((result) => {
  console.log('BCC support:', result.status);
  console.log('BCC recipients:', result.value);
});
```

**Fixes:**

#### Fix 1: Check BCC API Support
BCC reading requires **Mailbox 1.8** minimum.

```javascript
console.log(Office.context.requirements.isSetSupported('Mailbox', '1.8'))
```

If `false`, upgrade Outlook to Microsoft 365.

#### Fix 2: Verify Workspace Host
The add-in matches BCC against workspace domain.

1. Open taskpane
2. Check console for:
   ```
   [submail-detection] Checking recipients for submail
   workspaceHost: "valfor-demo.singlecase.ch"
   ```
3. If `workspaceHost` is empty or wrong, sign out and sign in again

#### Fix 3: Check Case Name Format
The add-in looks for case key in parentheses:
```
"Internal Know How (2023-0006)"
```

Console should show:
```
[resolveSubmailCaseKey] Resolved case: {caseKey: "2023-0006", caseId: "123", caseName: "..."}
```

If you see:
```
[resolveSubmailCaseKey] No case found for key: 2023-0006
```

Then case names don't match the expected format.

#### Fix 4: Check Recipients Polling
BCC detection runs every 350ms while composing.

Check console for repeated logs:
```
[submail-detection] Checking recipients for submail
```

If missing, recipient polling isn't working.

---

### 🔴 Issue: Event Runtime Not Loading

**Symptoms:**
- Taskpane works fine
- But On Send doesn't fire
- Console shows no runtime logs

**Diagnosis:**
1. Open `https://localhost:3000/commands.html` directly in browser
2. Should see blank page with no errors
3. Open console - should see:
   ```
   Office.js loaded
   Commands runtime initialized
   ```

**Fixes:**

#### Fix 1: Check Commands Entry Point
Verify `webpack.config.js` has entry:
```javascript
entry: {
  commands: './src/commands/commands.ts',
  // ...
}
```

#### Fix 2: Verify Commands Initialization
Open `src/commands/commands.ts` and ensure:
```typescript
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
```

The function name **must** match the manifest:
```xml
<LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" />
```

#### Fix 3: Check for Runtime Errors
Open `https://localhost:3000/commands.html` in browser.

Check console for errors:
- ❌ `Office is not defined` → Office.js not loaded
- ❌ `onMessageSendHandler is not defined` → Function not exported
- ❌ `actions.associate is not a function` → Wrong Office.js version

#### Fix 4: Use Runtime Logging (Windows)
1. Create file: `%TEMP%\OutlookLogging.txt`
2. Add content:
   ```
   [LogSettings]
   Level=verbose
   ```
3. Restart Outlook
4. Compose email, click Send
5. Check logs: `%TEMP%\Outlook Logging\`

Look for:
```
[Runtime] Loading event handler: onMessageSendHandler
[Runtime] Handler executed successfully
```

---

## Environment Differences Checklist

If add-in works on one machine but not another, check:

- [ ] **Node.js version** matches (`node --version`)
- [ ] **npm dependencies** installed (`npm install` clean)
- [ ] **Dev certificates** trusted (`npx office-addin-dev-certs verify`)
- [ ] **Outlook version** supports Mailbox 1.13+ (`Help → About`)
- [ ] **Manifest** sideloaded in same Outlook profile
- [ ] **Cache cleared** completely (see section above)
- [ ] **Firewall** allows `localhost:3000` (Windows Defender / macOS Firewall)
- [ ] **Antivirus** not blocking webpack dev server
- [ ] **Microsoft 365 account** signed into Outlook (not local account)
- [ ] **Same workspace** configured in add-in (check console logs)

---

## Advanced Debugging

### Enable Verbose Logging

Add to `src/commands/onMessageSendHandler.ts` (top of function):
```typescript
export async function onMessageSendHandler(event: SendEvent) {
  console.log("=== ON SEND FIRED ===");
  console.log("Event:", event);
  console.log("Office version:", Office.context.diagnostics);
  // ... rest of handler
}
```

### Inspect Manifest at Runtime

```javascript
// In browser console (taskpane):
Office.context.manifest
```

Should show parsed manifest object. If `undefined`, manifest not loaded.

### Check Network Requests

1. Open **Network** tab in browser DevTools
2. Filter by `localhost:3000`
3. All resources should return `200 OK`
4. If any `404 Not Found`, check webpack output

### Test Without Cache

Add `?nocache=true` to manifest URLs:
```xml
<bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html?nocache=true" />
```

Forces Outlook to reload from server every time.

---

## Getting Help

If issues persist after following this guide:

1. **Check logs** (`console.log` in browser DevTools)
2. **Check Outlook logs** (Windows: `%TEMP%\Outlook Logging\`)
3. **Compare working vs non-working**:
   - Outlook version
   - Node.js version
   - Certificate trust status
4. **Report issue** with:
   - Exact steps to reproduce
   - Console logs
   - Outlook version
   - OS version

---

## Quick Reference

### Start Development
```bash
npm start
# Keep running, open https://localhost:3000 in browser to verify
```

### Clear Cache & Reload
```bash
# Windows:
del /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*"

# macOS:
rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/*
```

Then: Remove add-in → Restart Outlook → Re-sideload manifest

### Verify Runtime
```bash
# Check certificates:
npx office-addin-dev-certs verify

# Check Office.js support:
# (in browser console with taskpane open)
Office.context.requirements.isSetSupported('Mailbox', '1.13')
```

### Test BCC Detection
1. Compose email
2. BCC: `2023-0006@your-workspace.singlecase.ch`
3. Console should show: `[submail-detection] Found case keys`

### Test On Send
1. Select case + enable auto-file
2. Click Send
3. Should see: Toast "SingleCase: email uložen při odeslání."

---

**🎉 You're ready to develop!** If everything works, you should see:
- ✅ Taskpane opens
- ✅ BCC detects cases
- ✅ Categories apply
- ✅ On Send files emails automatically

If any step fails, refer to the **Troubleshooting** section above.
