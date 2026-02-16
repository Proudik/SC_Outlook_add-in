# Developer Setup Guide

Welcome to the SingleCase Outlook Add-in project! This guide will get you up and running.

## 📚 Documentation

- **[QUICK_START.md](./QUICK_START.md)** - Complete setup guide for new developers
- **[TROUBLESHOOTING.md](./TROUBLESHOOTING.md)** - Fix common issues (On Send, Categories, BCC)
- **[DEBUG_DESKTOP.md](./DEBUG_DESKTOP.md)** - Debugging Outlook Desktop specifically

## 🚀 Quick Setup (5 minutes)

```bash
# 1. Install dependencies
npm install

# 2. Trust dev certificates
npx office-addin-dev-certs install

# 3. Start dev server
npm start

# 4. Sideload manifest in Outlook
# File → Get Add-ins → My add-ins → Add from File → manifest.xml
```

Open taskpane: Compose email → Click **SingleCase** button

## ⚠️ Common Issues

If the add-in doesn't work after setup:

### 🔴 On Send not firing
→ See [TROUBLESHOOTING.md#on-send-event-not-firing](./TROUBLESHOOTING.md#-on-send-event-not-firing)

### 🔴 Categories not applying
→ See [TROUBLESHOOTING.md#categories-not-working](./TROUBLESHOOTING.md#-categories-not-working)

### 🔴 BCC not detected
→ See [TROUBLESHOOTING.md#bcc-not-detected](./TROUBLESHOOTING.md#-bcc-not-detected)

## 🔧 Most Important Step: Clear Outlook Cache

**Windows:**
```bash
del /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*"
```

**macOS:**
```bash
rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/*
```

Then: **Close Outlook → Remove add-in → Restart → Re-sideload**

## 🧪 Verify Setup

After setup, test these features:

1. **Taskpane opens**: Compose email → SingleCase button → Panel appears
2. **BCC detection**: Add BCC `2023-0006@your-workspace.singlecase.ch` → Case auto-selected
3. **Categories**: File email → Check for "SC: Zařazeno" category
4. **On Send**: Select case + enable auto-file → Send → Toast notification appears

If any test fails, see [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)

## 📋 Requirements

- **Node.js** 18+
- **Outlook** Microsoft 365 or Outlook 2021+ (version 16.0.14326+)
- **Microsoft 365 account** (required for development)

Check your Outlook version: **Help → About Outlook**

## 🏗️ Project Structure

```
src/
├── commands/
│   ├── commands.ts           # Event runtime entry point
│   └── onMessageSendHandler.ts  # On Send handler
├── taskpane/
│   ├── index.tsx              # Taskpane entry point
│   └── components/
│       └── MainWorkspace/     # Main UI component
├── services/
│   ├── singlecase.ts          # SingleCase API
│   ├── graphMail.ts           # Microsoft Graph (categories)
│   └── singlecaseDocuments.ts # Document upload
└── utils/
    └── storage.ts             # Storage helpers
```

## 🔨 Development Commands

```bash
# Start dev server (keep running)
npm start

# Build for production
npm run build

# Lint code
npm run lint

# Validate manifest
npm run validate
```

## 🐛 Debugging

### Browser Console
1. Open taskpane
2. Press **F12** (Developer Tools)
3. Check **Console** tab for logs

### Network Requests
1. Open **Network** tab
2. All requests to `localhost:3000` should be `200 OK`
3. Check for certificate errors or 404s

### Outlook Desktop Logs (Windows)
1. Create: `%TEMP%\OutlookLogging.txt` with content:
   ```
   [LogSettings]
   Level=verbose
   ```
2. Restart Outlook
3. Check logs: `%TEMP%\Outlook Logging\`

## 🔐 Security & Certificates

The add-in uses **HTTPS** for local development. You must trust the development certificate:

```bash
npx office-addin-dev-certs install
npx office-addin-dev-certs verify
```

Expected: ✅ `The developer certificates are installed and trusted.`

If verification fails:
- **Windows**: Run PowerShell as Administrator
- **macOS**: Enter your password when prompted

## 🌐 Microsoft Graph Permissions

The add-in requires Graph API permissions for:
- Reading/writing email categories
- Accessing mailbox properties

These are configured in the Azure AD app registration. Users must consent on first use.

## 📦 Key Dependencies

- **office-js**: Official Office Add-ins API
- **React**: UI framework
- **TypeScript**: Type safety
- **Webpack**: Module bundler + dev server

## 🔄 Update Process

When pulling new changes:

```bash
# 1. Update dependencies
npm install

# 2. Clear Outlook cache (CRITICAL)
# See commands above

# 3. Restart dev server
npm start

# 4. Re-sideload manifest
# Remove old → Restart Outlook → Sideload new
```

## 🤝 Working with Other Developers

If the add-in works on one machine but not another:

1. **Compare Outlook versions** (Help → About)
2. **Compare Node.js versions** (`node --version`)
3. **Verify certificate trust** (`npx office-addin-dev-certs verify`)
4. **Clear cache on both machines**
5. **Use same Microsoft 365 account**

See [TROUBLESHOOTING.md](./TROUBLESHOOTING.md) for detailed comparison checklist.

## 📖 Architecture Notes

### Event-Based Activation
The add-in uses **event-based activation** for On Send:
- Defined in `manifest.xml` under `<Runtimes>` and `<LaunchEvent>`
- Handler in `src/commands/onMessageSendHandler.ts`
- Runs in **separate runtime** (not the taskpane)

### Storage Strategy
- **Compose mode**: Uses `OfficeRuntime.storage` + `roamingSettings`
- **Read mode**: Uses `OfficeRuntime.storage` + `localStorage`
- **Fallback keys**: For emails without stable ID before send

### Submail Detection
- Reads **BCC, To, Cc** recipients every 350ms
- Extracts case keys from SingleCase submail addresses
- Normalizes dashes/periods for matching (e.g., `2023-0005-001` ↔ `2023-0005.001`)
- **Highest priority** - overrides all other case suggestions

## 🎯 Testing Checklist

Before committing changes, test:

- [ ] Taskpane opens in compose mode
- [ ] Taskpane opens in read mode
- [ ] BCC detection works (submail → case auto-selected)
- [ ] Categories apply correctly
- [ ] On Send files email automatically
- [ ] Filing works for both new emails and replies
- [ ] Deleted documents show correct message
- [ ] Works in both Outlook Desktop and OWA

## 🚨 Known Issues

### macOS Outlook Desktop
- On Send may not fire on older macOS Outlook versions
- Categories API requires newer Outlook (16.54+)
- Use OWA for testing if Desktop fails

### Windows Outlook Desktop
- Cache is more aggressive - clear frequently
- Runtime logs available in `%TEMP%\Outlook Logging\`
- Antivirus may interfere with certificate trust

### Outlook Web (OWA)
- BCC API availability varies by browser
- Always use Chrome/Edge for best compatibility

## 📞 Getting Help

1. **Check logs**: Browser console (F12) shows most errors
2. **Read troubleshooting**: [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)
3. **Compare environments**: Check versions, certificates, cache
4. **Collect debug info**:
   - Console logs
   - Outlook version
   - Node.js version
   - Certificate verification output

---

**Ready to code?** Start with [QUICK_START.md](./QUICK_START.md) →
