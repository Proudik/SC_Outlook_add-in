# SingleCase Outlook Add-in

A Microsoft Outlook add-in that lets you file emails and attachments directly into cases in [SingleCase](https://singlecase.ch) — from Outlook Desktop, Outlook Web, or Outlook on Mac.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Sideloading the Add-in](#sideloading-the-add-in)
- [Available Scripts](#available-scripts)
- [Project Structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [Building for Production](#building-for-production)

---

## Features

- **File emails to cases** — attach received or outgoing emails to SingleCase cases in one click
- **Smart case suggestions** — AI-powered suggestions based on email content, participants, and history
- **Attachment selection** — choose which attachments to include when filing
- **Auto-file on send** — optionally file outgoing emails automatically when you click Send
- **BCC detection** — add a SingleCase submail address as BCC to auto-select the case
- **Document management** — view, rename, and delete filed documents directly from Outlook
- **Duplicate prevention** — detects already-filed emails to avoid duplicates
- Works in **Outlook Desktop (Windows & Mac)**, **Outlook Web (OWA)**, and **New Outlook**

---

## Prerequisites

Before you start, make sure you have:

| Requirement | Version / Details |
|---|---|
| **Node.js** | v18 or higher — [download](https://nodejs.org) |
| **npm** | Comes with Node.js |
| **Microsoft Outlook** | Microsoft 365 or Outlook 2021+ (build 16.0.14326 or later) |
| **Microsoft 365 account** | Required for development and sideloading |
| **SingleCase instance** | You need access to a SingleCase workspace |

Check your versions:
```bash
node --version   # should print v18.x.x or higher
npm --version
```

Check your Outlook version: **Help → About Microsoft Outlook**

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/Proudik/SC_Outlook_add-in.git
cd SC_Outlook_add-in
```

### 2. Install dependencies

```bash
npm install
```

### 3. Install and trust development certificates

The add-in runs over HTTPS locally. You need to trust the development certificate once:

```bash
npx office-addin-dev-certs install
```

When prompted, enter your system password (macOS) or run as Administrator (Windows).

Verify it worked:
```bash
npx office-addin-dev-certs verify
```

Expected output: `The developer certificates are installed and trusted.`

> **Windows**: If verification fails, run PowerShell as Administrator and try again.
> **macOS**: If verification fails, open Keychain Access, find the `localhost` certificate, and set it to "Always Trust".

### 4. Start the development server

```bash
npm start
```

This starts the webpack dev server at `https://localhost:3000` and opens Outlook for debugging. Keep this terminal running while you use the add-in.

---

## Sideloading the Add-in

After the dev server is running, you need to sideload `manifest.xml` into Outlook so it knows about the add-in.

### Outlook Desktop (Windows)

1. Open Outlook
2. Go to **File → Manage Add-ins** (or **Home → Get Add-ins**)
3. Click **My add-ins** in the left sidebar
4. Scroll down to **Custom add-ins** → click **Add a custom add-in → Add from File...**
5. Select `manifest.xml` from the project folder
6. Click **Install** when warned about trusted add-ins
7. Close the dialog — the **SingleCase** button should now appear in the ribbon

### Outlook Desktop (macOS)

1. Open Outlook
2. Go to **Tools → Get Add-ins**
3. Click **My add-ins** in the left sidebar
4. Scroll down to **Custom add-ins** → click **+ Add a custom add-in → Add from File...**
5. Select `manifest.xml` from the project folder
6. Click **Install**

### Outlook Web (OWA)

1. Open [outlook.office.com](https://outlook.office.com) in Chrome or Edge
2. Open any email → click the **three dots (...)** menu at the top of the email
3. Select **Get Add-ins**
4. Click **My add-ins** in the left sidebar
5. Scroll down to **Custom add-ins** → click **+ Add a custom add-in → Add from File...**
6. Upload `manifest.xml`
7. Click **Install**

### Verify the add-in is working

- Open or compose an email
- Click the **SingleCase** button in the ribbon (or the `...` more actions menu)
- The SingleCase panel should open on the right side

---

## Available Scripts

| Command | Description |
|---|---|
| `npm start` | Start debugging (launches dev server + Outlook) |
| `npm stop` | Stop the debugging session |
| `npm run dev-server` | Start only the webpack dev server (port 3000) |
| `npm run build` | Build for production (output to `dist/`) |
| `npm run build:dev` | Build for development |
| `npm run watch` | Watch mode — rebuild on file changes |
| `npm run validate` | Validate `manifest.xml` against Office schema |
| `npm run lint` | Run ESLint |
| `npm run lint:fix` | Fix ESLint issues automatically |
| `npm run prettier` | Format code with Prettier |

---

## Project Structure

```
SC_Outlook_add-in/
├── src/
│   ├── commands/
│   │   ├── commands.ts              # Event runtime entry point
│   │   └── onMessageSendHandler.ts  # On Send handler (auto-file on send)
│   ├── taskpane/
│   │   ├── index.tsx                # Taskpane entry point
│   │   └── components/
│   │       ├── MainWorkspace/       # Main filing interface
│   │       ├── CaseSelector/        # Case search and selection
│   │       ├── PromptBubble/        # User prompts and messages
│   │       └── SettingsModal/       # User settings
│   ├── services/
│   │   ├── singlecase.ts            # SingleCase REST API client
│   │   ├── graphMail.ts             # Microsoft Graph (email categories)
│   │   └── singlecaseDocuments.ts   # Document upload/management
│   └── utils/
│       ├── storage.ts               # Storage abstraction (OfficeRuntime + localStorage)
│       ├── caseSuggestionEngine.ts  # AI-powered case suggestions
│       └── sentPillStore.ts         # Filed email state persistence
├── assets/                          # Icons and images
├── manifest.xml                     # Add-in manifest (Office schema)
├── webpack.config.js                # Webpack configuration
└── package.json
```

---

## Troubleshooting

### Add-in doesn't appear after sideloading

Clear the Outlook add-in cache, then remove and re-sideload the add-in:

**Windows:**
```bash
del /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*"
```

**macOS:**
```bash
rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/*
```

Then: **Quit Outlook → Remove the add-in → Restart Outlook → Sideload again**

---

### "This site can't be reached" / certificate errors

The dev server uses HTTPS with a self-signed certificate. If Outlook shows a certificate error:

1. Open `https://localhost:3000/taskpane.html` directly in your browser
2. Accept/trust the certificate warning
3. Go back to Outlook and reload the add-in

---

### On Send not firing

The On Send handler requires event-based activation support:

- Outlook must be version **16.0.14326** or later
- On **macOS Outlook Desktop** older than 16.54, On Send may not fire — use OWA instead
- Make sure you are **not** in cached Exchange mode (Windows) — try switching to online mode

---

### Panel shows blank or loading forever

1. Press **F12** in Outlook to open Developer Tools
2. Check the **Console** tab for errors
3. Check the **Network** tab — all requests to `localhost:3000` should return `200 OK`
4. Make sure `npm start` (or `npm run dev-server`) is still running

---

### Debugging tips

- **Browser console**: Press F12 in Outlook Desktop or OWA → Console tab
- **Validate manifest**: Run `npm run validate` to check for schema errors
- **Check certificates**: Run `npx office-addin-dev-certs verify`

---

## Building for Production

1. Update the version number in `manifest.xml` (line: `<Version>`)
2. Build the project:
   ```bash
   npm run build
   ```
3. The `dist/` folder contains all bundled assets
4. Deploy `dist/` to your HTTPS hosting (e.g., Azure Static Web Apps, Nginx, IIS)
5. Update all `localhost:3000` URLs in `manifest.xml` to your production domain
6. Distribute the updated `manifest.xml` to users (or submit to Microsoft AppSource)

---

## Authentication

The add-in uses **Microsoft Authentication Library (MSAL)** for secure login using Office credentials. Users authenticate via Microsoft 365 SSO — no separate password is needed.

You will need a **SingleCase workspace URL** configured in the add-in settings on first use.

---

## License

MIT

---

**Built with React, TypeScript, and Office.js**
