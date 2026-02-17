# Outlook Add-in V1 – Production Specification

**Last Updated:** February 17, 2026
**Status:** Ready for Implementation
**Implementation Approach:** Greenfield (Start from scratch, use existing demo as reference)

---

## Jira Links

This specification maps to the following Jira epics:
- [SC-ADDIN-EPIC-001: Outlook Add-in V1](link-to-epic-001)
- [SC-ADDIN-EPIC-002: Server-Side Source of Truth](link-to-epic-002)
- [SC-ADDIN-EPIC-003: Frontend Production Hardening](link-to-epic-003)
- [SC-ADDIN-EPIC-004: DevOps & Deployment](link-to-epic-004)
- [SC-ADDIN-EPIC-005: QA & UAT](link-to-epic-005)
- [SC-ADDIN-EPIC-006: Security Hardening](link-to-epic-006)

---

## Implementation Approach

### Starting Fresh
Developers will **build this add-in from scratch** rather than modifying the existing prototype. The current demo implementation serves as:
- ✅ **Reference for functionality** – See how features work
- ✅ **UI/UX guide** – Understand the intended user experience
- ✅ **Technical proof-of-concept** – Validate that Office.js APIs work as expected

### Why Start Fresh?
The existing prototype was built to validate the concept and iterate quickly. The production version requires:
- Clean architecture from day one
- Proper separation of concerns
- Production-grade error handling
- Scalable backend infrastructure
- Comprehensive testing from the start

### Reference Implementation
The demo is available at: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/`

Key reference points:
- **Frontend:** `src/taskpane/components/` – UI components and user flows
- **Services:** `src/services/` – API client patterns and Office.js integration
- **Manifest:** `manifest.xml` – Office Add-in configuration
- **Build:** `webpack.config.js` – Build and development setup

---

## What Is This?

The **Outlook Add-in** lets users save emails and their attachments directly to SingleCase cases without leaving Outlook. It works in Outlook on Windows, Mac, and the web.

Think of it like a "Save to Case" button that appears right inside your email client.

---

## The Problem We're Solving

**Current Situation:**
- Users must manually download emails and upload them to SingleCase
- Context switching between Outlook and SingleCase disrupts workflow
- Email attachments require multiple steps to file
- No automatic case suggestions based on email content

**The Issue:**
This creates friction, slows down case management, and increases the risk of emails not being properly documented.

---

## The Solution

A production-ready Outlook Add-in where **all filing information lives on our servers**, not on user devices.

**What This Means:**
- File an email on your laptop → It shows as "filed" on your desktop instantly
- Try to file the same email twice → The system stops you automatically
- Your filing history is permanent and backed up
- Everything works the same whether you're on Windows, Mac, or web

---

## What Users Can Do

### Core Features

✅ **Log in securely** – OAuth authentication with workspace selection

✅ **See smart suggestions** – The add-in suggests which case to file to based on email metadata and history

✅ **Search for cases** – Full-text search with client filtering

✅ **File emails in one click** – Reading an email? Click "File to Case" and it's saved

✅ **File while composing** – Writing an email? Select a case and it files automatically when you hit Send

✅ **Attach files automatically** – All email attachments are saved to the case

✅ **Never file twice** – Server-side duplicate detection prevents filing the same email twice

✅ **Color-coded emails** – Filed emails get a green "SC: Filed" category in Outlook

✅ **Works everywhere** – Same experience on Windows, Mac, and Outlook web

---

## Technical Requirements

### Technology Stack (Recommended)

**Frontend:**
- TypeScript + React
- Office.js SDK
- Build: Webpack or Vite
- Testing: Jest + React Testing Library

**Backend:**
- Language: PHP (Symfony) or Node.js
- Database: MySQL/PostgreSQL
- API: REST with JSON
- Authentication: OAuth 2.0 with PKCE

**Infrastructure:**
- Hosting: HTTPS required (Office Add-ins requirement)
- CDN for static assets
- SSL certificates
- CORS properly configured

---

## API Endpoints (To Be Built)

### Authentication
```
POST /oauth/authorize    - OAuth authorization endpoint
POST /oauth/token        - Token exchange
POST /oauth/refresh      - Refresh access token
POST /oauth/revoke       - Revoke token
```

### Email Filing
```
POST /api/v1/email-filings              - File email to case
GET  /api/v1/email-filings/status       - Check filing status
GET  /api/v1/email-filings/:id          - Get filing details
GET  /api/v1/email-filings              - List user's filings
```

### Cases
```
GET  /api/v1/cases                      - List cases (with search)
GET  /api/v1/cases/:id                  - Get case details
GET  /api/v1/cases/suggestions          - Get case suggestions
```

### Favourites
```
GET    /api/v1/users/me/favourite-cases      - List favourites
POST   /api/v1/users/me/favourite-cases/:id  - Add to favourites
DELETE /api/v1/users/me/favourite-cases/:id  - Remove from favourites
```

### Documents
```
POST /api/v1/cases/:id/documents        - Upload document to case
GET  /api/v1/documents/:id              - Get document metadata
```

---

## Data Models (To Be Built)

### EmailFiling
```typescript
{
  id: UUID
  internetMessageId: string (indexed, unique)
  conversationId: string (indexed)
  subject: string

  // Filing metadata
  caseId: UUID (foreign key)
  documentId: UUID (foreign key)

  // Audit trail
  userId: UUID
  submittedByEmail: string
  submittedVia: enum("read_mode", "compose_mode", "sent_items")

  // Idempotency
  idempotencyKey: string (unique)

  // Timestamps
  emailReceivedAt: datetime
  filedAt: datetime

  // Attachments
  attachmentCount: integer
  attachmentMetadata: JSONB[]
}
```

### Document (Extended)
```typescript
{
  // ... existing fields ...

  // Email-specific metadata
  submitted_by_user_id: UUID (nullable)
  submitted_by_email: string (nullable)
  submitted_via: enum("email_addon", "web_upload", "api")
  email_message_id: string (indexed, nullable)
  email_conversation_id: string (nullable)
  email_from: string (nullable)
  email_to: string[] (nullable)
  email_received_at: datetime (nullable)
}
```

---

## User Flows

### 1. Initial Setup
1. User installs add-in from Microsoft 365 Admin Center
2. Opens Outlook, sees SingleCase button
3. Clicks button → Add-in panel opens
4. Enters workspace URL (e.g., `company.singlecase.com`)
5. Redirected to OAuth login
6. Authenticates and returns to add-in
7. Ready to file emails

### 2. Filing a Received Email
1. User opens email in Outlook
2. Add-in queries: `GET /api/v1/email-filings/status?internetMessageId=...`
3. **If already filed:** Show "Filed to [Case Name]" with link
4. **If not filed:** Show case suggestions or search
5. User selects case
6. User clicks "File to Case"
7. Add-in calls: `POST /api/v1/email-filings` (with idempotency key)
8. Backend creates EmailFiling record and uploads document
9. Add-in applies "SC: Filed" category via Microsoft Graph
10. Success message shown

### 3. File on Send (Compose Mode)
1. User composes email
2. User selects case from suggestions or search
3. User enables "File on Send" toggle
4. User clicks Send in Outlook
5. Add-in intercepts send event (SoftBlock mode)
6. Add-in calls: `POST /api/v1/email-filings`
7. **If success:** Email sends normally
8. **If failure:** Email still sends (SoftBlock), user can retry from Sent Items

---

## Work Breakdown

| Epic | Stories | Focus Area |
|------|---------|-----------|
| **Backend** | BE-001 to BE-010 | Server-side filing logic, APIs, duplicate prevention |
| **Frontend** | FE-001 to FE-010 | User interface, authentication, error handling |
| **DevOps** | DO-001 to DO-005 | Deployment, environments, automation |
| **QA** | QA-001 to QA-006 | Testing, user acceptance, performance |
| **Security** | SEC-001 to SEC-006 | OAuth, logging, input validation |

**Total: 37 stories, 180 story points, ~9-10 weeks**

---

## Story Dependencies

```
┌─────────────────────────────────────────────────────────────────┐
│                      PHASE 1: FOUNDATION                        │
│                         (29 points)                             │
├─────────────────────────────────────────────────────────────────┤
│  • SEC-001: Token Management & Secure Storage         [5 pts]   │
│  • FE-001: OAuth Integration                          [8 pts]   │
│  • BE-001: Database Schema & Data Model               [5 pts]   │
│  • DO-001: Environment-Specific Manifests             [3 pts]   │
│  • DO-004: Environment Variable Management            [3 pts]   │
│  • BE-004: Document Metadata Storage                  [5 pts]   │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                    PHASE 2: CORE FILING                         │
│                         (39 points)                             │
├─────────────────────────────────────────────────────────────────┤
│  • BE-002: Idempotent Filing API                      [8 pts]   │
│  • BE-003: Filing Status Query API                    [5 pts]   │
│  • BE-009: Attachment Filing Logic                    [5 pts]   │
│  • BE-010: Audit Logging & Analytics                  [3 pts]   │
│  • FE-004: Integrate Idempotent Filing API            [8 pts]   │
│  • FE-002: Replace Local Duplicate Logic              [5 pts]   │
│  • SEC-006: Idempotency Key Security                  [3 pts]   │
│  • FE-010: Offline Mode & Sync Strategy               [2 pts]   │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                  PHASE 3: FEATURES & POLISH                     │
│                         (44 points)                             │
├─────────────────────────────────────────────────────────────────┤
│  • BE-007: Suggestion Intelligence Engine             [8 pts]   │
│  • BE-005: Case Search API                            [5 pts]   │
│  • BE-006: Favourites Management API                  [3 pts]   │
│  • BE-008: Document Rename API                        [3 pts]   │
│  • FE-003: Integrate Server-Side Suggestions          [5 pts]   │
│  • FE-005: Error Handling & User Feedback             [5 pts]   │
│  • FE-007: Favourites UI Integration                  [3 pts]   │
│  • FE-008: Progress Indicators & Loading States       [3 pts]   │
│  • FE-006: Categories UI Improvements                 [3 pts]   │
│  • SEC-005: Input Validation & Sanitization           [5 pts]   │
│  • SEC-004: Rate Limiting & DDoS Protection           [3 pts]   │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│               PHASE 4: DEPLOYMENT & SECURITY                    │
│                         (29 points)                             │
├─────────────────────────────────────────────────────────────────┤
│  • DO-002: CI/CD Pipeline Setup                       [8 pts]   │
│  • DO-003: Centralized Deployment Script              [5 pts]   │
│  • DO-005: Production Monitoring Setup                [5 pts]   │
│  • FE-009: CORS & Proxy Configuration                 [5 pts]   │
│  • SEC-002: PII Logging Audit & Remediation           [8 pts]   │
│  • SEC-003: CORS Policy Configuration                 [3 pts]   │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│                 PHASE 5: TESTING & LAUNCH                       │
│                         (39 points)                             │
├─────────────────────────────────────────────────────────────────┤
│  • QA-001: Backend Unit & Integration Tests           [8 pts]   │
│  • QA-002: Frontend Unit Tests                        [5 pts]   │
│  • QA-003: E2E Test Suite                             [8 pts]   │
│  • QA-004: UAT Test Plan & Scenarios                  [5 pts]   │
│  • QA-005: Performance Testing                        [5 pts]   │
│  • QA-006: UAT Execution & Sign-off                   [8 pts]   │
└─────────────────────────────────────────────────────────────────┘
```

---

## Success Criteria

We'll know we're successful when:

✅ No duplicate documents are created
✅ Filing on one device is visible on another within 5 seconds
✅ Users can't lose filing history
✅ Email filing takes less than 2 seconds
✅ System handles 100+ users without slowing down
✅ No critical bugs in first 30 days of production
✅ Positive user feedback (NPS > 50)

---

## Reference Documentation

For developers starting the implementation:

**Study these files in the demo:**
- `src/taskpane/components/MainWorkspace/MainWorkspace.tsx` – Main UI component
- `src/services/singlecase.ts` – API client pattern
- `src/services/graphMail.ts` – Microsoft Graph integration
- `manifest.xml` – Office Add-in manifest structure
- `webpack.config.js` – Development proxy configuration

**Office.js Documentation:**
- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Outlook Add-in API](https://learn.microsoft.com/en-us/javascript/api/outlook)
- [Event-based activation](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch)

---

## Story Files

Individual story specifications are located in:
- `docs/stories/backend/` – BE-001 through BE-010
- `docs/stories/frontend/` – FE-001 through FE-010
- `docs/stories/devops/` – DO-001 through DO-005
- `docs/stories/qa/` – QA-001 through QA-006
- `docs/stories/security/` – SEC-001 through SEC-006

---

## Questions?

**Point of Contact:** Martin Polasek (Product Lead)
