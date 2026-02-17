# Outlook Add-in V1 – Production Specification

**Last Updated:** February 17, 2026
**Status:** Ready for Implementation

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

## What Is This?

The **Outlook Add-in** lets users save emails and their attachments directly to SingleCase cases without leaving Outlook. It works in Outlook on Windows, Mac, and the web.

Think of it like a "Save to Case" button that appears right inside your email client.

---

## The Problem We're Solving

**Current Situation:**
- Our prototype works, but it saves information on your computer (like cookies in a browser)
- If you file an email on your laptop, your desktop computer doesn't know about it
- You might accidentally file the same email twice, creating duplicate documents
- If your browser cache clears, you lose all history of what you've filed

**The Issue:**
This causes confusion, duplicate work, and makes it hard to trust the system across multiple devices.

---

## The Solution

We're rebuilding the add-in so that **all filing information lives on our servers**, not on your computer.

**What This Means:**
- File an email on your laptop → It shows as "filed" on your desktop instantly
- Try to file the same email twice → The system stops you and says "You already filed this"
- Your filing history is permanent and backed up
- Everything works the same whether you're on Windows, Mac, or web

It's like moving from saving documents on your computer to saving them in the cloud - they're available everywhere.

---

## What Can Users Do?

### Core Features

✅ **Log in securely** – Enter your workspace name, authenticate, and you're in

✅ **See smart suggestions** – The add-in suggests which case to file to based on the email

✅ **Search for cases** – Can't find the right case? Search by name or number

✅ **File emails in one click** – Reading an email? Click "File to Case" and it's saved

✅ **File while composing** – Writing an email? Select a case and it files automatically when you hit Send

✅ **Attach files automatically** – All email attachments are saved to the case too

✅ **Never file twice** – The system remembers what you've filed and prevents duplicates

✅ **Color-coded emails** – Filed emails get a green "SC: Filed" tag in Outlook

✅ **Works everywhere** – Same experience on Windows, Mac, and Outlook web

---

## What's NOT Included (For Now)

These features are planned for future versions:

❌ Advanced analytics dashboards
❌ AI-powered case suggestions (we store the data, smart suggestions come in V2)
❌ Bulk filing multiple emails at once
❌ Email threading (grouping related emails)

---

## How It Works

### For Users Reading Email

1. Open an email in Outlook
2. The add-in checks: "Has this been filed already?"
   - **Yes** → Shows you which case it's in, with a link to view it
   - **No** → Shows suggested cases or a search box
3. Click a case, click "File"
4. Done! The email is saved to SingleCase and tagged with "SC: Filed"

### For Users Sending Email

1. Compose a new email
2. Use the add-in to pick a case
3. Toggle on "File on Send"
4. Click Send
5. Your email sends normally AND gets filed to the case automatically

**Important:** Even if filing fails, your email still sends. We never block your work.

---

## Key Improvements Over Current Version

### 1. Cross-Device Consistency
**Before:** Filed on your laptop? Your desktop doesn't know.
**After:** File anywhere, see it everywhere instantly.

### 2. No Duplicate Documents
**Before:** Click "File" twice by accident → Two copies in SingleCase.
**After:** System detects you already filed it and stops you.

### 3. Reliable History
**Before:** Clear your browser → Lose all filing history.
**After:** Everything is saved permanently on our servers.

### 4. Better Error Messages
**Before:** "Error 500" (what does that mean?).
**After:** "Filing failed due to network issue. Click to retry."

### 5. Audit Trail
**Before:** Can't see who filed what or when.
**After:** Every filing is recorded with who, when, and how.

---

## Technical Architecture (Simplified)

### The Old Way
```
Your Computer → Stores filing info locally → Uploads to SingleCase
```
**Problem:** Each device has its own info. They don't talk to each other.

### The New Way
```
Your Computer → Asks Server: "Is this filed?" → Server knows everything
                ↓
            Server stores all filing records
```
**Benefit:** One source of truth. Everyone sees the same information.

---

## What Needs to Be Built

### Backend Work (Server-Side)
We need to build systems that:
- Remember which emails have been filed
- Prevent duplicate documents
- Provide case suggestions
- Store filing history
- Handle secure login

**10 stories** covering data storage, APIs, and business logic

### Frontend Work (User Interface)
We need to update the add-in to:
- Ask the server for filing status instead of checking locally
- Show better error messages
- Display filed status across devices
- Integrate secure login

**10 stories** covering UI, error handling, and authentication

### DevOps Work (Deployment)
We need to:
- Set up automatic deployments
- Configure different environments (development, staging, production)
- Deploy through Microsoft 365 Admin Center

**5 stories** covering deployment pipelines and infrastructure

### Quality Assurance
We need to:
- Test on Windows, Mac, and web
- Test cross-device scenarios
- Verify no duplicates are created
- Get sign-off from users

**6 stories** covering testing and user acceptance

### Security
We need to:
- Implement secure login (OAuth)
- Ensure no sensitive data is logged
- Validate all user inputs
- Set up rate limiting

**6 stories** covering security hardening

**Total: 37 stories across 6 work streams**

---

## Project Timeline

### Phase 1: Foundation (Weeks 1-2)
Set up secure login and server infrastructure

### Phase 2: Core Filing (Weeks 3-4)
Build filing system with duplicate prevention

### Phase 3: Features (Weeks 5-6)
Add suggestions, search, and error handling

### Phase 4: Deployment (Week 7)
Automate deployment and lock down security

### Phase 5: Testing (Weeks 8-9)
Comprehensive testing on all platforms

### Phase 6: Launch (Week 10)
Deploy to production with monitoring

**Total Duration: ~10 weeks**

---

## Dependencies

### What We Need Before We Start

**Backend:**
- Server-side authentication system (OAuth)
- Database to store filing records
- API endpoints for filing and status checks

**Infrastructure:**
- Production server environment
- Hosting for the add-in files
- SSL certificates for secure connections

**Microsoft 365:**
- Admin access to deploy add-ins
- Permissions to use Microsoft Graph API

---

## Success Criteria

We'll know we're successful when:

✅ No duplicate documents are created
✅ Filing on one device is visible on another within 5 seconds
✅ Users can't lose filing history
✅ Email filing takes less than 2 seconds
✅ System handles 100+ users without slowing down
✅ No critical bugs in first 30 days of production
✅ Positive user feedback (we'll measure this)

---

## Work Breakdown

| Epic | Stories | Focus Area |
|------|---------|-----------|
| **Backend** | BE-001 to BE-010 | Server-side filing logic, APIs, duplicate prevention |
| **Frontend** | FE-001 to FE-010 | User interface, authentication, error handling |
| **DevOps** | DO-001 to DO-005 | Deployment, environments, automation |
| **QA** | QA-001 to QA-006 | Testing, user acceptance, performance |
| **Security** | SEC-001 to SEC-006 | OAuth, logging, input validation |

---

## Dependency Map

Here's how the work flows (must be done in this order):

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

### Story Point Summary

| Phase | Points | Duration Estimate |
|-------|--------|-------------------|
| Phase 1: Foundation | 29 pts | 2 weeks |
| Phase 2: Core Filing | 39 pts | 2 weeks |
| Phase 3: Features & Polish | 44 pts | 2 weeks |
| Phase 4: Deployment & Security | 29 pts | 1 week |
| Phase 5: Testing & Launch | 39 pts | 2 weeks |
| **TOTAL** | **180 pts** | **9-10 weeks** |

### Point Scale Guide
- **1-2 points:** Simple task, clear path, minimal risk (2-4 hours)
- **3 points:** Straightforward task with some complexity (1 day)
- **5 points:** Moderate complexity, requires design decisions (2-3 days)
- **8 points:** Complex task, significant integration work (3-5 days)
- **13 points:** Very complex, multiple unknowns (1+ week) - *None in this project*

**Notes:**
- Some stories can run in parallel within phases (e.g., BE-005 and BE-006)
- Point estimates assume experienced developers familiar with the stack
- Buffer time included in phase durations for code review, deployment, and unexpected issues

---

## Testing Plan

### Platforms We'll Test
- ✅ Outlook Web (Chrome, Edge, Safari)
- ✅ Outlook Desktop Windows
- ✅ Outlook Desktop Mac

### What We'll Test

**Authentication**
- Can users log in?
- Does the session stay active?
- Does token refresh work?

**Filing Emails**
- Does filing work in one click?
- Are attachments included?
- Is the right case selected?

**Duplicate Prevention**
- File the same email twice → Does it stop you?
- File on laptop, try again on desktop → Does it know?

**File on Send**
- Does filing happen when you send?
- If filing fails, does the email still send?
- Can you retry from Sent Items?

**Cross-Device**
- File on Windows → See it on Mac?
- File on web → See it on desktop?
- How fast does it sync?

**Error Handling**
- Network timeout → Does it show a retry button?
- Server error → Is the message user-friendly?
- Bad login → Does it ask you to log in again?

---

## Security & Privacy

**What We're Doing to Keep Data Safe:**

✅ **Secure Login** – Using industry-standard OAuth (same as "Sign in with Google")
✅ **Encrypted Connections** – All data sent over HTTPS
✅ **No Sensitive Logging** – Email content and personal info never go in logs
✅ **Access Control** – Users can only file to cases they have permission for
✅ **Rate Limiting** – Prevents abuse by limiting how many requests can be made
✅ **Input Validation** – All user input is checked for security issues

---

## Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Users file same email twice | Low | High | Server-side duplicate detection prevents this |
| Cross-device sync is slow | Medium | Medium | We'll load test and optimize before launch |
| OAuth is complex to set up | Medium | High | Start early, dedicate focused time |
| UAT reveals blocking issues | Medium | Medium | Built-in buffer week in timeline |

---

## Sign-off Checklist

Before we go to production, we need approval from:

- [ ] **Engineering Lead** – Code is production-ready
- [ ] **Product Manager** – Features are complete and match requirements
- [ ] **QA Lead** – All tests passed, no critical bugs
- [ ] **Security Team** – Security review completed
- [ ] **Users (UAT)** – Real users have tested and approved

---

## Questions?

**For more details:**
- Technical specifications → See engineering documentation
- API details → See API documentation
- Test cases → See QA test plan
- Security review → See security documentation

**Point of Contact:** [Your Name/Team]
