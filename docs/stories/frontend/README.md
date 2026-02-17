# Frontend Stories Overview

This directory contains comprehensive story files for building the new React + TypeScript + Office.js Outlook add-in from scratch. Each story is detailed, actionable, and includes specific technical requirements, code examples, and acceptance criteria.

## Story Index

### Core Authentication & Filing (Critical Path)
1. **[FE-001: OAuth Integration](FE-001-oauth-integration.md)** - 8 points
   - Build OAuth 2.0 authentication flow from scratch
   - Dialog-based login with Office.js
   - Token management and refresh
   - Cross-runtime storage synchronization

2. **[FE-002: Replace Local Duplicate Logic](FE-002-replace-local-duplicate-logic.md)** - 5 points
   - Remove client-side duplicate detection
   - Integrate server-side duplicate check API
   - Display duplicate warnings with case info
   - Fail-open error handling

3. **[FE-003: Integrate Server-Side Suggestions](FE-003-integrate-server-side-suggestions.md)** - 5 points
   - Replace local suggestion algorithm with ML-powered API
   - Display top 3 suggestions with confidence scores
   - Auto-select high-confidence suggestions
   - Suggestion feedback loop

4. **[FE-004: Integrate Idempotent Filing API](FE-004-integrate-idempotent-filing-api.md)** - 8 points
   - Complete email filing integration
   - Attachment handling (Base64 upload)
   - Idempotent requests with internetMessageId
   - Progress indicators for multi-step filing
   - Category application after success

### User Experience & Polish
5. **[FE-005: Error Handling & User Feedback](FE-005-error-handling-user-feedback.md)** - 5 points
   - Centralized error handling system
   - Toast notifications (success, error, warning, info)
   - Error boundary for React errors
   - Offline detection banner
   - User-friendly error messages

6. **[FE-006: Categories UI Improvements](FE-006-categories-ui-improvements.md)** - 3 points
   - Category status badges (Filed/Unfiled/None)
   - Category quick actions menu
   - Office.js + Graph API fallback
   - Master category management
   - Cross-platform consistency

7. **[FE-007: Favourites UI Integration](FE-007-favourites-ui-integration.md)** - 3 points
   - Star/favorite cases for quick access
   - Favorites filter view
   - Backend API integration
   - Optimistic UI updates
   - Sync favorites across devices

8. **[FE-008: Progress Indicators & Loading States](FE-008-progress-indicators-loading-states.md)** - 3 points
   - Skeleton loaders for case lists
   - Progress bars for uploads
   - Loading buttons with inline spinners
   - Contextual loading messages
   - Smooth transitions

### Infrastructure & Advanced Features
9. **[FE-009: CORS & Proxy Configuration](FE-009-cors-proxy-configuration.md)** - 5 points
   - Dynamic workspace host routing
   - Webpack dev server proxy enhancement
   - Production proxy infrastructure (nginx or Lambda)
   - CORS header configuration
   - Security and rate limiting

10. **[FE-010: Offline Mode & Sync Strategy](FE-010-offline-mode-sync-strategy.md)** - 2 points
    - Offline detection and indicators
    - Cache management (cases, favorites, preferences)
    - Graceful degradation
    - Auto-refresh on reconnection
    - Cache-first API strategy

## Total Story Points: 47 points

## Development Sequence

### Phase 1: Foundation (16 points)
- FE-001 (OAuth) - 8 points
- FE-005 (Error Handling) - 5 points
- FE-009 (CORS/Proxy) - 5 points (partial - dev environment only)

**Rationale**: Authentication and error handling are required for all other stories. Proxy configuration enables API calls.

### Phase 2: Core Filing (18 points)
- FE-002 (Duplicate Detection) - 5 points
- FE-003 (Suggestions) - 5 points
- FE-004 (Filing API) - 8 points

**Rationale**: The main value proposition - filing emails to cases with smart suggestions and duplicate prevention.

### Phase 3: Polish & UX (11 points)
- FE-006 (Categories) - 3 points
- FE-007 (Favourites) - 3 points
- FE-008 (Loading States) - 3 points
- FE-010 (Offline Mode) - 2 points

**Rationale**: Improve user experience and productivity. Can be done in parallel.

### Phase 4: Production Deployment (2 points)
- FE-009 (CORS/Proxy - production) - 2 points (remaining)

**Rationale**: Deploy production proxy infrastructure and finalize CORS configuration.

## Key Technologies

- **Frontend Framework**: React 18 + TypeScript
- **Office Add-ins**: Office.js API
- **Microsoft Graph**: Graph API for fallback operations
- **UI Library**: Fluent UI (React components)
- **Build Tool**: Webpack 5
- **Auth**: OAuth 2.0 with MSAL
- **API**: REST APIs with fetch
- **State Management**: React hooks + Context API
- **Storage**: sessionStorage + OfficeRuntime.storage + localStorage

## Reference Implementation

The demo codebase at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` serves as a reference implementation. Key files to review:

- **Auth**: `src/services/auth.ts`, `src/dialog/DialogAuth.tsx`
- **Office.js**: `src/services/graphMail.ts`, `src/taskpane/taskpane.ts`
- **API**: `src/services/singlecase.ts`
- **Components**: `src/taskpane/components/`
- **Webpack**: `webpack.config.js`

**Important**: The demo uses public token authentication and local suggestion logic. These must be replaced with OAuth and server-side APIs as specified in the stories.

## Story Format

Each story follows this structure:

1. **Header**: Story ID, points, epic, status
2. **Description**: What needs to be built and why
3. **Acceptance Criteria**: Specific, testable requirements
4. **Technical Requirements**: Components, services, hooks, APIs
5. **API Integration Patterns**: Code examples and patterns
6. **Reference Implementation**: What to review in the demo
7. **Dependencies**: Prerequisites and blockers
8. **Notes**: Edge cases, testing, performance, accessibility
9. **Definition of Done**: Checklist of completion criteria

## Getting Started

1. Read FE-001 (OAuth Integration) first - it's the foundation
2. Review the demo codebase to understand existing patterns
3. Set up development environment (Node.js, npm, Office.js debugger)
4. Start with Phase 1 stories (authentication and error handling)
5. Test on Outlook Desktop, Web, and Mac throughout development

## Documentation Standards

- All code examples are TypeScript
- Office.js APIs are properly typed
- Error handling is comprehensive
- Accessibility (ARIA) is considered
- Performance implications are noted
- Testing strategies are provided

## Questions or Clarifications

If any story requirements are unclear:
1. Check the Reference Implementation section for patterns
2. Review related stories for context
3. Consult Office.js documentation for API specifics
4. Ask for clarification on technical approach

## Success Criteria

The new add-in will be considered complete when:
- All 10 stories meet their Definition of Done
- Add-in works on Windows, Mac, and Web Outlook
- Users can authenticate, file emails, and manage cases
- Error handling is comprehensive and user-friendly
- Performance is acceptable (< 2s for filing)
- Code is tested and documented

---

**Last Updated**: February 17, 2026
**Total Stories**: 10
**Total Story Points**: 47
