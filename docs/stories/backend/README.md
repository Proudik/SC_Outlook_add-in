# Backend Stories - Outlook Add-in

This directory contains comprehensive story files for building the backend API for the SingleCase Outlook Add-in. These stories are for **greenfield development** - you are building from scratch, not modifying existing code.

## Story Overview

| Story ID | Title | Points | Priority | Dependencies |
|----------|-------|--------|----------|--------------|
| [BE-001](./BE-001-database-schema.md) | Database Schema & Data Model | 5 | Critical | None |
| [BE-002](./BE-002-idempotent-filing-api.md) | Idempotent Filing API | 8 | Critical | BE-001 |
| [BE-003](./BE-003-filing-status-query-api.md) | Filing Status Query API | 5 | High | BE-001, BE-004 |
| [BE-004](./BE-004-document-metadata-storage.md) | Document Metadata Storage | 5 | High | BE-001 |
| [BE-005](./BE-005-case-search-api.md) | Case Search API | 5 | High | BE-001 |
| [BE-006](./BE-006-favourites-management-api.md) | Favourites Management API | 3 | Medium | BE-001 |
| [BE-007](./BE-007-suggestion-intelligence-engine.md) | Suggestion Intelligence Engine | 8 | High | BE-001, BE-004 |
| [BE-008](./BE-008-document-rename-api.md) | Document Rename API | 3 | Medium | BE-001 |
| [BE-009](./BE-009-attachment-filing-logic.md) | Attachment Filing Logic | 5 | High | BE-001, BE-002 |
| [BE-010](./BE-010-audit-logging-analytics.md) | Audit Logging & Analytics | 3 | Medium | BE-001 |

**Total Story Points:** 50

## Development Phases

### Phase 1: Foundation (13 points)
**Goal:** Set up core data structures and basic filing

1. **BE-001: Database Schema** (5 pts) - Create all tables, indexes, migrations
2. **BE-002: Idempotent Filing API** (8 pts) - Core filing endpoint with idempotency

### Phase 2: Core Features (18 points)
**Goal:** Enable complete filing workflow with status tracking

3. **BE-004: Document Metadata Storage** (5 pts) - Store email metadata
4. **BE-003: Filing Status Query API** (5 pts) - Check "already filed" status
5. **BE-009: Attachment Filing Logic** (5 pts) - File attachments with emails
6. **BE-005: Case Search API** (3 pts) - Search cases for filing

### Phase 3: Intelligence & UX (11 points)
**Goal:** Add smart suggestions and user experience features

7. **BE-007: Suggestion Intelligence Engine** (8 pts) - Smart case suggestions
8. **BE-006: Favourites Management API** (3 pts) - Favorite cases

### Phase 4: Polish & Operations (8 points)
**Goal:** Document management and observability

9. **BE-008: Document Rename API** (3 pts) - Rename filed documents
10. **BE-010: Audit Logging & Analytics** (5 pts) - Logging and analytics

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` is a **reference implementation** showing:

- Expected data structures and API responses
- Client-side logic that calls your backend
- Suggestion algorithm to replicate server-side
- Error handling patterns
- Metadata structures

**Important:** You are NOT modifying the demo. You are building a new backend API from scratch that the demo can eventually call.

## Key Technical Decisions

### Database
- **Recommended:** PostgreSQL 14+ for JSONB, full-text search, and trigram indexes
- **Alternative:** MySQL 8+ with JSON columns
- Must support: JSON columns, full-text search, complex indexes

### Architecture
- **RESTful API** with standard HTTP methods (GET, POST, PATCH, DELETE)
- **Authentication:** Bearer token (JWT recommended)
- **File Storage:** Separate blob storage (S3, Azure Blob, etc.)
- **Caching:** Redis or in-memory for performance

### Core Concepts

#### Idempotency
Filing the same email twice should not create duplicate documents. Use **conversationId + normalized subject** as natural key.

#### Metadata-Driven Suggestions
Store email metadata (sender, domain, subject) to power intelligent case suggestions that learn from user behavior.

#### Transaction Safety
File email and attachments together in a single transaction - either all succeed or all rollback.

#### Folder Organization
Automatically organize filed documents into "Outlook add-in" folder within each case.

## API Standards

### Response Status Codes
- `200 OK` - Successful GET/PATCH/DELETE
- `201 Created` - Resource created
- `204 No Content` - Successful DELETE with no response body
- `400 Bad Request` - Validation error
- `401 Unauthorized` - Missing/invalid authentication
- `403 Forbidden` - Insufficient permissions
- `404 Not Found` - Resource doesn't exist
- `409 Conflict` - Concurrent modification or duplicate
- `413 Payload Too Large` - File size exceeds limit
- `500 Internal Server Error` - Server error

### Error Response Format
```json
{
  "error": "error_code",
  "message": "Human-readable error message",
  "details": [
    {
      "field": "field_name",
      "error": "Field-specific error"
    }
  ]
}
```

### Pagination
```json
{
  "items": [...],
  "total": 1234,
  "limit": 50,
  "offset": 0
}
```

## Performance Targets

| Operation | Target |
|-----------|--------|
| File email (< 5MB) | < 2s |
| File email + 5 attachments (50MB) | < 3s |
| Check filing status | < 100ms (P95) |
| Case search | < 200ms (P95) |
| Case suggestions | < 200ms (P95) |
| Get document metadata | < 50ms (P95) |
| Rename document | < 100ms (P95) |

## Security Requirements

- **Authentication:** All endpoints require valid bearer token
- **Authorization:** Verify user has access to case/document
- **Input Validation:** Sanitize all user input
- **SQL Injection Prevention:** Use parameterized queries exclusively
- **File Scanning:** Integrate antivirus for uploaded files
- **Rate Limiting:** Prevent abuse (configurable per endpoint)
- **HTTPS Only:** Never allow HTTP in production

## Testing Requirements

Each story must include:

- **Unit Tests:** Business logic, validation, utilities
- **Integration Tests:** API endpoints with database
- **Load Tests:** Performance under load (1000 concurrent requests)
- **Edge Case Tests:** Error conditions, concurrent access
- **Security Tests:** Authorization, injection prevention

## Documentation Requirements

Each story must include:

- **API Documentation:** OpenAPI/Swagger spec
- **Code Comments:** Explain complex logic
- **README:** Setup instructions and architecture
- **Database Migrations:** Up and down scripts
- **Deployment Guide:** Environment setup and configuration

## Monitoring & Observability

### Metrics to Track
- API response times (P50, P95, P99)
- Error rates by endpoint
- Filing success rate
- Suggestion acceptance rate
- Database query performance
- Storage usage growth

### Logging
- All API requests (method, path, status, duration)
- All errors with stack traces
- Audit events (see BE-010)
- Performance warnings (slow queries)

### Alerts
- Error rate > 2%
- P95 response time > target
- Database connection pool exhaustion
- Storage approaching limits
- Security events (failed auth, unusual patterns)

## Getting Started

1. **Read BE-001** first to understand the data model
2. **Examine reference implementation** in demo project
3. **Set up development environment** (database, IDE, tools)
4. **Implement Phase 1** (foundation stories)
5. **Test thoroughly** with integration tests
6. **Review with team** before proceeding to next phase

## Questions & Clarifications

For questions about these stories, contact:
- Product Owner: [Name]
- Technical Lead: [Name]
- Architecture Review: [Name]

## Additional Resources

- Demo project: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/`
- API design guidelines: [Link]
- Database standards: [Link]
- Security requirements: [Link]
- Deployment pipeline: [Link]

---

**Note:** These stories are for backend development only. Frontend stories are in the `/docs/stories/frontend/` directory.
