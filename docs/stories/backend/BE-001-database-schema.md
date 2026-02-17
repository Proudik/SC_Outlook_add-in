# BE-001: Database Schema & Data Model

**Story ID:** BE-001
**Story Points:** 5
**Epic Link:** Backend Infrastructure
**Priority:** Critical
**Status:** Ready for Development

---

## Description

Design and implement the complete database schema for the Outlook Add-in backend. This foundation will support document management, case associations, user tracking, suggestion intelligence, and audit logging. The schema must be optimized for fast queries, support idempotent operations, and scale to handle thousands of documents per case.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

---

## Acceptance Criteria

- [ ] Database schema implemented with all required tables and relationships
- [ ] Indexes created for all frequently queried fields
- [ ] Foreign key constraints properly defined
- [ ] Migration scripts created and tested
- [ ] Schema supports efficient querying for:
  - Filing status checks by conversationId + subject
  - Case document listings
  - Suggestion engine queries (sender/domain history)
  - Audit log retrieval
- [ ] Database can handle 10,000+ documents per case without performance degradation
- [ ] Schema documentation created with ER diagram
- [ ] Rollback scripts tested and verified

---

## Technical Requirements

### Core Tables

#### 1. `documents`
Primary table for filed documents (emails and attachments).

```sql
documents:
  - id (uuid, primary key)
  - case_id (string/uuid, indexed, foreign key to cases.id)
  - name (string) - filename (e.g., "Meeting Notes.eml")
  - mime_type (string) - e.g., "message/rfc822", "application/pdf"
  - size_bytes (bigint)
  - directory_id (uuid, nullable, foreign key to directories.id)
  - created_at (timestamp with timezone)
  - created_by_user_id (string, indexed)
  - modified_at (timestamp with timezone)
  - latest_version_id (uuid, foreign key to document_versions.id)
  - metadata (jsonb) - flexible storage for custom fields

Indexes:
  - (case_id, created_at DESC)
  - (case_id, mime_type)
  - (created_by_user_id, created_at DESC)
  - GIN index on metadata jsonb column
```

#### 2. `document_versions`
Version history for documents (supports email re-filing).

```sql
document_versions:
  - id (uuid, primary key)
  - document_id (uuid, indexed, foreign key to documents.id)
  - version_number (integer) - starts at 1
  - name (string)
  - size_bytes (bigint)
  - storage_key (string) - blob storage reference
  - checksum_sha256 (string)
  - created_at (timestamp with timezone)
  - created_by_user_id (string)

Indexes:
  - (document_id, version_number DESC)
  - (storage_key) UNIQUE
```

#### 3. `document_metadata`
Email-specific metadata for suggestion engine and search.

```sql
document_metadata:
  - id (uuid, primary key)
  - document_id (uuid, indexed, foreign key to documents.id)
  - conversation_id (string, indexed) - Office.js conversationId
  - internet_message_id (string, indexed) - RFC822 Message-ID
  - subject (text, indexed with trigram/full-text)
  - subject_normalized (string, indexed) - lowercase, Re:/Fw: stripped
  - from_email (string, indexed)
  - from_name (string)
  - to_emails (text[])
  - cc_emails (text[])
  - date_sent (timestamp with timezone)
  - body_preview (text) - first 500 chars
  - has_attachments (boolean)

Indexes:
  - (conversation_id, subject_normalized) - for idempotency checks
  - (internet_message_id)
  - (from_email, case_id)
  - (subject_normalized) - trigram index for fuzzy search
  - (date_sent DESC)
```

#### 4. `directories`
Folder structure within cases (e.g., "Outlook add-in" folder).

```sql
directories:
  - id (uuid, primary key)
  - case_id (string/uuid, indexed)
  - parent_id (uuid, nullable, foreign key to directories.id)
  - name (string)
  - created_at (timestamp with timezone)
  - created_by_user_id (string)

Indexes:
  - (case_id, parent_id, name) UNIQUE
  - (parent_id)
```

#### 5. `cases`
Case metadata (may already exist - extend if needed).

```sql
cases:
  - id (string/uuid, primary key)
  - name (string)
  - case_id_visible (string) - e.g., "2024-0123"
  - client_id (uuid, nullable)
  - root_directory_id (uuid, nullable)
  - created_at (timestamp with timezone)
  - status (enum: active, archived, closed)

Indexes:
  - (case_id_visible) UNIQUE
  - (client_id, status)
```

#### 6. `user_favorites`
User's favorited cases for quick access.

```sql
user_favorites:
  - id (uuid, primary key)
  - user_id (string, indexed)
  - case_id (string/uuid, foreign key to cases.id)
  - created_at (timestamp with timezone)
  - position (integer) - for custom ordering

Indexes:
  - (user_id, case_id) UNIQUE
  - (user_id, position)
```

#### 7. `suggestion_history`
Track filing patterns for suggestion engine.

```sql
suggestion_history:
  - id (uuid, primary key)
  - user_id (string, indexed)
  - case_id (string/uuid, indexed)
  - sender_email (string, indexed)
  - sender_domain (string, indexed)
  - conversation_id (string, nullable)
  - filed_at (timestamp with timezone)
  - source (enum: manual, suggested, auto)

Indexes:
  - (user_id, sender_email, filed_at DESC)
  - (user_id, sender_domain, filed_at DESC)
  - (user_id, case_id, filed_at DESC)
  - (filed_at) - for cleanup of old entries
```

#### 8. `audit_logs`
Comprehensive audit trail for compliance.

```sql
audit_logs:
  - id (uuid, primary key)
  - timestamp (timestamp with timezone, indexed)
  - user_id (string, indexed)
  - action (string) - e.g., "document.filed", "document.renamed"
  - resource_type (string) - e.g., "document", "case"
  - resource_id (string)
  - details (jsonb) - flexible storage for action-specific data
  - ip_address (inet, nullable)
  - user_agent (text, nullable)

Indexes:
  - (timestamp DESC)
  - (user_id, timestamp DESC)
  - (resource_type, resource_id, timestamp DESC)
  - (action, timestamp DESC)
```

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` shows the expected data structures:

### Document Upload Structure
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:
- Lines 5-15: `UploadDocumentResponse` type
- Lines 30-35: `DocumentMeta` type
- Lines 218-260: Document upload payload structure with metadata

### Metadata Storage
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:
- Lines 212-217: Metadata fields (subject, fromEmail, fromName)
- Lines 705-710: Multiple metadata field locations for flexibility
- Lines 835-845: conversationId-based matching

### Suggestion Data
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/suggestionStorage.ts`:
- Lines 9-15: `CaseSuggestState` structure
- Lines 3-4: Sender and domain statistics structure
- Lines 144-192: `recordSuccessfulAttach` shows what data to track

### Audit Context
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:
- Lines 221-226: Request logging context
- Lines 283-287: Response status tracking

---

## Dependencies

- **Blocks:** BE-002 (Filing API), BE-003 (Status Query), BE-004 (Metadata Storage)
- **Requires:** Database server (PostgreSQL 14+ recommended)
- **Prerequisites:**
  - Database instance provisioned
  - Migration tool selected (e.g., Flyway, Liquibase, or framework-native)
  - Backup and restore procedures defined

---

## Notes

### Database Choice
- **Recommended:** PostgreSQL 14+ for JSONB support, excellent indexing, and mature ecosystem
- **Alternative:** MySQL 8+ with JSON column type
- Must support:
  - JSONB/JSON columns for flexible metadata
  - Trigram indexes for fuzzy text search (PostgreSQL `pg_trgm` extension)
  - Full-text search capabilities
  - Partial indexes for performance optimization

### Performance Considerations
- **Partitioning:** Consider partitioning `audit_logs` by month for better query performance
- **Archival:** Plan for archiving closed cases and old audit logs
- **Connection pooling:** Use PgBouncer or similar for PostgreSQL
- **Read replicas:** Consider read replicas for analytics and reporting

### Schema Evolution
- Use semantic versioning for migrations (e.g., V1.0.0__initial_schema.sql)
- Include both up and down migrations
- Test migrations on copy of production data
- Plan for zero-downtime schema changes (e.g., adding nullable columns first)

### Security
- Enable row-level security (RLS) if supporting multi-tenant architecture
- Encrypt sensitive fields (e.g., email content previews) at rest
- Use parameterized queries exclusively to prevent SQL injection
- Implement database user roles with least privilege principle

### Monitoring & Alerting
- Set up slow query logging (queries > 100ms)
- Monitor index usage and identify missing indexes
- Track database size growth and plan for scaling
- Alert on failed transactions and connection pool exhaustion

### Testing Strategy
- Unit tests for migration scripts
- Integration tests for all queries in isolation
- Load tests with 10,000+ documents per case
- Test concurrent filing scenarios for idempotency verification
- Verify cascading deletes work as expected

### Sample Queries to Optimize
```sql
-- Check if email already filed (idempotency)
SELECT d.id, d.case_id, dm.subject
FROM documents d
JOIN document_metadata dm ON dm.document_id = d.id
WHERE dm.conversation_id = $1
  AND dm.subject_normalized = $2
LIMIT 1;

-- Get sender filing history for suggestions
SELECT case_id, COUNT(*) as filing_count, MAX(filed_at) as last_filed
FROM suggestion_history
WHERE user_id = $1 AND sender_email = $2
  AND filed_at > NOW() - INTERVAL '90 days'
GROUP BY case_id
ORDER BY filing_count DESC, last_filed DESC
LIMIT 10;

-- List case documents with pagination
SELECT d.id, d.name, d.mime_type, d.created_at,
       u.name as created_by_name
FROM documents d
LEFT JOIN users u ON u.id = d.created_by_user_id
WHERE d.case_id = $1
ORDER BY d.created_at DESC
LIMIT $2 OFFSET $3;
```

### Data Retention Policy
- Keep all documents indefinitely (compliance requirement)
- Archive suggestion_history > 1 year old
- Archive audit_logs > 2 years old to separate table
- Implement soft deletes with `deleted_at` timestamp for documents

### Backup Strategy
- Daily full backups retained for 30 days
- Transaction log backups every 15 minutes
- Test restore procedures monthly
- Geographic redundancy for disaster recovery
