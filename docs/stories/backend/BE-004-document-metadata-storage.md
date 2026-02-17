# BE-004: Document Metadata Storage

**Story ID:** BE-004
**Story Points:** 5
**Epic Link:** Core Filing System
**Priority:** High
**Status:** Ready for Development

---

## Description

Implement comprehensive metadata storage and retrieval for email documents. Metadata powers the suggestion engine, enables idempotent filing, supports advanced search, and provides context for document management operations.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

**Key Requirements:**
- Store email-specific metadata (from, to, subject, body preview)
- Support flexible schema via JSONB for future extensibility
- Enable fast queries for suggestion engine
- Maintain metadata consistency with document records

---

## Acceptance Criteria

- [ ] Metadata automatically stored when documents are filed via BE-002 API
- [ ] `document_metadata` table populated with email-specific fields
- [ ] JSONB metadata field supports custom properties
- [ ] GET endpoint retrieves document with full metadata
- [ ] UPDATE endpoint allows metadata updates
- [ ] All email metadata fields searchable via indexes
- [ ] Response time < 50ms for metadata queries
- [ ] Metadata validation enforces data quality
- [ ] Metadata migrations handle schema evolution
- [ ] Integration tests verify metadata consistency
- [ ] API documentation with metadata schema

---

## Technical Requirements

### Metadata Schema

#### Core Email Metadata Fields
```typescript
interface EmailMetadata {
  // Identifiers (critical for idempotency)
  conversationId: string;          // Office.js conversationId
  internetMessageId?: string;      // RFC822 Message-ID header

  // Email properties
  subject: string;                 // Original subject
  subjectNormalized: string;       // Normalized for matching
  fromEmail: string;               // Sender email
  fromName?: string;               // Sender display name
  toEmails: string[];              // Recipients
  ccEmails?: string[];             // CC recipients
  bccEmails?: string[];            // BCC recipients (if available)
  dateSent: string;                // ISO 8601 timestamp

  // Content metadata
  bodyPreview?: string;            // First 500 chars of body
  bodyFormat?: 'text' | 'html';    // Body format
  hasAttachments: boolean;         // Attachment flag
  attachmentCount?: number;        // Number of attachments
  importance?: 'low' | 'normal' | 'high';

  // Filing context
  filedFromMailbox?: string;       // User's mailbox email
  filedFromFolder?: string;        // Source folder (Inbox, Sent, etc.)
  filingMethod?: 'manual' | 'suggested' | 'auto';

  // Custom fields (flexible JSONB)
  custom?: Record<string, any>;
}
```

### API Specification

#### 1. Get Document with Metadata
```
GET /api/v1/documents/{documentId}
```

**Response:**
```json
{
  "id": "doc-uuid-123",
  "name": "Project Update.eml",
  "mime_type": "message/rfc822",
  "size_bytes": 45678,
  "case_id": "case-456",
  "directory_id": "dir-uuid-789",
  "created_at": "2024-02-15T10:35:00Z",
  "created_by_user_id": "user-abc",
  "modified_at": "2024-02-15T10:35:00Z",
  "latest_version": {
    "id": "version-uuid-1",
    "version_number": 1,
    "size_bytes": 45678,
    "created_at": "2024-02-15T10:35:00Z"
  },
  "metadata": {
    "conversationId": "AAQkADU1YjJj...",
    "internetMessageId": "<CADxFGQ1234@mail.gmail.com>",
    "subject": "Re: Project Update",
    "subjectNormalized": "project update",
    "fromEmail": "john@example.com",
    "fromName": "John Doe",
    "toEmails": ["jane@example.com"],
    "ccEmails": [],
    "dateSent": "2024-02-15T09:00:00Z",
    "bodyPreview": "Here is the latest project update...",
    "hasAttachments": true,
    "attachmentCount": 2,
    "importance": "normal",
    "filedFromMailbox": "user@company.com",
    "filingMethod": "manual"
  }
}
```

#### 2. Update Document Metadata
```
PATCH /api/v1/documents/{documentId}/metadata
```

**Request Body:**
```json
{
  "metadata": {
    "custom": {
      "department": "Legal",
      "priority": "high",
      "tags": ["contract", "review"]
    }
  }
}
```

**Use Cases:**
- Add custom tags/labels
- Update categorization
- Add notes or annotations

#### 3. Search Documents by Metadata
```
POST /api/v1/documents/search
```

**Request Body:**
```json
{
  "case_id": "case-456",
  "filters": {
    "from_email": "john@example.com",
    "date_sent_after": "2024-01-01T00:00:00Z",
    "has_attachments": true,
    "subject_contains": "contract"
  },
  "sort": {
    "field": "date_sent",
    "order": "desc"
  },
  "limit": 50,
  "offset": 0
}
```

**Response:**
```json
{
  "total": 123,
  "documents": [
    {
      "id": "doc-uuid-1",
      "name": "Contract Review.eml",
      "metadata": {
        "subject": "Contract Review",
        "fromEmail": "john@example.com",
        "dateSent": "2024-02-10T14:30:00Z"
      }
    }
  ],
  "limit": 50,
  "offset": 0
}
```

### Database Operations

#### Insert Metadata (during filing)
```sql
INSERT INTO document_metadata (
  id,
  document_id,
  conversation_id,
  internet_message_id,
  subject,
  subject_normalized,
  from_email,
  from_name,
  to_emails,
  cc_emails,
  date_sent,
  body_preview,
  has_attachments
) VALUES (
  $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13
);
```

#### Query by conversationId + subject
```sql
SELECT d.*, dm.*
FROM documents d
JOIN document_metadata dm ON dm.document_id = d.id
WHERE dm.conversation_id = $1
  AND dm.subject_normalized = $2
LIMIT 1;
```

#### Query by sender email (for suggestions)
```sql
SELECT d.case_id, COUNT(*) as filing_count
FROM documents d
JOIN document_metadata dm ON dm.document_id = d.id
WHERE dm.from_email = $1
  AND d.created_at > NOW() - INTERVAL '90 days'
GROUP BY d.case_id
ORDER BY filing_count DESC
LIMIT 5;
```

#### Full-text search on subject and body
```sql
SELECT d.*, dm.*,
       ts_rank(to_tsvector('english', dm.subject || ' ' || COALESCE(dm.body_preview, '')),
               plainto_tsquery('english', $1)) as rank
FROM documents d
JOIN document_metadata dm ON dm.document_id = d.id
WHERE to_tsvector('english', dm.subject || ' ' || COALESCE(dm.body_preview, ''))
      @@ plainto_tsquery('english', $1)
ORDER BY rank DESC
LIMIT 20;
```

### Metadata Validation

```javascript
function validateEmailMetadata(metadata) {
  const errors = [];

  // Required fields
  if (!metadata.conversationId && !metadata.internetMessageId) {
    errors.push('At least one identifier (conversationId or internetMessageId) required');
  }

  if (!metadata.subject || metadata.subject.trim() === '') {
    errors.push('Subject is required');
  }

  // Email validation
  if (metadata.fromEmail && !isValidEmail(metadata.fromEmail)) {
    errors.push('Invalid fromEmail format');
  }

  if (metadata.toEmails) {
    for (const email of metadata.toEmails) {
      if (!isValidEmail(email)) {
        errors.push(`Invalid toEmail: ${email}`);
      }
    }
  }

  // Date validation
  if (metadata.dateSent && !isValidISO8601(metadata.dateSent)) {
    errors.push('dateSent must be ISO 8601 format');
  }

  // Length limits
  if (metadata.subject.length > 998) {
    errors.push('Subject exceeds max length (998 chars)');
  }

  if (metadata.bodyPreview && metadata.bodyPreview.length > 1000) {
    errors.push('bodyPreview exceeds max length (1000 chars)');
  }

  return errors;
}
```

### Metadata Extraction

```javascript
async function extractMetadataFromEmail(emlContent) {
  // Parse .eml file (RFC822 format)
  const parsed = await parseEmail(emlContent);

  return {
    conversationId: parsed.headers['thread-index'] || null,
    internetMessageId: parsed.headers['message-id'] || null,
    subject: parsed.subject || '(no subject)',
    subjectNormalized: normalizeSubject(parsed.subject),
    fromEmail: extractEmail(parsed.from),
    fromName: extractName(parsed.from),
    toEmails: parsed.to.map(extractEmail),
    ccEmails: parsed.cc?.map(extractEmail) || [],
    dateSent: parsed.date.toISOString(),
    bodyPreview: extractBodyPreview(parsed.text || parsed.html, 500),
    bodyFormat: parsed.html ? 'html' : 'text',
    hasAttachments: parsed.attachments.length > 0,
    attachmentCount: parsed.attachments.length,
    importance: parsed.headers['importance'] || 'normal'
  };
}
```

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` shows the expected metadata structure:

### Metadata Structure
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 212-217:** Metadata fields in upload payload
  ```typescript
  metadata?: {
    subject?: string;
    fromEmail?: string;
    fromName?: string;
    [key: string]: any;
  }
  ```

- **Lines 249-260:** Full upload payload with metadata
  - Shows metadata is nested within document object
  - Flexible structure allows custom fields

### Metadata Querying
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 705-714:** Multiple metadata field locations
  ```javascript
  let docSubject =
    doc.metadata?.subject ||    // Metadata field (if we stored it)
    doc.subject ||              // Direct field
    doc.properties?.subject ||  // Properties object
    "";
  ```
  - Shows flexibility needed for different data sources

- **Lines 835-846:** conversationId and subject matching
  - Critical for idempotency checks
  - Demonstrates normalized comparison

### Filed Cache Structure
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/filedCache.ts`:

- **Lines 5-14:** `FiledEmailCache` type
  - Shows what metadata clients need for "Already Filed" detection
  - Includes case context (caseId, caseName, caseKey)

### Suggestion Engine Data
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/suggestionStorage.ts`:

- **Lines 3-4:** Sender/domain statistics structure
  - Shows how metadata feeds suggestion algorithm
  - Tracks count and lastSeenAt per case

---

## Dependencies

- **Requires:** BE-001 (Database Schema) - document_metadata table must exist
- **Integrates with:** BE-002 (Filing API) - metadata stored during filing
- **Powers:** BE-007 (Suggestion Engine) - provides data for suggestions
- **Enables:** BE-003 (Status Query) - metadata enables fast lookups

---

## Notes

### Why Separate Metadata Table?

**Benefits of `document_metadata` as separate table:**
1. **Performance:** Email-specific indexes don't bloat main documents table
2. **Flexibility:** Non-email documents (PDFs, etc.) don't need email fields
3. **Querying:** Optimize indexes for email-specific queries
4. **Schema Evolution:** Add email fields without affecting other document types

**Alternative:** Store in JSONB column in documents table
- Simpler schema
- Harder to index efficiently
- Mixing email and non-email documents

**Recommendation:** Use separate table for email metadata, JSONB for custom fields.

### Metadata Extraction Challenges

1. **Email Parsing:** Use robust library (e.g., mailparser for Node.js)
2. **Character Encoding:** Handle international characters correctly (UTF-8)
3. **HTML vs Text:** Prefer plain text for body preview, fallback to HTML
4. **Malformed Headers:** Handle emails with missing or invalid headers
5. **Large Emails:** Limit body preview to 500 chars to avoid storage bloat

### Performance Optimization

1. **Indexes:**
   - Composite: (conversation_id, subject_normalized)
   - Single: (internet_message_id)
   - Single: (from_email)
   - GIN: full-text search on subject + body_preview

2. **Query Patterns:**
   - Most common: conversationId + subject lookup (< 50ms)
   - Sender history: from_email + date range (< 100ms)
   - Full-text search: subject/body search (< 500ms)

3. **Caching:**
   - Cache recent metadata queries (5 min TTL)
   - Cache sender statistics (15 min TTL)
   - Invalidate on document creation/update

### Testing Strategy

```javascript
// Test Cases to Implement

1. Metadata Storage:
   - File email with full metadata → all fields stored
   - File email with minimal metadata → defaults applied
   - File email with custom metadata → stored in JSONB
   - Update metadata → changes persisted

2. Metadata Queries:
   - Get document → returns full metadata
   - Search by sender → correct documents returned
   - Search by date range → correct documents returned
   - Full-text search → ranked results

3. Validation:
   - Invalid email format → 400 Bad Request
   - Missing required fields → 400 Bad Request
   - Subject too long → truncated or rejected
   - Invalid date format → 400 Bad Request

4. Performance:
   - Query 100k documents by conversationId → < 50ms
   - Search by sender in 100k documents → < 100ms
   - Full-text search → < 500ms

5. Edge Cases:
   - Null conversationId → handled gracefully
   - Empty subject → stored as "(no subject)"
   - Special characters in subject → sanitized
   - Very long body preview → truncated
```

### Security & Privacy

- **PII Handling:** Email addresses are personal data (GDPR)
- **Body Preview:** Limit to 500 chars to avoid storing sensitive content
- **Attachments:** Don't store attachment content in metadata
- **Access Control:** Only return metadata for documents user can access
- **Redaction:** Support redacting sensitive metadata fields

### Monitoring & Observability

Track metrics:
- Metadata completeness (% of documents with all fields)
- Query performance by type
- Full-text search usage
- Metadata update frequency

Alert on:
- Query performance degradation
- High % of documents with missing metadata
- Full-text search index corruption

### Future Enhancements

- **Automatic Categorization:** ML-based email categorization
- **Named Entity Recognition:** Extract people, companies, dates from body
- **Sentiment Analysis:** Classify email sentiment
- **Thread Reconstruction:** Build email thread trees
- **Email Forwarding Detection:** Track forwarding chains
- **Attachment Metadata:** Extract metadata from attachments (PDF title, etc.)
