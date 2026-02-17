# BE-002: Idempotent Filing API

**Story ID:** BE-002
**Story Points:** 8
**Epic Link:** Core Filing System
**Priority:** Critical
**Status:** Ready for Development

---

## Description

Implement a robust, idempotent API endpoint for filing emails to cases. The API must prevent duplicate filings when the same email is filed multiple times (from different mailboxes, re-sent emails, or network retries). Support both new document creation and versioning of existing documents.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

**Key Challenge:** Detect when an email with the same conversationId and subject has already been filed, and either return the existing document or create a new version, depending on configuration.

---

## Acceptance Criteria

- [ ] `POST /api/v1/cases/{caseId}/documents` endpoint implemented
- [ ] Idempotency guaranteed: filing same email twice returns existing document
- [ ] Supports both single document and batch uploads (multiple attachments)
- [ ] Validates request payload (required fields, file size limits, MIME types)
- [ ] Handles base64-encoded document data (email .eml and attachments)
- [ ] Returns detailed response with document IDs and version information
- [ ] HTTP status codes:
  - 200 OK - Document(s) already exist (idempotent return)
  - 201 Created - New document(s) created
  - 400 Bad Request - Invalid payload
  - 401 Unauthorized - Missing/invalid token
  - 404 Not Found - Case doesn't exist
  - 409 Conflict - Document locked (concurrent edit)
  - 413 Payload Too Large - File size exceeds limit
  - 500 Internal Server Error - Server error
- [ ] Response time < 2 seconds for 10MB email with 5 attachments
- [ ] Concurrent filing of same email handled gracefully (no duplicate documents)
- [ ] Integration tests cover all scenarios

---

## Technical Requirements

### API Specification

#### Endpoint
```
POST /api/v1/cases/{caseId}/documents
```

#### Request Headers
```
Authorization: Bearer <token>
Content-Type: application/json
Accept: application/json
```

#### Request Body
```json
{
  "documents": [
    {
      "name": "Meeting Notes.eml",
      "mime_type": "message/rfc822",
      "data_base64": "base64EncodedEmailContent...",
      "directory_id": "uuid-optional",
      "metadata": {
        "conversationId": "AAQkADU...",
        "internetMessageId": "<msg123@example.com>",
        "subject": "Project Status Update",
        "fromEmail": "john@example.com",
        "fromName": "John Doe",
        "toEmails": ["jane@example.com"],
        "ccEmails": [],
        "dateSent": "2024-02-15T10:30:00Z",
        "bodyPreview": "Here is the status update...",
        "hasAttachments": true
      }
    }
  ]
}
```

#### Response Body (201 Created)
```json
{
  "documents": [
    {
      "id": "doc-uuid-1",
      "name": "Meeting Notes.eml",
      "mime_type": "message/rfc822",
      "size_bytes": 45678,
      "case_id": "case-123",
      "directory_id": "dir-uuid",
      "created_at": "2024-02-15T10:35:00Z",
      "created_by_user_id": "user-456",
      "latest_version": {
        "id": "version-uuid-1",
        "version_number": 1,
        "name": "Meeting Notes.eml",
        "size_bytes": 45678,
        "created_at": "2024-02-15T10:35:00Z"
      },
      "metadata": {
        "conversationId": "AAQkADU...",
        "subject": "Project Status Update"
      }
    }
  ],
  "status": "created"
}
```

#### Response Body (200 OK - Already Filed)
```json
{
  "documents": [
    {
      "id": "doc-uuid-1",
      "name": "Meeting Notes.eml",
      "mime_type": "message/rfc822",
      "size_bytes": 45678,
      "case_id": "case-123",
      "created_at": "2024-02-14T09:20:00Z",
      "latest_version": {
        "id": "version-uuid-2",
        "version_number": 2,
        "created_at": "2024-02-15T10:35:00Z"
      }
    }
  ],
  "status": "already_filed",
  "message": "Email already filed to this case"
}
```

### Idempotency Logic

1. **Primary Check:** conversationId + normalized subject
   - Normalize subject: lowercase, strip "Re:", "Fw:", multiple spaces
   - Query: `document_metadata WHERE conversation_id = ? AND subject_normalized = ?`
   - If match found: return existing document (200 OK)

2. **Fallback Check:** internetMessageId (if conversationId missing)
   - Query: `document_metadata WHERE internet_message_id = ?`
   - If match found: return existing document (200 OK)

3. **Versioning Mode (Optional):**
   - If `?mode=version` query param provided, create new version instead of returning existing
   - Increment version_number for existing document
   - Use case: Re-filing same email with updated content

4. **Concurrency Control:**
   - Use database row-level locking during check-and-insert
   - PostgreSQL: `SELECT ... FOR UPDATE SKIP LOCKED`
   - If lock conflict detected, return 409 Conflict with retry-after header

### Validation Rules

- `case_id`: Must exist in database, user must have write access
- `name`: Required, max 255 chars, sanitize filename
- `mime_type`: Required, must be in allowed list (see below)
- `data_base64`: Required, valid base64, max 25MB decoded
- `directory_id`: Optional, must belong to same case if provided
- `metadata.conversationId`: Highly recommended for idempotency
- `metadata.subject`: Required for email documents
- `metadata.fromEmail`: Validated email format

### Allowed MIME Types
```
Email:
  - message/rfc822

Attachments:
  - application/pdf
  - application/msword
  - application/vnd.openxmlformats-officedocument.wordprocessingml.document
  - application/vnd.ms-excel
  - application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
  - application/vnd.ms-powerpoint
  - application/vnd.openxmlformats-officedocument.presentationml.presentation
  - image/jpeg
  - image/png
  - image/gif
  - text/plain
  - text/html
  - application/zip
  - application/x-zip-compressed
```

### Storage Flow

1. **Decode Base64** → binary data
2. **Calculate Checksum** → SHA-256 hash
3. **Check Duplicate Content** → if same checksum exists, don't re-upload blob
4. **Upload to Blob Storage** → store in object storage (S3, Azure Blob, etc.)
5. **Create Database Records:**
   - Insert into `documents` table
   - Insert into `document_versions` table
   - Insert into `document_metadata` table (if metadata provided)
6. **Return Response** with document IDs

### Error Handling

```javascript
// Validation Error Response
{
  "error": "validation_failed",
  "message": "Invalid request payload",
  "details": [
    {
      "field": "documents[0].name",
      "error": "Filename is required"
    },
    {
      "field": "documents[1].data_base64",
      "error": "File size exceeds 25MB limit"
    }
  ]
}

// Case Not Found Response
{
  "error": "case_not_found",
  "message": "Case with ID 'case-123' does not exist"
}

// Concurrent Modification Response
{
  "error": "document_locked",
  "message": "Document is currently being modified. Please retry in a moment.",
  "retry_after_seconds": 5
}
```

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` shows the expected behavior:

### Document Upload Logic
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 206-307:** `uploadDocumentToCase` function
  - Request payload structure (lines 249-260)
  - Metadata handling (lines 212-217)
  - Error handling (lines 293-302)
  - Response structure (lines 304-306)

- **Lines 117-141:** `getDocumentMeta` for checking existing documents
  - Shows how to query document by ID
  - Null handling for 404 cases

### Idempotency Check Pattern
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 766-873:** `checkFiledStatusByConversationAndSubject` function
  - conversationId + normalized subject matching (lines 782-786)
  - Iterating through documents to find match (lines 827-866)
  - Returns document info if match found (lines 858-864)

- **Lines 625-741:** `findDocumentBySubject` function (case-specific search)
  - Multiple API endpoint attempts (lines 635-639)
  - Email file filtering (lines 698-702)
  - Subject normalization and matching (lines 687-723)

### Subject Normalization
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 587-607:** `normalizeSubject` function
  - Trim and lowercase (line 590)
  - Collapse whitespace (line 593)
  - Strip Re:/Fw:/Fwd: prefixes (lines 596-603)
  - Handle nested prefixes (lines 598-603)

### Version Upload
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 143-204:** `uploadDocumentVersion` function
  - Multiple endpoint attempts for compatibility (lines 170-177)
  - Error handling for unsupported endpoints (lines 193-196)
  - Version-specific payload (lines 157-168)

### Error Handling Pattern
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 96-115:** `expectJson` helper
  - Document lock detection (lines 100-104) - returns 423 status
  - JSON validation (lines 109-112)
  - Error message formatting (line 106)

---

## Dependencies

- **Requires:** BE-001 (Database Schema) - must be completed first
- **Blocks:** BE-003 (Status Query), BE-004 (Metadata Storage), BE-009 (Attachment Filing)
- **Related:** BE-008 (Document Rename API)

### External Dependencies
- Blob storage service (AWS S3, Azure Blob Storage, Google Cloud Storage, or MinIO)
- Authentication/Authorization middleware
- Rate limiting middleware

---

## Notes

### Idempotency Strategy
The API uses **conversationId + normalized subject** as the natural idempotency key. This is more reliable than internetMessageId because:
- conversationId is available at send time in Outlook
- internetMessageId may not be assigned until after send
- conversationId groups all replies in a thread
- Subject normalization handles "Re: Re: Fw:" variations

### Performance Optimization
1. **Check Before Upload:** Query for existing document BEFORE decoding base64 and uploading to blob storage
2. **Parallel Processing:** For batch uploads, process documents in parallel (but with max concurrency limit of 5)
3. **Streaming:** For large files, use streaming upload to blob storage
4. **Connection Pooling:** Reuse database connections
5. **CDN Caching:** Cache document metadata for recently filed documents

### Security Considerations
- **File Scanning:** Integrate with antivirus API for malware scanning before storing
- **Content Validation:** Verify MIME type matches actual file content (magic bytes check)
- **Size Limits:** Enforce strict file size limits (25MB default, configurable)
- **Rate Limiting:** Max 100 requests/minute per user
- **Authentication:** Verify JWT token signature and expiration
- **Authorization:** Check user has write access to case

### Testing Checklist
```javascript
// Test Cases to Implement

1. Happy Path:
   - File single email → 201 Created
   - File email with 3 attachments → 201 Created
   - File same email again → 200 OK (idempotent)

2. Idempotency:
   - File email twice with same conversationId → returns existing doc
   - File email with "Re:" prefix → matches original subject
   - Concurrent filing of same email → only one document created

3. Validation:
   - Missing required fields → 400 Bad Request
   - Invalid base64 → 400 Bad Request
   - File too large → 413 Payload Too Large
   - Invalid MIME type → 400 Bad Request
   - Case doesn't exist → 404 Not Found

4. Versioning:
   - File email with ?mode=version → creates version 2
   - File email without mode param → returns existing doc

5. Error Handling:
   - Blob storage unavailable → 500 Internal Server Error
   - Database connection lost → 500 Internal Server Error
   - Invalid auth token → 401 Unauthorized

6. Performance:
   - 10MB email uploads in < 2 seconds
   - Batch of 5 attachments processes in < 3 seconds
   - Idempotency check completes in < 100ms
```

### Monitoring & Observability
- Log all filing attempts with:
  - User ID
  - Case ID
  - Document count
  - Total payload size
  - Processing time
  - Outcome (created/already_filed/error)
- Track metrics:
  - Filing success rate
  - Average file size
  - P95/P99 response times
  - Idempotency hit rate
  - Error rate by error type
- Alert on:
  - Filing success rate < 95%
  - P95 response time > 3 seconds
  - Error rate > 5%
  - Blob storage errors

### Future Enhancements
- Support for streaming large files (> 100MB)
- Webhook notifications when filing completes
- Async filing for very large batches
- OCR for scanned PDF attachments
- Automatic metadata extraction from email headers
- Content indexing for full-text search
