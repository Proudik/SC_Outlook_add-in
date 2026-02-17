# BE-003: Filing Status Query API

**Story ID:** BE-003
**Story Points:** 5
**Epic Link:** Core Filing System
**Priority:** High
**Status:** Ready for Development

---

## Description

Implement an efficient API endpoint to check if an email has already been filed, supporting cross-mailbox detection. This enables the add-in to show "Already Filed" status when a user opens an email that has been filed by themselves or another user in the organization.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

**Key Use Cases:**
1. User receives email, files it → later opens same email in Sent Items → should show "Already Filed"
2. User A files email → User B receives same email → should show "Already Filed to Case X"
3. Email thread: reply is filed → later replies in same thread can reference original filing

---

## Acceptance Criteria

- [ ] `POST /api/v1/documents/check-filed-status` endpoint implemented
- [ ] Supports querying by conversationId + subject (primary method)
- [ ] Supports fallback querying by internetMessageId
- [ ] Returns document ID, case info, and filing timestamp
- [ ] Response time < 100ms for 95% of requests
- [ ] Handles missing/null identifiers gracefully
- [ ] HTTP status codes:
  - 200 OK - Request processed (document may or may not be filed)
  - 400 Bad Request - Invalid request payload
  - 401 Unauthorized - Missing/invalid token
  - 500 Internal Server Error - Server error
- [ ] Query optimization: indexed fields, no full table scans
- [ ] Integration tests cover all scenarios
- [ ] API documentation with examples

---

## Technical Requirements

### API Specification

#### Endpoint
```
POST /api/v1/documents/check-filed-status
```

Why POST instead of GET? Because conversationId can be very long (200+ chars), and we're passing multiple search parameters.

#### Request Headers
```
Authorization: Bearer <token>
Content-Type: application/json
Accept: application/json
```

#### Request Body
```json
{
  "conversationId": "AAQkADU1YjJjOTE4LTkyMzUtNGU0Ny04YzRlLTk3MTNkODc5YzBkZQAQAN...",
  "subject": "Re: Fw: Project Status Update",
  "internetMessageId": "<CADxFGQ1234@mail.gmail.com>",
  "workspaceId": "workspace-uuid-optional"
}
```

**Fields:**
- `conversationId` (string, optional): Office.js conversationId - primary identifier
- `subject` (string, optional): Email subject for matching with normalization
- `internetMessageId` (string, optional): RFC822 Message-ID header - fallback identifier
- `workspaceId` (string, optional): Limit search to specific workspace/tenant

**Note:** At least one identifier must be provided. Best practice: send all three when available.

#### Response Body - Document Found
```json
{
  "filed": true,
  "document": {
    "id": "doc-uuid-123",
    "name": "Project Status Update.eml",
    "case_id": "case-456",
    "case_name": "Acme Corp - Q4 Planning",
    "case_key": "2024-0123",
    "directory_id": "dir-uuid-789",
    "directory_path": "/Outlook add-in",
    "filed_at": "2024-02-15T10:35:00Z",
    "filed_by_user_id": "user-abc",
    "filed_by_user_name": "John Doe",
    "latest_version": {
      "id": "version-uuid-1",
      "version_number": 1
    }
  }
}
```

#### Response Body - Document Not Found
```json
{
  "filed": false,
  "message": "Email has not been filed"
}
```

#### Response Body - Multiple Matches (Edge Case)
```json
{
  "filed": true,
  "multiple_matches": true,
  "documents": [
    {
      "id": "doc-uuid-123",
      "case_id": "case-456",
      "case_name": "Acme Corp - Q4 Planning",
      "filed_at": "2024-02-15T10:35:00Z"
    },
    {
      "id": "doc-uuid-789",
      "case_id": "case-999",
      "case_name": "Acme Corp - Legal",
      "filed_at": "2024-02-16T14:20:00Z"
    }
  ],
  "message": "Email filed to multiple cases"
}
```

### Query Logic

#### Primary Search Strategy (conversationId + subject)
```sql
SELECT
  d.id,
  d.name,
  d.case_id,
  d.directory_id,
  d.created_at as filed_at,
  d.created_by_user_id as filed_by_user_id,
  dm.conversation_id,
  dm.subject,
  dm.subject_normalized,
  c.name as case_name,
  c.case_id_visible as case_key
FROM documents d
JOIN document_metadata dm ON dm.document_id = d.id
JOIN cases c ON c.id = d.case_id
WHERE dm.conversation_id = ?
  AND dm.subject_normalized = ?
ORDER BY d.created_at DESC
LIMIT 1;
```

**Subject Normalization (server-side):**
```javascript
function normalizeSubject(subject) {
  if (!subject) return '';

  let normalized = subject.trim().toLowerCase();

  // Collapse multiple spaces
  normalized = normalized.replace(/\s+/g, ' ');

  // Strip Re:/Fw:/Fwd: prefixes (handle nested)
  let prevLength;
  do {
    prevLength = normalized.length;
    normalized = normalized.replace(/^(re|fw|fwd):\s*/i, '');
  } while (normalized.length !== prevLength && normalized.length > 0);

  return normalized.trim();
}
```

#### Fallback Search Strategy (internetMessageId)
```sql
SELECT
  d.id,
  d.name,
  d.case_id,
  d.created_at as filed_at,
  dm.internet_message_id
FROM documents d
JOIN document_metadata dm ON dm.document_id = d.id
WHERE dm.internet_message_id = ?
ORDER BY d.created_at DESC
LIMIT 1;
```

#### Search Flow
1. **Primary:** If `conversationId` AND `subject` provided → search by both
2. **Fallback 1:** If no match, try `internetMessageId`
3. **Fallback 2:** If no match, try `conversationId` alone (less reliable)
4. **Return:** First match found, or `filed: false`

### Performance Requirements

- **Response Time:**
  - P50: < 50ms
  - P95: < 100ms
  - P99: < 200ms

- **Indexing Strategy:**
  - Composite index: `(conversation_id, subject_normalized)`
  - Index: `(internet_message_id)`
  - Index: `(conversation_id)` for fallback

- **Query Optimization:**
  - Use `LIMIT 1` to stop at first match
  - Avoid `SELECT *` - only fetch needed columns
  - Use query planner EXPLAIN to verify index usage
  - Consider read replica for high-traffic scenarios

### Caching Strategy

Implement in-memory caching with short TTL:

```javascript
// Cache key: hash(conversationId + subject_normalized)
// Cache value: { filed, document }
// TTL: 5 minutes

// Cache hit rate should be > 60% for typical usage
// Cache invalidation: on document creation/deletion
```

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` shows the expected behavior:

### Filing Status Check Logic
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 766-873:** `checkFiledStatusByConversationAndSubject` function
  - Input validation (lines 770-778)
  - conversationId + normalized subject matching (lines 782-786)
  - Document listing and filtering (lines 798-824)
  - Metadata comparison (lines 834-846)
  - Return structure (lines 858-864)

### Client-Side Filing Cache
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/filedCache.ts`:

- **Lines 101-149:** `getFiledEmailFromCache` function
  - Shows what data client needs (lines 132-139)
  - Cache structure (lines 5-14)
  - Platform detection logging (lines 109-115)

### Subject Normalization
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 587-607:** `normalizeSubject` function
  - Implementation details for server-side matching
  - Handles nested prefixes (lines 598-603)

### Document Search Pattern
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 625-741:** `findDocumentBySubject` function
  - Multiple endpoint attempts (lines 635-639)
  - Response structure handling (lines 666-671)
  - Email file filtering (lines 698-702)
  - Metadata field locations (lines 705-714)

---

## Dependencies

- **Requires:** BE-001 (Database Schema) - must be completed first
- **Requires:** BE-004 (Metadata Storage) - document_metadata table must exist
- **Related:** BE-002 (Filing API) - works together for idempotency

---

## Notes

### Why This API is Critical

This API enables the add-in to show "Already Filed" status instantly, preventing:
- Duplicate filing attempts
- User confusion ("Did I already file this?")
- Wasted time searching through cases

### Cross-Mailbox Detection

**Problem:** Office.js doesn't provide internetMessageId at send time (only available after send). conversationId is available immediately.

**Solution:** Use conversationId + subject as composite key. This works because:
- conversationId is unique per email thread
- Subject normalization handles "Re:" and "Fw:" variations
- Both fields are available at send time

**Edge Case:** If user changes subject mid-thread, it creates new conversationId. This is acceptable - treat as new email.

### Handling Missing Identifiers

| Scenario | Available Data | Strategy |
|----------|---------------|----------|
| New compose email | subject only (no conversationId yet) | Search by subject (less reliable) |
| Sent email opened later | conversationId + subject | Primary strategy |
| Received email | all three identifiers | Primary + fallback |
| Forward as new email | new conversationId, original subject | Treat as new email |

### Testing Strategy

```javascript
// Test Cases to Implement

1. Happy Path:
   - Email filed → check status → filed: true
   - Email not filed → check status → filed: false
   - Multiple cases → returns first match

2. Identifier Matching:
   - Match by conversationId + subject → found
   - Match by internetMessageId → found
   - conversationId match, subject mismatch → not found
   - Subject only (no conversationId) → found if unique

3. Subject Normalization:
   - "Project Update" → matches "Re: Project Update"
   - "Re: Fw: Project" → matches "Project"
   - "PROJECT UPDATE" → matches "project update"
   - "Project    Update" → matches "Project Update"

4. Edge Cases:
   - Null conversationId → fallback to internetMessageId
   - Empty subject → use conversationId only
   - Very long conversationId (200+ chars) → handled
   - Special characters in subject → sanitized

5. Performance:
   - 1000 concurrent requests → no degradation
   - Large database (100k+ documents) → < 100ms response
   - Cache hit → < 10ms response

6. Security:
   - Invalid token → 401 Unauthorized
   - User queries case without access → not found (don't leak existence)
```

### Security Considerations

- **Access Control:** Only return documents user has read access to
- **Data Leakage:** Don't reveal case names for cases user can't access
- **Rate Limiting:** Max 500 requests/minute per user (checking is frequent)
- **Query Injection:** Always use parameterized queries

### Monitoring & Observability

Track metrics:
- Query response time (P50, P95, P99)
- Cache hit rate
- Found vs not-found ratio
- Query type distribution (conversationId vs internetMessageId)
- Error rate

Alert on:
- P95 response time > 100ms
- Cache hit rate < 50%
- Error rate > 2%

### Future Enhancements

- **Batch API:** Check status for multiple emails in one request
- **Webhook:** Subscribe to filing events for real-time updates
- **Search by date range:** Limit search to recent filings for better performance
- **Fuzzy subject matching:** Use trigram similarity for typos
- **Full-text search:** Search email body content for more accurate matching
