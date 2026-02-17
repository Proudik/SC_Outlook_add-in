# BE-010: Audit Logging & Analytics

**Story ID:** BE-010
**Story Points:** 3
**Epic Link:** Observability & Compliance
**Priority:** Medium
**Status:** Ready for Development

---

## Description

Implement comprehensive audit logging for all user actions and provide analytics API to query usage patterns. Support compliance requirements, security monitoring, and usage analytics for product improvement.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

**Key Requirements:**
- Log all document operations (file, rename, delete)
- Log all case operations (search, favorite, access)
- Log authentication events
- Provide query API for audit logs
- Generate usage analytics and reports

---

## Acceptance Criteria

- [ ] All API operations automatically logged to `audit_logs` table
- [ ] Logs include: timestamp, user, action, resource, details, IP address
- [ ] `GET /api/v1/audit-logs` endpoint for querying logs
- [ ] `GET /api/v1/analytics/usage` endpoint for analytics
- [ ] Supports filtering by: user, action, date range, resource
- [ ] Response time < 500ms for log queries
- [ ] Log retention: 2 years minimum
- [ ] Old logs archived to separate storage
- [ ] Privacy: redact sensitive data (email content)
- [ ] HTTP status codes:
  - 200 OK - Logs/analytics returned
  - 400 Bad Request - Invalid query parameters
  - 401 Unauthorized - Missing/invalid token
  - 403 Forbidden - Insufficient permissions (admin only)
  - 500 Internal Server Error - Server error
- [ ] Integration tests verify logging
- [ ] API documentation with examples

---

## Technical Requirements

### Events to Log

#### Document Operations
```javascript
const DOCUMENT_EVENTS = {
  'document.filed': {
    description: 'Document filed to case',
    details: { case_id, document_id, mime_type, size_bytes }
  },
  'document.renamed': {
    description: 'Document renamed',
    details: { from, to, case_id }
  },
  'document.deleted': {
    description: 'Document deleted',
    details: { case_id, document_id, name }
  },
  'document.downloaded': {
    description: 'Document downloaded',
    details: { document_id, case_id }
  },
  'document.viewed': {
    description: 'Document metadata viewed',
    details: { document_id, case_id }
  }
};
```

#### Case Operations
```javascript
const CASE_EVENTS = {
  'case.searched': {
    description: 'Case search performed',
    details: { query, results_count }
  },
  'case.viewed': {
    description: 'Case details viewed',
    details: { case_id }
  },
  'case.favorited': {
    description: 'Case added to favorites',
    details: { case_id }
  },
  'case.unfavorited': {
    description: 'Case removed from favorites',
    details: { case_id }
  }
};
```

#### Suggestion Operations
```javascript
const SUGGESTION_EVENTS = {
  'suggestion.generated': {
    description: 'Case suggestions generated',
    details: { suggestions_count, top_case_id, confidence }
  },
  'suggestion.accepted': {
    description: 'User accepted suggested case',
    details: { case_id, confidence }
  },
  'suggestion.rejected': {
    description: 'User rejected suggestions',
    details: { suggested_case_ids }
  }
};
```

#### Authentication Events
```javascript
const AUTH_EVENTS = {
  'auth.login': {
    description: 'User logged in',
    details: { method }
  },
  'auth.logout': {
    description: 'User logged out',
    details: {}
  },
  'auth.token_refreshed': {
    description: 'Auth token refreshed',
    details: {}
  },
  'auth.failed': {
    description: 'Authentication failed',
    details: { reason, attempted_user_id }
  }
};
```

### API Specification

#### 1. Query Audit Logs
```
GET /api/v1/audit-logs
```

**Query Parameters:**
- `user_id` (string, optional): Filter by user
- `action` (string, optional): Filter by action type
- `resource_type` (string, optional): Filter by resource type (document, case, etc.)
- `resource_id` (string, optional): Filter by specific resource
- `from_date` (ISO 8601, optional): Start of date range
- `to_date` (ISO 8601, optional): End of date range
- `limit` (integer, optional): Max results (default: 100, max: 1000)
- `offset` (integer, optional): Pagination offset (default: 0)

**Example Request:**
```
GET /api/v1/audit-logs?user_id=user-123&action=document.filed&from_date=2024-02-01T00:00:00Z&limit=50
```

**Response:**
```json
{
  "logs": [
    {
      "id": "log-uuid-1",
      "timestamp": "2024-02-17T15:30:00Z",
      "user_id": "user-123",
      "user_name": "Jane Doe",
      "action": "document.filed",
      "resource_type": "document",
      "resource_id": "doc-uuid-456",
      "details": {
        "case_id": "case-789",
        "document_name": "Contract Review.eml",
        "mime_type": "message/rfc822",
        "size_bytes": 45678
      },
      "ip_address": "192.168.1.100",
      "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0"
    },
    {
      "id": "log-uuid-2",
      "timestamp": "2024-02-17T14:20:00Z",
      "user_id": "user-123",
      "action": "case.searched",
      "resource_type": "case",
      "details": {
        "query": "Acme Corp",
        "results_count": 12
      },
      "ip_address": "192.168.1.100"
    }
  ],
  "total": 2,
  "limit": 50,
  "offset": 0
}
```

#### 2. Usage Analytics
```
GET /api/v1/analytics/usage
```

**Query Parameters:**
- `period` (enum): "day", "week", "month", "year"
- `from_date` (ISO 8601, optional): Start date
- `to_date` (ISO 8601, optional): End date
- `user_id` (string, optional): Specific user (admin can query all users)

**Response:**
```json
{
  "period": "month",
  "from_date": "2024-02-01T00:00:00Z",
  "to_date": "2024-02-29T23:59:59Z",
  "total_users": 45,
  "active_users": 32,
  "metrics": {
    "documents_filed": 1250,
    "emails_filed": 980,
    "attachments_filed": 270,
    "cases_searched": 3200,
    "suggestions_generated": 1800,
    "suggestions_accepted": 1350,
    "suggestion_acceptance_rate": 0.75
  },
  "top_users": [
    {
      "user_id": "user-123",
      "user_name": "Jane Doe",
      "documents_filed": 120,
      "suggestion_acceptance_rate": 0.85
    }
  ],
  "top_cases": [
    {
      "case_id": "case-456",
      "case_name": "Acme Corp - Q4 Planning",
      "documents_filed": 85,
      "unique_users": 8
    }
  ],
  "filing_by_day": [
    { "date": "2024-02-01", "count": 45 },
    { "date": "2024-02-02", "count": 52 },
    { "date": "2024-02-03", "count": 38 }
  ]
}
```

#### 3. User Activity Summary
```
GET /api/v1/analytics/users/{userId}/activity
```

**Response:**
```json
{
  "user_id": "user-123",
  "user_name": "Jane Doe",
  "period": {
    "from": "2024-02-01T00:00:00Z",
    "to": "2024-02-29T23:59:59Z"
  },
  "summary": {
    "total_actions": 156,
    "documents_filed": 45,
    "cases_accessed": 12,
    "favorite_cases": 3,
    "suggestion_acceptance_rate": 0.82
  },
  "most_active_cases": [
    {
      "case_id": "case-456",
      "case_name": "Acme Corp - Q4 Planning",
      "actions": 28
    }
  ],
  "recent_activity": [
    {
      "timestamp": "2024-02-17T15:30:00Z",
      "action": "document.filed",
      "description": "Filed Contract Review.eml to Acme Corp - Q4 Planning"
    }
  ]
}
```

### Database Queries

#### Log Event
```sql
INSERT INTO audit_logs (
  id,
  timestamp,
  user_id,
  action,
  resource_type,
  resource_id,
  details,
  ip_address,
  user_agent
) VALUES (
  $1,
  NOW(),
  $2,
  $3,
  $4,
  $5,
  $6,
  $7,
  $8
);
```

#### Query Logs with Filters
```sql
SELECT
  al.id,
  al.timestamp,
  al.user_id,
  u.name as user_name,
  al.action,
  al.resource_type,
  al.resource_id,
  al.details,
  al.ip_address,
  al.user_agent
FROM audit_logs al
LEFT JOIN users u ON u.id = al.user_id
WHERE ($1::text IS NULL OR al.user_id = $1)
  AND ($2::text IS NULL OR al.action = $2)
  AND ($3::text IS NULL OR al.resource_type = $3)
  AND ($4::text IS NULL OR al.resource_id = $4)
  AND ($5::timestamp IS NULL OR al.timestamp >= $5)
  AND ($6::timestamp IS NULL OR al.timestamp <= $6)
ORDER BY al.timestamp DESC
LIMIT $7 OFFSET $8;
```

#### Analytics: Documents Filed by Period
```sql
SELECT
  DATE_TRUNC($1, timestamp) as period,
  COUNT(*) FILTER (WHERE action = 'document.filed') as documents_filed,
  COUNT(DISTINCT user_id) as unique_users
FROM audit_logs
WHERE action = 'document.filed'
  AND timestamp >= $2
  AND timestamp <= $3
GROUP BY period
ORDER BY period;
```

#### Analytics: Suggestion Acceptance Rate
```sql
WITH suggestions AS (
  SELECT
    COUNT(*) FILTER (WHERE action = 'suggestion.generated') as generated,
    COUNT(*) FILTER (WHERE action = 'suggestion.accepted') as accepted
  FROM audit_logs
  WHERE action IN ('suggestion.generated', 'suggestion.accepted')
    AND timestamp >= $1
    AND timestamp <= $2
)
SELECT
  generated,
  accepted,
  CASE WHEN generated > 0 THEN accepted::float / generated ELSE 0 END as acceptance_rate
FROM suggestions;
```

#### Analytics: Top Cases by Activity
```sql
SELECT
  c.id,
  c.name,
  c.case_id_visible,
  COUNT(*) as action_count,
  COUNT(DISTINCT al.user_id) as unique_users
FROM audit_logs al
JOIN cases c ON c.id = (al.details->>'case_id')::uuid
WHERE al.timestamp >= $1
  AND al.timestamp <= $2
  AND al.resource_type IN ('document', 'case')
GROUP BY c.id, c.name, c.case_id_visible
ORDER BY action_count DESC
LIMIT 10;
```

### Logging Middleware

Automatically log all API requests:

```javascript
function auditLogMiddleware(req, res, next) {
  // Capture original res.json to intercept response
  const originalJson = res.json.bind(res);

  res.json = function(body) {
    // Log successful operations
    if (res.statusCode >= 200 && res.statusCode < 300) {
      logAuditEvent({
        userId: req.user.id,
        action: mapRouteToAction(req.method, req.path),
        resourceType: extractResourceType(req.path),
        resourceId: extractResourceId(req.path, body),
        details: sanitizeDetails(req.body, body),
        ipAddress: req.ip,
        userAgent: req.get('user-agent')
      });
    }

    return originalJson(body);
  };

  next();
}

function mapRouteToAction(method, path) {
  // POST /api/v1/cases/{id}/documents → "document.filed"
  // PATCH /api/v1/documents/{id} → "document.renamed"
  // GET /api/v1/cases/search → "case.searched"
  // etc.
}
```

### Data Retention & Archival

```javascript
async function archiveOldLogs() {
  const cutoffDate = new Date();
  cutoffDate.setFullYear(cutoffDate.getFullYear() - 2); // 2 years ago

  // 1. Export to archive storage (S3, BigQuery, etc.)
  const oldLogs = await db.query(
    'SELECT * FROM audit_logs WHERE timestamp < $1',
    [cutoffDate]
  );

  await exportToArchive(oldLogs.rows);

  // 2. Delete from primary database
  await db.query(
    'DELETE FROM audit_logs WHERE timestamp < $1',
    [cutoffDate]
  );

  console.log(`Archived ${oldLogs.rows.length} logs older than ${cutoffDate}`);
}

// Run monthly via cron job
```

### Privacy & Redaction

```javascript
function sanitizeDetails(requestBody, responseBody) {
  const sanitized = { ...requestBody };

  // Redact sensitive fields
  const REDACTED_FIELDS = [
    'password',
    'token',
    'secret',
    'data_base64',  // Don't log file content
    'body_preview'  // Don't log email content
  ];

  for (const field of REDACTED_FIELDS) {
    if (sanitized[field]) {
      sanitized[field] = '[REDACTED]';
    }
  }

  return sanitized;
}
```

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` shows logging context:

### Request Logging
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 221-226:** Upload request logging
  ```javascript
  console.log("[uploadDocumentToCase] Starting upload", {
    caseId,
    fileName,
    mimeType,
    dataLength: dataBase64.length,
  });
  ```

- **Lines 283-287:** Response status logging
  ```javascript
  console.log("[uploadDocumentToCase] Fetch completed", {
    status: res.status,
    statusText: res.statusText,
    ok: res.ok,
  });
  ```

- **Lines 305-306:** Success logging
  ```javascript
  console.log("[uploadDocumentToCase] Upload successful", {
    documentIds: (json as any).documents?.map((d: any) => d.id)
  });
  ```

These patterns show what context to capture in audit logs.

### Suggestion Tracking
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/suggestionStorage.ts`:

- **Lines 144-192:** `recordSuccessfulAttach` function
  - Shows what data to track for suggestion analytics
  - Demonstrates learning from user behavior

---

## Dependencies

- **Requires:** BE-001 (Database Schema) - audit_logs table
- **Integrates with:** All other stories - logs all operations

---

## Notes

### Why Audit Logging Matters

1. **Compliance:** GDPR, HIPAA, SOC 2 require audit trails
2. **Security:** Detect unauthorized access or suspicious activity
3. **Debugging:** Troubleshoot user issues by reviewing their actions
4. **Analytics:** Understand product usage patterns
5. **Accountability:** Know who did what and when

### What to Log

**DO Log:**
- User actions (file, rename, delete)
- Authentication events
- Access patterns
- API errors
- Search queries
- System events

**DON'T Log:**
- Passwords or tokens
- Email/document content (privacy)
- Excessive personal data
- High-frequency events (every keystroke)

### Performance Considerations

1. **Async Logging:**
   ```javascript
   // Don't block API response waiting for log write
   logAuditEvent(event).catch(err =>
     console.error('Failed to log audit event:', err)
   );
   ```

2. **Batch Logging:**
   ```javascript
   // Buffer logs and write in batches every 5 seconds
   const logBuffer = [];
   setInterval(async () => {
     if (logBuffer.length > 0) {
       await db.query('INSERT INTO audit_logs ... VALUES ...', logBuffer);
       logBuffer.length = 0;
     }
   }, 5000);
   ```

3. **Partitioning:**
   ```sql
   -- Partition by month for better query performance
   CREATE TABLE audit_logs_2024_02 PARTITION OF audit_logs
   FOR VALUES FROM ('2024-02-01') TO ('2024-03-01');
   ```

### Testing Strategy

```javascript
// Test Cases to Implement

1. Event Logging:
   - File document → logged
   - Rename document → logged
   - Delete document → logged
   - Failed API call → logged (with error)

2. Log Queries:
   - Query by user → correct logs returned
   - Query by date range → correct logs returned
   - Query by action → correct logs returned
   - Pagination → works correctly

3. Analytics:
   - Calculate acceptance rate → correct percentage
   - Top users → sorted correctly
   - Filing by day → counts correct

4. Privacy:
   - Sensitive fields redacted → no passwords/tokens in logs
   - Email content not logged → details don't contain body
   - PII handling → compliant with regulations

5. Performance:
   - Log 1000 events → no API slowdown
   - Query 1M logs → < 500ms response
   - Archive old logs → database size maintained

6. Edge Cases:
   - Logging fails → API still succeeds
   - Database down → logs buffered
   - Invalid data → sanitized before logging
```

### Security Considerations

- **Access Control:** Only admins can query all users' logs
- **Data Encryption:** Encrypt logs at rest
- **Log Tampering:** Prevent modification of logs (append-only)
- **Retention:** Comply with data retention regulations

### Monitoring & Alerting

Track metrics:
- Log write latency
- Log query latency
- Storage size growth
- Failed log writes

Alert on:
- High failed login rate (security)
- Unusual activity patterns (anomaly detection)
- Storage approaching limits
- Log write failures > 1%

### Future Enhancements

- **Real-time Streaming:** Stream logs to SIEM (Splunk, ELK, etc.)
- **Anomaly Detection:** ML-based detection of unusual patterns
- **Compliance Reports:** Auto-generate compliance reports
- **User Notifications:** Alert users of unusual activity on their account
- **Log Replay:** Replay actions for debugging
- **Tamper Detection:** Cryptographic proof of log integrity
- **Export API:** Export logs in various formats (CSV, JSON, PDF)
- **Custom Analytics:** Allow admins to define custom analytics queries
