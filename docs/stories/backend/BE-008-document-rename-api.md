# BE-008: Document Rename API

**Story ID:** BE-008
**Story Points:** 3
**Epic Link:** Document Management
**Priority:** Medium
**Status:** Ready for Development

---

## Description

Implement API endpoint to rename filed documents. Users should be able to update document names for better organization without affecting the underlying file content or metadata. Support validation, conflict detection, and audit logging.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

**Use Cases:**
- Rename "message.eml" to "Contract Review - Q4 2024.eml"
- Fix typos in document names
- Standardize naming conventions across case

---

## Acceptance Criteria

- [ ] `PATCH /api/v1/documents/{documentId}` endpoint implemented
- [ ] Updates document name in database
- [ ] Validates new name (length, characters, uniqueness)
- [ ] Prevents concurrent modifications (optimistic locking)
- [ ] Preserves file extension
- [ ] Returns updated document metadata
- [ ] Logs rename action in audit_logs
- [ ] HTTP status codes:
  - 200 OK - Document renamed successfully
  - 400 Bad Request - Invalid name or validation error
  - 401 Unauthorized - Missing/invalid token
  - 403 Forbidden - User lacks permission
  - 404 Not Found - Document doesn't exist
  - 409 Conflict - Document locked or name conflict
  - 500 Internal Server Error - Server error
- [ ] Response time < 100ms for 95% of requests
- [ ] Integration tests cover all scenarios
- [ ] API documentation with examples

---

## Technical Requirements

### API Specification

#### Endpoint
```
PATCH /api/v1/documents/{documentId}
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
  "name": "Contract Review - Q4 2024.eml",
  "preserve_extension": true
}
```

**Fields:**
- `name` (string, required): New document name
- `preserve_extension` (boolean, optional): Keep original extension (default: true)

#### Response Body - Success (200 OK)
```json
{
  "id": "doc-uuid-123",
  "name": "Contract Review - Q4 2024.eml",
  "name_previous": "message.eml",
  "mime_type": "message/rfc822",
  "case_id": "case-456",
  "modified_at": "2024-02-17T16:30:00Z",
  "modified_by_user_id": "user-789",
  "modified_by_user_name": "Jane Doe"
}
```

#### Response Body - Validation Error (400 Bad Request)
```json
{
  "error": "validation_failed",
  "message": "Invalid document name",
  "details": [
    {
      "field": "name",
      "error": "Name cannot contain special characters: / \\ : * ? \" < > |"
    }
  ]
}
```

#### Response Body - Conflict (409 Conflict)
```json
{
  "error": "conflict",
  "message": "Document with this name already exists in the same directory",
  "existing_document_id": "doc-uuid-999"
}
```

### Validation Rules

```javascript
function validateDocumentName(name, preserveExtension, originalName) {
  const errors = [];

  // Required
  if (!name || name.trim() === '') {
    errors.push('Name is required');
    return errors;
  }

  const trimmedName = name.trim();

  // Length limits
  if (trimmedName.length > 255) {
    errors.push('Name exceeds maximum length of 255 characters');
  }

  if (trimmedName.length < 1) {
    errors.push('Name must be at least 1 character');
  }

  // Forbidden characters (Windows/Unix file system restrictions)
  const forbiddenChars = /[\/\\:*?"<>|]/g;
  if (forbiddenChars.test(trimmedName)) {
    errors.push('Name cannot contain special characters: / \\ : * ? " < > |');
  }

  // Cannot start or end with space or period
  if (trimmedName.startsWith(' ') || trimmedName.startsWith('.')) {
    errors.push('Name cannot start with space or period');
  }

  if (trimmedName.endsWith(' ') || trimmedName.endsWith('.')) {
    errors.push('Name cannot end with space or period');
  }

  // Preserve extension if requested
  if (preserveExtension && originalName) {
    const originalExt = path.extname(originalName);
    const newExt = path.extname(trimmedName);

    if (originalExt && newExt !== originalExt) {
      errors.push(`Extension must be preserved: ${originalExt}`);
    }
  }

  // Reserved names (Windows)
  const reservedNames = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'LPT1'];
  const baseName = path.basename(trimmedName, path.extname(trimmedName));
  if (reservedNames.includes(baseName.toUpperCase())) {
    errors.push('Name uses reserved system name');
  }

  return errors;
}
```

### Database Operations

#### Get Document for Update
```sql
SELECT
  d.id,
  d.name,
  d.case_id,
  d.directory_id,
  d.modified_at,
  d.created_by_user_id
FROM documents d
WHERE d.id = $1
FOR UPDATE; -- Lock row for update
```

#### Check Name Uniqueness
```sql
SELECT id
FROM documents
WHERE case_id = $1
  AND directory_id IS NOT DISTINCT FROM $2
  AND LOWER(name) = LOWER($3)
  AND id != $4
LIMIT 1;
```

**Note:** `IS NOT DISTINCT FROM` handles NULL directory_id comparison correctly.

#### Update Document Name
```sql
UPDATE documents
SET
  name = $1,
  modified_at = NOW(),
  modified_by_user_id = $2
WHERE id = $3
  AND modified_at = $4  -- Optimistic locking
RETURNING *;
```

#### Log Audit Entry
```sql
INSERT INTO audit_logs (
  id,
  timestamp,
  user_id,
  action,
  resource_type,
  resource_id,
  details
) VALUES (
  $1,
  NOW(),
  $2,
  'document.renamed',
  'document',
  $3,
  $4  -- JSON: { from: 'old.eml', to: 'new.eml', case_id: '...' }
);
```

### Implementation Logic

```javascript
async function renameDocument(documentId, newName, userId, preserveExtension = true) {
  // 1. Start transaction
  await db.query('BEGIN');

  try {
    // 2. Get and lock document
    const docResult = await db.query(
      'SELECT id, name, case_id, directory_id, modified_at FROM documents WHERE id = $1 FOR UPDATE',
      [documentId]
    );

    if (!docResult.rows.length) {
      throw new NotFoundError('Document not found');
    }

    const doc = docResult.rows[0];

    // 3. Validate new name
    const errors = validateDocumentName(newName, preserveExtension, doc.name);
    if (errors.length) {
      throw new ValidationError('Invalid document name', errors);
    }

    // 4. Preserve extension if needed
    let finalName = newName.trim();
    if (preserveExtension) {
      const originalExt = path.extname(doc.name);
      const newExt = path.extname(finalName);

      if (originalExt && !newExt) {
        finalName = finalName + originalExt;
      }
    }

    // 5. Check for name conflicts in same directory
    const conflictResult = await db.query(
      `SELECT id FROM documents
       WHERE case_id = $1
         AND directory_id IS NOT DISTINCT FROM $2
         AND LOWER(name) = LOWER($3)
         AND id != $4
       LIMIT 1`,
      [doc.case_id, doc.directory_id, finalName, documentId]
    );

    if (conflictResult.rows.length) {
      throw new ConflictError('Document with this name already exists', {
        existingDocumentId: conflictResult.rows[0].id
      });
    }

    // 6. Update document name
    const updateResult = await db.query(
      `UPDATE documents
       SET name = $1, modified_at = NOW(), modified_by_user_id = $2
       WHERE id = $3 AND modified_at = $4
       RETURNING *`,
      [finalName, userId, documentId, doc.modified_at]
    );

    if (!updateResult.rows.length) {
      throw new ConflictError('Document was modified by another user');
    }

    // 7. Log audit entry
    await db.query(
      `INSERT INTO audit_logs (id, timestamp, user_id, action, resource_type, resource_id, details)
       VALUES ($1, NOW(), $2, 'document.renamed', 'document', $3, $4)`,
      [
        uuidv4(),
        userId,
        documentId,
        JSON.stringify({
          from: doc.name,
          to: finalName,
          case_id: doc.case_id
        })
      ]
    );

    // 8. Commit transaction
    await db.query('COMMIT');

    return updateResult.rows[0];

  } catch (error) {
    // Rollback on any error
    await db.query('ROLLBACK');
    throw error;
  }
}
```

### Optimistic Locking

Prevent lost updates when multiple users try to rename simultaneously:

```sql
-- Update only if modified_at hasn't changed
UPDATE documents
SET name = $1, modified_at = NOW()
WHERE id = $2 AND modified_at = $3
RETURNING *;
```

If `modified_at` changed, update fails (0 rows affected) → return 409 Conflict.

### Extension Preservation

```javascript
function preserveExtension(newName, originalName) {
  const originalExt = path.extname(originalName); // '.eml'
  const newExt = path.extname(newName);

  // If new name has no extension, append original
  if (originalExt && !newExt) {
    return newName + originalExt;
  }

  // If extensions differ, warn or reject
  if (originalExt && newExt && newExt !== originalExt) {
    throw new ValidationError(`Extension must be ${originalExt}`);
  }

  return newName;
}
```

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` shows rename functionality:

### Rename Document Logic
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 332-350:** `renameDocument` function
  - Multiple endpoint attempts (lines 336-342)
  - Token-based authentication (line 310)
  - Error handling for unsupported endpoints (line 349)

- **Lines 309-330:** `tryRename` helper
  - Request structure (lines 310-318)
  - JSON response handling (lines 321-323)
  - 404/405 handling for endpoint discovery (line 326)
  - Error message formatting (line 329)

### API Endpoint Discovery Pattern
The demo tries multiple endpoints to handle API variations:

```javascript
const candidates = [
  { url: `${base}/documents/${id}`, method: 'PUT' },
  { url: `${base}/documents/${id}`, method: 'PATCH' },
  { url: `${base}/documents/${id}/rename`, method: 'POST' },
  { url: `${base}/documents/${id}/name`, method: 'PUT' },
];
```

**Your implementation should use:** `PATCH /documents/{id}` (RESTful standard)

---

## Dependencies

- **Requires:** BE-001 (Database Schema) - documents table
- **Requires:** BE-010 (Audit Logging) - audit_logs table
- **Related:** BE-002 (Filing API) - documents to rename come from filing

---

## Notes

### Why Rename Matters

Email attachments often have generic names:
- "message.eml"
- "Document.pdf"
- "image001.png"

Users need descriptive names for:
- Better organization
- Easier searching
- Document management compliance

### Concurrency Handling

**Scenario:** Two users rename same document simultaneously.

**Solution 1: Optimistic Locking (Recommended)**
```javascript
// Update only if modified_at unchanged
UPDATE documents
SET name = $1, modified_at = NOW()
WHERE id = $2 AND modified_at = $3;

// If 0 rows updated → conflict, return 409
```

**Solution 2: Pessimistic Locking**
```javascript
// Lock row during read
SELECT * FROM documents WHERE id = $1 FOR UPDATE;
// Other transactions wait
```

**Recommendation:** Use optimistic locking for better performance.

### Name Sanitization

```javascript
function sanitizeFilename(name) {
  // Remove forbidden characters
  name = name.replace(/[\/\\:*?"<>|]/g, '_');

  // Remove leading/trailing spaces and periods
  name = name.trim().replace(/^\.+/, '').replace(/\.+$/, '');

  // Collapse multiple spaces
  name = name.replace(/\s+/g, ' ');

  // Limit length
  if (name.length > 255) {
    const ext = path.extname(name);
    const base = path.basename(name, ext);
    name = base.substring(0, 255 - ext.length) + ext;
  }

  return name;
}
```

### Testing Strategy

```javascript
// Test Cases to Implement

1. Valid Renames:
   - "message.eml" → "Contract Review.eml" (success)
   - "doc.pdf" → "Document" (adds .pdf extension)
   - "Test" → "Test " (trims trailing space)

2. Validation Errors:
   - Empty name → 400
   - Name with "/" → 400
   - Name > 255 chars → 400
   - Reserved name "CON.eml" → 400

3. Conflicts:
   - Name already exists in directory → 409
   - Document locked by another user → 409
   - Modified since read → 409

4. Extension Handling:
   - "doc.eml" → "doc.pdf" (rejects if preserve=true)
   - "doc" → "doc.eml" (adds extension)
   - "doc.txt" → "doc" (removes extension if preserve=false)

5. Concurrency:
   - Two users rename simultaneously → one succeeds, one gets 409
   - User renames while another deletes → 404

6. Permissions:
   - User without access → 403
   - Document in archived case → success (if allowed)

7. Audit Logging:
   - Successful rename → logged
   - Failed rename → not logged
   - Log contains old and new names
```

### Security Considerations

- **Authorization:** Verify user has write access to case
- **Path Traversal:** Prevent "../../../etc/passwd" in names
- **SQL Injection:** Always use parameterized queries
- **XSS:** Sanitize names when displayed in UI

### Performance Optimization

1. **Indexes:**
   ```sql
   CREATE INDEX idx_documents_case_dir_name ON documents (case_id, directory_id, LOWER(name));
   ```

2. **Query Optimization:**
   - Use `FOR UPDATE` only when necessary
   - Keep transaction short
   - Check uniqueness before updating

3. **Caching:**
   - Invalidate document cache on rename
   - Update any cached directory listings

### Monitoring & Observability

Track metrics:
- Rename success rate
- Validation error rate by type
- Conflict rate (concurrent renames)
- Average rename time

Alert on:
- High conflict rate (> 5%)
- Slow renames (P95 > 500ms)
- Error rate > 2%

### Future Enhancements

- **Bulk Rename:** Rename multiple documents at once
- **Rename Patterns:** Apply naming templates (e.g., "Email - {sender} - {date}")
- **Auto-Naming:** Suggest names based on email subject
- **Rename History:** Track all previous names (versions)
- **Undo Rename:** Restore previous name within 24 hours
- **Name Validation Rules:** Custom rules per case/client
