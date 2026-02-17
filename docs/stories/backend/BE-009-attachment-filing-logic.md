# BE-009: Attachment Filing Logic

**Story ID:** BE-009
**Story Points:** 5
**Epic Link:** Core Filing System
**Priority:** High
**Status:** Ready for Development

---

## Description

Implement logic to handle filing email attachments alongside the email document. Support batch filing, attachment selection, automatic folder organization, and maintain parent-child relationships between emails and their attachments.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

**Key Requirements:**
- File email (.eml) and attachments in single transaction
- Support selective attachment filing (user chooses which attachments)
- Organize into "Outlook add-in" folder automatically
- Link attachments to parent email via metadata
- Handle large attachments (streaming upload)

---

## Acceptance Criteria

- [ ] Filing API (BE-002) supports batch document upload
- [ ] Email and attachments filed together in single transaction
- [ ] Attachments organized into "Outlook add-in" folder
- [ ] Metadata tracks parent email relationship
- [ ] Supports selective filing (user can exclude attachments)
- [ ] Handles attachment name conflicts
- [ ] Rollback entire filing operation if any attachment fails
- [ ] Response includes all filed document IDs
- [ ] HTTP status codes follow BE-002 specification
- [ ] Response time < 3 seconds for email + 5 attachments (50MB total)
- [ ] Integration tests cover all scenarios
- [ ] API documentation with batch examples

---

## Technical Requirements

### Batch Filing Request

Building on BE-002 API, support multiple documents in single request:

```json
{
  "case_id": "case-123",
  "documents": [
    {
      "name": "Project Update.eml",
      "mime_type": "message/rfc822",
      "data_base64": "base64EmailContent...",
      "directory_id": "dir-uuid-outlook-folder",
      "metadata": {
        "conversationId": "AAQkADU...",
        "subject": "Project Update",
        "fromEmail": "john@example.com",
        "isParentEmail": true
      }
    },
    {
      "name": "Budget_Q4.xlsx",
      "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "data_base64": "base64AttachmentContent...",
      "directory_id": "dir-uuid-outlook-folder",
      "metadata": {
        "parentConversationId": "AAQkADU...",
        "attachmentIndex": 0,
        "attachmentOf": "Project Update.eml"
      }
    },
    {
      "name": "Roadmap.pdf",
      "mime_type": "application/pdf",
      "data_base64": "base64AttachmentContent...",
      "directory_id": "dir-uuid-outlook-folder",
      "metadata": {
        "parentConversationId": "AAQkADU...",
        "attachmentIndex": 1,
        "attachmentOf": "Project Update.eml"
      }
    }
  ]
}
```

### Response Structure

```json
{
  "documents": [
    {
      "id": "doc-uuid-email",
      "name": "Project Update.eml",
      "type": "email",
      "mime_type": "message/rfc822",
      "size_bytes": 45678,
      "directory_id": "dir-uuid-outlook-folder",
      "attachments": [
        {
          "id": "doc-uuid-att1",
          "name": "Budget_Q4.xlsx",
          "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "size_bytes": 125000
        },
        {
          "id": "doc-uuid-att2",
          "name": "Roadmap.pdf",
          "mime_type": "application/pdf",
          "size_bytes": 890000
        }
      ]
    }
  ],
  "status": "created",
  "total_documents": 3,
  "total_size_bytes": 1060678
}
```

### Folder Management

#### Ensure "Outlook add-in" Folder Exists

See reference implementation in demo:
`/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts` lines 527-574

```javascript
async function ensureOutlookAddinFolder(caseId) {
  const FOLDER_NAME = "Outlook add-in";

  // 1. Check cache first
  const cachedFolderId = await getFolderIdFromCache(caseId);
  if (cachedFolderId) {
    return cachedFolderId;
  }

  // 2. Get case root directory
  const rootDirId = await getCaseRootDirectory(caseId);
  if (!rootDirId) {
    console.warn('Case has no root directory');
    return null; // Upload to case root
  }

  // 3. List root directory contents
  const listing = await listDirectory(rootDirId);

  // 4. Check if folder already exists
  const existingFolder = listing.items.find(
    item => item.type === 'directory' && item.name === FOLDER_NAME
  );

  if (existingFolder) {
    await cacheFolderId(caseId, existingFolder.id);
    return existingFolder.id;
  }

  // 5. Create folder
  try {
    const created = await createDirectory(rootDirId, FOLDER_NAME);
    await cacheFolderId(caseId, created.id);
    return created.id;
  } catch (error) {
    // If creation fails (e.g., concurrent creation), retry listing
    const retryListing = await listDirectory(rootDirId);
    const retryFolder = retryListing.items.find(
      item => item.type === 'directory' && item.name === FOLDER_NAME
    );
    if (retryFolder) {
      await cacheFolderId(caseId, retryFolder.id);
      return retryFolder.id;
    }
    throw error;
  }
}
```

### Transaction Handling

All documents must be filed together or none at all:

```javascript
async function fileEmailWithAttachments(caseId, emailDoc, attachments, userId) {
  // Start database transaction
  await db.query('BEGIN');

  try {
    // 1. Ensure "Outlook add-in" folder exists
    const folderId = await ensureOutlookAddinFolder(caseId);

    // 2. File email document
    const emailResult = await fileDocument({
      caseId,
      name: emailDoc.name,
      mimeType: emailDoc.mimeType,
      data: emailDoc.data,
      directoryId: folderId,
      metadata: {
        ...emailDoc.metadata,
        isParentEmail: true,
        hasAttachments: attachments.length > 0
      },
      userId
    });

    const emailDocId = emailResult.id;
    const filedAttachments = [];

    // 3. File each attachment
    for (let i = 0; i < attachments.length; i++) {
      const att = attachments[i];

      const attResult = await fileDocument({
        caseId,
        name: att.name,
        mimeType: att.mimeType,
        data: att.data,
        directoryId: folderId,
        metadata: {
          parentEmailDocumentId: emailDocId,
          parentConversationId: emailDoc.metadata.conversationId,
          attachmentIndex: i,
          attachmentOf: emailDoc.name
        },
        userId
      });

      filedAttachments.push(attResult);
    }

    // 4. Update email metadata with attachment IDs
    await db.query(
      `UPDATE document_metadata
       SET attachment_document_ids = $1
       WHERE document_id = $2`,
      [JSON.stringify(filedAttachments.map(a => a.id)), emailDocId]
    );

    // 5. Commit transaction
    await db.query('COMMIT');

    return {
      email: emailResult,
      attachments: filedAttachments
    };

  } catch (error) {
    // Rollback on any error
    await db.query('ROLLBACK');

    // Cleanup any uploaded blobs (if blob storage is separate)
    // await cleanupBlobs([emailDocId, ...attachmentDocIds]);

    throw error;
  }
}
```

### Attachment Metadata Schema

Extend `document_metadata` table (from BE-001):

```sql
ALTER TABLE document_metadata
ADD COLUMN parent_email_document_id UUID REFERENCES documents(id),
ADD COLUMN attachment_index INTEGER,
ADD COLUMN attachment_document_ids JSONB;

CREATE INDEX idx_document_metadata_parent ON document_metadata (parent_email_document_id);
```

### Selective Filing

Client sends only selected attachments:

```json
{
  "case_id": "case-123",
  "email": {
    "name": "Project Update.eml",
    "data_base64": "...",
    "metadata": { ... }
  },
  "attachments": [
    {
      "name": "Budget_Q4.xlsx",
      "data_base64": "...",
      "selected": true
    },
    {
      "name": "Image001.png",
      "data_base64": "...",
      "selected": false  // User unchecked this
    }
  ]
}
```

Server only files attachments where `selected: true`.

### Name Conflict Handling

If attachment name already exists in folder:

**Strategy 1: Auto-Rename**
```javascript
function resolveNameConflict(name, existingNames) {
  if (!existingNames.includes(name)) {
    return name;
  }

  const ext = path.extname(name);
  const base = path.basename(name, ext);

  let counter = 1;
  let newName;

  do {
    newName = `${base} (${counter})${ext}`;
    counter++;
  } while (existingNames.includes(newName));

  return newName;
}

// "Budget.xlsx" → "Budget (1).xlsx" → "Budget (2).xlsx"
```

**Strategy 2: Return Conflict Error**
```json
{
  "error": "name_conflict",
  "message": "Attachment name already exists",
  "conflicts": [
    {
      "name": "Budget.xlsx",
      "existing_document_id": "doc-uuid-existing"
    }
  ]
}
```

**Recommendation:** Use Strategy 1 (auto-rename) for better UX.

### Large Attachment Handling

For attachments > 10MB, consider streaming upload:

```javascript
async function uploadLargeAttachment(attachment, caseId, directoryId) {
  // 1. Create multipart upload
  const uploadId = await initiateMultipartUpload({
    caseId,
    name: attachment.name,
    mimeType: attachment.mimeType,
    directoryId
  });

  // 2. Upload in chunks (5MB each)
  const chunkSize = 5 * 1024 * 1024; // 5MB
  const chunks = [];

  for (let offset = 0; offset < attachment.data.length; offset += chunkSize) {
    const chunk = attachment.data.slice(offset, offset + chunkSize);
    const etag = await uploadPart({
      uploadId,
      partNumber: chunks.length + 1,
      data: chunk
    });
    chunks.push({ partNumber: chunks.length + 1, etag });
  }

  // 3. Complete upload
  const result = await completeMultipartUpload({
    uploadId,
    parts: chunks
  });

  return result;
}
```

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` shows attachment filing logic:

### Folder Management
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 527-574:** `ensureOutlookAddinFolder` function - CRITICAL REFERENCE
  - Cache checking (lines 531-535)
  - Root directory retrieval (lines 538-542)
  - Directory listing (lines 547-549)
  - Existing folder check (lines 552-560)
  - Folder creation (lines 563-568)
  - Idempotent handling (lines 569-573)

- **Lines 361-388:** Folder caching logic
  - `getCachedFolderId` (lines 361-372)
  - `cacheFolderId` (lines 377-388)

- **Lines 424-461:** `listDirectory` function
  - Shows how to query directory contents
  - Response structure parsing (lines 447-455)

- **Lines 466-519:** `createDirectory` function
  - Multiple endpoint attempts (lines 476-480)
  - Error handling (lines 497-505)
  - Response structure (lines 508-512)

### Document Upload Structure
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecaseDocuments.ts`:

- **Lines 206-307:** `uploadDocumentToCase` function
  - Batch upload support: `documents` array (line 251)
  - Metadata structure (lines 252-258)
  - Directory ID specification (line 256)

### Attachment Context
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/MainWorkspace/components/AttachmentsStep.tsx`:

This client component shows expected attachment structure (for reference only - you're building the backend).

---

## Dependencies

- **Requires:** BE-001 (Database Schema) - directories table
- **Requires:** BE-002 (Filing API) - base filing functionality
- **Extends:** BE-004 (Metadata Storage) - attachment metadata

---

## Notes

### Why "Outlook add-in" Folder?

**Benefits:**
1. **Organization:** Keeps Outlook-filed documents separate from manual uploads
2. **Bulk Operations:** Easy to find all emails filed via add-in
3. **Permissions:** Can set different permissions on this folder
4. **Cleanup:** Easy to archive or delete all add-in documents

**Alternative:** File to case root without subfolder (simpler, but less organized)

### Parent-Child Relationships

Link attachments to parent email via metadata:

```javascript
// Email metadata
{
  "conversationId": "AAQkADU...",
  "isParentEmail": true,
  "hasAttachments": true,
  "attachmentDocumentIds": ["doc-uuid-att1", "doc-uuid-att2"]
}

// Attachment metadata
{
  "parentEmailDocumentId": "doc-uuid-email",
  "parentConversationId": "AAQkADU...",
  "attachmentIndex": 0,
  "attachmentOf": "Project Update.eml"
}
```

**Benefits:**
- Navigate from email to attachments
- Navigate from attachment back to email
- Query all attachments for an email
- Maintain relationship even if renamed

### Idempotent Folder Creation

**Problem:** Two users file emails to same case simultaneously → both try to create "Outlook add-in" folder.

**Solution:**
1. Try to create folder
2. If conflict (folder already exists), query again
3. Return existing folder ID

```sql
-- PostgreSQL: Use ON CONFLICT
INSERT INTO directories (id, case_id, parent_id, name)
VALUES ($1, $2, $3, 'Outlook add-in')
ON CONFLICT (case_id, parent_id, name)
DO NOTHING
RETURNING id;

-- If INSERT returned nothing, query existing:
SELECT id FROM directories
WHERE case_id = $1 AND parent_id = $2 AND name = 'Outlook add-in';
```

### Testing Strategy

```javascript
// Test Cases to Implement

1. Basic Filing:
   - Email + 2 attachments → all 3 documents created
   - Email only (no attachments) → 1 document created
   - Attachments without email → error (not supported)

2. Folder Management:
   - First filing → folder created
   - Second filing → folder reused (not recreated)
   - Concurrent filings → no duplicate folders
   - Folder creation fails → documents filed to case root

3. Transaction Handling:
   - Email succeeds, attachment fails → rollback all
   - Blob upload fails → rollback database records
   - Network error → no partial data left

4. Selective Filing:
   - 3 attachments, only 2 selected → 2 filed
   - 0 attachments selected → only email filed

5. Name Conflicts:
   - Attachment name exists → auto-renamed
   - Multiple conflicts → unique names generated

6. Large Attachments:
   - 25MB attachment → filed successfully
   - 100MB total (email + 5 attachments) → filed in < 10 sec

7. Metadata Linking:
   - Email metadata includes attachment IDs
   - Attachment metadata includes parent email ID
   - Query attachments by parent ID → correct results

8. Edge Cases:
   - Attachment without name → generate name
   - Duplicate attachment names in same email → disambiguate
   - Email subject too long → truncate filename
```

### Performance Considerations

1. **Parallel Processing:**
   ```javascript
   // Upload attachments in parallel (max 3 concurrent)
   const CONCURRENCY = 3;
   const results = await Promise.map(
     attachments,
     att => uploadAttachment(att),
     { concurrency: CONCURRENCY }
   );
   ```

2. **Streaming for Large Files:**
   - Don't load entire attachment into memory
   - Stream from client to blob storage
   - Use multipart upload for > 10MB

3. **Caching:**
   - Cache folder ID per case (avoid repeated lookups)
   - Cache TTL: 10 minutes

### Error Handling

```javascript
class AttachmentFilingError extends Error {
  constructor(message, failedAttachments, partialResults) {
    super(message);
    this.failedAttachments = failedAttachments;
    this.partialResults = partialResults;
  }
}

// Response for partial failure
{
  "error": "partial_failure",
  "message": "Some attachments failed to upload",
  "email": {
    "id": "doc-uuid-email",
    "status": "success"
  },
  "attachments": [
    {
      "name": "Budget.xlsx",
      "status": "success",
      "id": "doc-uuid-att1"
    },
    {
      "name": "LargeFile.zip",
      "status": "failed",
      "error": "File too large (exceeds 25MB limit)"
    }
  ]
}
```

**Recommendation:** Use all-or-nothing transaction (no partial success).

### Monitoring & Observability

Track metrics:
- Average attachments per email
- Attachment size distribution
- Folder creation success rate
- Transaction rollback rate
- Upload time by file size

Alert on:
- High rollback rate (> 5%)
- Slow uploads (P95 > 5 seconds)
- Folder creation failures

### Future Enhancements

- **Attachment Deduplication:** Don't re-upload identical files (checksum matching)
- **Attachment Preview:** Generate thumbnails for images/PDFs
- **Attachment Virus Scanning:** Integrate with antivirus API
- **Attachment Compression:** Compress large files before storage
- **Attachment Linking:** Link same attachment filed multiple times
- **Attachment Extraction:** Extract attachments from email body (inline images)
