# BE-006: Favourites Management API

**Story ID:** BE-006
**Story Points:** 3
**Epic Link:** Case Management
**Priority:** Medium
**Status:** Ready for Development

---

## Description

Implement API endpoints to manage user's favorite cases. Favorites provide quick access to frequently used cases and appear at the top of the case selector. Users should be able to add, remove, reorder, and list their favorites.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

---

## Acceptance Criteria

- [ ] `GET /api/v1/users/{userId}/favorites` - List user's favorite cases
- [ ] `POST /api/v1/users/{userId}/favorites` - Add case to favorites
- [ ] `DELETE /api/v1/users/{userId}/favorites/{caseId}` - Remove from favorites
- [ ] `PUT /api/v1/users/{userId}/favorites/order` - Reorder favorites
- [ ] Response time < 50ms for all operations
- [ ] Favorites persist across sessions
- [ ] Max 20 favorites per user
- [ ] Favorites automatically removed if case is deleted
- [ ] HTTP status codes:
  - 200 OK - Operation successful
  - 201 Created - Favorite added
  - 204 No Content - Favorite removed
  - 400 Bad Request - Invalid request
  - 401 Unauthorized - Missing/invalid token
  - 404 Not Found - Case or user not found
  - 409 Conflict - Favorite already exists
  - 500 Internal Server Error - Server error
- [ ] Integration tests cover all scenarios
- [ ] API documentation with examples

---

## Technical Requirements

### API Specification

#### 1. List User's Favorites
```
GET /api/v1/users/{userId}/favorites
```

**Response:**
```json
{
  "favorites": [
    {
      "id": "fav-uuid-1",
      "case_id": "case-123",
      "case_name": "Acme Corp - Q4 Planning",
      "case_id_visible": "2024-0123",
      "case_status": "active",
      "client_name": "Acme Corporation",
      "position": 0,
      "created_at": "2024-01-15T10:00:00Z",
      "last_activity_at": "2024-02-15T14:30:00Z"
    },
    {
      "id": "fav-uuid-2",
      "case_id": "case-456",
      "case_name": "Beta Inc - Contract Review",
      "case_id_visible": "2024-0234",
      "case_status": "active",
      "position": 1,
      "created_at": "2024-02-01T09:00:00Z"
    }
  ],
  "total": 2,
  "max_favorites": 20
}
```

**Query Parameters:**
- `include_archived` (boolean, optional): Include archived cases (default: false)

#### 2. Add Case to Favorites
```
POST /api/v1/users/{userId}/favorites
```

**Request Body:**
```json
{
  "case_id": "case-789"
}
```

**Response (201 Created):**
```json
{
  "id": "fav-uuid-3",
  "case_id": "case-789",
  "case_name": "Gamma LLC - Due Diligence",
  "position": 2,
  "created_at": "2024-02-17T15:45:00Z"
}
```

**Response (409 Conflict - Already Favorite):**
```json
{
  "error": "conflict",
  "message": "Case is already in favorites",
  "favorite": {
    "id": "fav-uuid-3",
    "position": 2
  }
}
```

**Response (400 Bad Request - Limit Reached):**
```json
{
  "error": "limit_exceeded",
  "message": "Maximum of 20 favorites allowed. Remove a favorite before adding a new one.",
  "current_count": 20,
  "max_favorites": 20
}
```

#### 3. Remove Case from Favorites
```
DELETE /api/v1/users/{userId}/favorites/{caseId}
```

**Response (204 No Content):**
```
(empty body)
```

**Response (404 Not Found):**
```json
{
  "error": "not_found",
  "message": "Case is not in favorites"
}
```

#### 4. Reorder Favorites
```
PUT /api/v1/users/{userId}/favorites/order
```

**Request Body:**
```json
{
  "order": [
    { "case_id": "case-456", "position": 0 },
    { "case_id": "case-123", "position": 1 },
    { "case_id": "case-789", "position": 2 }
  ]
}
```

**Response (200 OK):**
```json
{
  "message": "Favorites reordered successfully",
  "favorites": [
    {
      "case_id": "case-456",
      "position": 0
    },
    {
      "case_id": "case-123",
      "position": 1
    },
    {
      "case_id": "case-789",
      "position": 2
    }
  ]
}
```

### Database Operations

#### List Favorites
```sql
SELECT
  uf.id,
  uf.case_id,
  uf.position,
  uf.created_at,
  c.name as case_name,
  c.case_id_visible,
  c.status as case_status,
  c.client_name,
  c.last_activity_at
FROM user_favorites uf
JOIN cases c ON c.id = uf.case_id
WHERE uf.user_id = $1
  AND ($2 = TRUE OR c.status != 'archived')  -- include_archived flag
ORDER BY uf.position ASC, uf.created_at ASC;
```

#### Add Favorite
```sql
-- Check if already exists
SELECT id FROM user_favorites
WHERE user_id = $1 AND case_id = $2;

-- Check count limit
SELECT COUNT(*) FROM user_favorites
WHERE user_id = $1;

-- Insert new favorite (if checks pass)
INSERT INTO user_favorites (id, user_id, case_id, position, created_at)
VALUES ($1, $2, $3, $4, NOW())
RETURNING *;
```

#### Remove Favorite
```sql
DELETE FROM user_favorites
WHERE user_id = $1 AND case_id = $2
RETURNING id;

-- Reindex positions after deletion (optional, can be done lazily)
UPDATE user_favorites
SET position = position - 1
WHERE user_id = $1 AND position > $3;
```

#### Reorder Favorites
```sql
-- Update all positions in a transaction
BEGIN;

UPDATE user_favorites
SET position = CASE case_id
  WHEN 'case-456' THEN 0
  WHEN 'case-123' THEN 1
  WHEN 'case-789' THEN 2
  ELSE position
END
WHERE user_id = $1
  AND case_id IN ('case-456', 'case-123', 'case-789');

COMMIT;
```

### Business Logic

#### Adding Favorites
```javascript
async function addFavorite(userId, caseId) {
  // 1. Verify case exists and user has access
  const caseExists = await db.query(
    'SELECT id FROM cases WHERE id = $1',
    [caseId]
  );
  if (!caseExists.rows.length) {
    throw new NotFoundError('Case not found');
  }

  // 2. Check if already favorite
  const existing = await db.query(
    'SELECT id FROM user_favorites WHERE user_id = $1 AND case_id = $2',
    [userId, caseId]
  );
  if (existing.rows.length) {
    throw new ConflictError('Case is already in favorites');
  }

  // 3. Check favorites limit
  const countResult = await db.query(
    'SELECT COUNT(*) FROM user_favorites WHERE user_id = $1',
    [userId]
  );
  const currentCount = parseInt(countResult.rows[0].count);
  if (currentCount >= 20) {
    throw new BadRequestError('Maximum of 20 favorites allowed');
  }

  // 4. Get next position (append to end)
  const maxPosition = currentCount; // 0-indexed

  // 5. Insert favorite
  const result = await db.query(
    'INSERT INTO user_favorites (id, user_id, case_id, position, created_at) VALUES ($1, $2, $3, $4, NOW()) RETURNING *',
    [uuidv4(), userId, caseId, maxPosition]
  );

  return result.rows[0];
}
```

#### Removing Favorites
```javascript
async function removeFavorite(userId, caseId) {
  const result = await db.query(
    'DELETE FROM user_favorites WHERE user_id = $1 AND case_id = $2 RETURNING id, position',
    [userId, caseId]
  );

  if (!result.rows.length) {
    throw new NotFoundError('Case is not in favorites');
  }

  // Optional: Reindex positions to fill gap
  // (Can be done lazily on next list/reorder operation)
  const deletedPosition = result.rows[0].position;
  await db.query(
    'UPDATE user_favorites SET position = position - 1 WHERE user_id = $1 AND position > $2',
    [userId, deletedPosition]
  );

  return { deleted: true };
}
```

#### Reordering Favorites
```javascript
async function reorderFavorites(userId, orderArray) {
  // Validate all case IDs belong to user's favorites
  const favoriteIds = await db.query(
    'SELECT case_id FROM user_favorites WHERE user_id = $1',
    [userId]
  );
  const validIds = new Set(favoriteIds.rows.map(r => r.case_id));

  for (const item of orderArray) {
    if (!validIds.has(item.case_id)) {
      throw new BadRequestError(`Case ${item.case_id} is not in favorites`);
    }
  }

  // Build CASE statement for bulk update
  const updates = orderArray
    .map((item, idx) => `WHEN '${item.case_id}' THEN ${idx}`)
    .join(' ');

  const caseIds = orderArray.map(item => item.case_id);

  await db.query(`
    UPDATE user_favorites
    SET position = CASE case_id ${updates} ELSE position END
    WHERE user_id = $1 AND case_id = ANY($2)
  `, [userId, caseIds]);

  return { success: true };
}
```

### Cascade Deletion

When a case is deleted, automatically remove from all users' favorites:

```sql
-- Add foreign key with cascade
ALTER TABLE user_favorites
ADD CONSTRAINT fk_user_favorites_case
FOREIGN KEY (case_id)
REFERENCES cases(id)
ON DELETE CASCADE;
```

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` doesn't implement favorites API (client-side only), but shows the expected UX:

### Recent Cases Tracking
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/suggestionStorage.ts`:

- **Lines 7:** `RecentCase` type - similar structure for favorites
- **Lines 178-184:** Tracking recent usage
  - Shows how to update timestamps and counts
  - Similar pattern for favorites

### Case Data Structure
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/cases.ts`:

- **Lines 3-6:** `CaseOption` type
  - Favorites API should return similar structure
  - Include case name and ID

---

## Dependencies

- **Requires:** BE-001 (Database Schema) - user_favorites table must exist
- **Integrates with:** BE-005 (Case Search) - favorites shown in empty search
- **Related:** BE-007 (Suggestion Engine) - favorites can boost suggestions

---

## Notes

### UX Considerations

1. **Quick Add:** Support one-click favorite toggle from case selector
2. **Drag-and-Drop:** Reorder endpoint enables drag-and-drop UI
3. **Visual Indicator:** Show star icon or badge on favorited cases
4. **Max Limit:** Show remaining slots (e.g., "5 of 20 favorites")

### Position Management Strategies

**Option 1: Dense Positions (0, 1, 2, 3...)**
- Pros: Simple, sequential
- Cons: Reindexing needed on deletion

**Option 2: Sparse Positions (0, 100, 200, 300...)**
- Pros: Can insert between without reindexing
- Cons: Eventually need reindexing

**Recommendation:** Use dense positions with lazy reindexing.

### Performance Optimization

1. **Indexes:**
   ```sql
   CREATE INDEX idx_user_favorites_user_position ON user_favorites (user_id, position);
   CREATE UNIQUE INDEX idx_user_favorites_user_case ON user_favorites (user_id, case_id);
   ```

2. **Caching:**
   - Cache user's favorites list (5 min TTL)
   - Invalidate on add/remove/reorder
   - Cache key: `favorites:${userId}`

3. **Query Optimization:**
   - Join with cases table only when needed
   - Use `LIMIT 20` to prevent over-fetching

### Testing Strategy

```javascript
// Test Cases to Implement

1. List Favorites:
   - Empty favorites → returns empty array
   - 3 favorites → returns all 3 in order
   - include_archived=true → includes archived cases

2. Add Favorite:
   - Valid case → 201 Created
   - Already favorite → 409 Conflict
   - 20 favorites exist → 400 Bad Request (limit)
   - Case doesn't exist → 404 Not Found

3. Remove Favorite:
   - Exists → 204 No Content
   - Doesn't exist → 404 Not Found
   - Positions reindexed → correct order

4. Reorder:
   - Valid reorder → positions updated
   - Invalid case ID → 400 Bad Request
   - Partial reorder → only specified cases moved

5. Cascade Deletion:
   - Delete case → removed from all users' favorites
   - Position gaps filled correctly

6. Concurrency:
   - Two users add same case → both succeed
   - User adds twice simultaneously → only one succeeds (409)
```

### Security Considerations

- **Authorization:** Verify userId in token matches {userId} in URL
- **Access Control:** Verify user has access to case before allowing favorite
- **Rate Limiting:** Max 100 operations/minute per user
- **Injection Prevention:** Parameterized queries for case IDs

### Monitoring & Observability

Track metrics:
- Favorites per user (avg, median, P95)
- Most favorited cases
- Reorder frequency
- Add/remove rate

Alert on:
- High error rate (> 2%)
- Unusual favorite counts (abuse detection)

### Future Enhancements

- **Shared Favorites:** Team favorites visible to all members
- **Favorite Folders:** Organize favorites into folders
- **Smart Favorites:** Auto-suggest cases to favorite based on usage
- **Favorite Notes:** Add personal notes to favorites
- **Export/Import:** Export favorites for backup
- **Sync Across Devices:** Real-time sync with WebSocket
