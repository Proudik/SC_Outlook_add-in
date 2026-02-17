# BE-005: Case Search API

**Story ID:** BE-005
**Story Points:** 5
**Epic Link:** Case Management
**Priority:** High
**Status:** Ready for Development

---

## Description

Implement a fast, flexible case search API to support the add-in's case selector. Users need to quickly find cases by name, visible ID, client name, or keywords. The API must handle partial matches, typos, and return results in < 200ms for excellent UX.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

**Key Use Cases:**
1. User types "acme" → finds "Acme Corp - Q4 Planning"
2. User types "2024-0123" → finds case with visible ID "2024-0123"
3. User types partial name → shows top 10 matches
4. Empty search → returns recent cases + favorites

---

## Acceptance Criteria

- [ ] `GET /api/v1/cases/search` endpoint implemented
- [ ] Supports search by:
  - Case name (partial match)
  - Case visible ID (exact and partial)
  - Client name
  - Keywords
- [ ] Returns top 10 most relevant results by default
- [ ] Response time < 200ms for 95% of requests
- [ ] Handles empty search (returns recent + favorites)
- [ ] Supports pagination for large result sets
- [ ] Case status filtering (active, archived, closed)
- [ ] HTTP status codes:
  - 200 OK - Search completed
  - 400 Bad Request - Invalid query parameters
  - 401 Unauthorized - Missing/invalid token
  - 500 Internal Server Error - Server error
- [ ] Relevance ranking algorithm implemented
- [ ] Integration tests cover all search scenarios
- [ ] API documentation with examples

---

## Technical Requirements

### API Specification

#### Endpoint
```
GET /api/v1/cases/search?q={query}&limit={limit}&offset={offset}&status={status}
```

#### Query Parameters
- `q` (string, optional): Search query (min 1 char for partial match, can be empty)
- `limit` (integer, optional): Max results to return (default: 10, max: 100)
- `offset` (integer, optional): Pagination offset (default: 0)
- `status` (string, optional): Filter by status - "active", "archived", "closed", "all" (default: "active")
- `client_id` (string, optional): Filter by specific client

#### Request Headers
```
Authorization: Bearer <token>
Accept: application/json
```

#### Response Body - With Query
```json
{
  "total": 47,
  "cases": [
    {
      "id": "case-uuid-123",
      "name": "Acme Corp - Q4 Planning",
      "case_id_visible": "2024-0123",
      "client_id": "client-456",
      "client_name": "Acme Corporation",
      "status": "active",
      "created_at": "2024-01-10T08:00:00Z",
      "last_activity_at": "2024-02-15T14:30:00Z",
      "document_count": 342,
      "match_score": 0.95,
      "match_reason": "Exact match on case name"
    },
    {
      "id": "case-uuid-456",
      "name": "Acme Corp - Legal Review",
      "case_id_visible": "2024-0089",
      "client_name": "Acme Corporation",
      "status": "active",
      "match_score": 0.87,
      "match_reason": "Partial match on case name and client"
    }
  ],
  "limit": 10,
  "offset": 0
}
```

#### Response Body - Empty Query (Recent + Favorites)
```json
{
  "total": 8,
  "cases": [
    {
      "id": "case-uuid-789",
      "name": "Beta Inc - Contract Negotiation",
      "case_id_visible": "2024-0234",
      "status": "active",
      "is_favorite": true,
      "last_activity_at": "2024-02-16T10:00:00Z"
    },
    {
      "id": "case-uuid-123",
      "name": "Acme Corp - Q4 Planning",
      "case_id_visible": "2024-0123",
      "status": "active",
      "is_favorite": false,
      "last_activity_at": "2024-02-15T14:30:00Z"
    }
  ],
  "sections": [
    {
      "name": "Favorites",
      "count": 3
    },
    {
      "name": "Recent",
      "count": 5
    }
  ],
  "limit": 10,
  "offset": 0
}
```

### Search Algorithm

#### 1. Exact Match (Highest Priority)
```sql
SELECT *
FROM cases
WHERE (case_id_visible = $1 OR LOWER(name) = LOWER($1))
  AND status = ANY($2)
LIMIT 1;
```

#### 2. Prefix Match
```sql
SELECT *,
       CASE
         WHEN case_id_visible LIKE $1 || '%' THEN 100
         WHEN LOWER(name) LIKE LOWER($1) || '%' THEN 90
         ELSE 0
       END as match_score
FROM cases
WHERE (case_id_visible LIKE $1 || '%' OR LOWER(name) LIKE LOWER($1) || '%')
  AND status = ANY($2)
ORDER BY match_score DESC, last_activity_at DESC
LIMIT 10;
```

#### 3. Partial Match (Trigram Similarity)
PostgreSQL with `pg_trgm` extension:

```sql
SELECT *,
       GREATEST(
         similarity(case_id_visible, $1),
         similarity(name, $1),
         similarity(COALESCE(client_name, ''), $1)
       ) as match_score
FROM cases
WHERE (
  case_id_visible % $1  -- Trigram similarity operator
  OR name % $1
  OR client_name % $1
)
  AND status = ANY($2)
ORDER BY match_score DESC, last_activity_at DESC
LIMIT 10;
```

#### 4. Full-Text Search (Fallback)
```sql
SELECT *,
       ts_rank(
         to_tsvector('english', name || ' ' || COALESCE(case_id_visible, '') || ' ' || COALESCE(client_name, '')),
         plainto_tsquery('english', $1)
       ) as match_score
FROM cases
WHERE to_tsvector('english', name || ' ' || COALESCE(case_id_visible, '') || ' ' || COALESCE(client_name, ''))
      @@ plainto_tsquery('english', $1)
  AND status = ANY($2)
ORDER BY match_score DESC, last_activity_at DESC
LIMIT 10;
```

### Empty Search Logic (No Query)

Return user's recent cases + favorites:

```sql
-- Get user's favorite cases
WITH user_favorites AS (
  SELECT c.*, 1.0 as relevance_score, 'favorite' as source
  FROM cases c
  JOIN user_favorites uf ON uf.case_id = c.id
  WHERE uf.user_id = $1
    AND c.status = 'active'
  ORDER BY uf.position, c.last_activity_at DESC
  LIMIT 5
),
-- Get user's recently accessed cases (from suggestion history or audit logs)
recent_cases AS (
  SELECT DISTINCT c.*, 0.8 as relevance_score, 'recent' as source
  FROM cases c
  JOIN suggestion_history sh ON sh.case_id = c.id
  WHERE sh.user_id = $1
    AND c.status = 'active'
    AND sh.filed_at > NOW() - INTERVAL '30 days'
    AND c.id NOT IN (SELECT case_id FROM user_favorites WHERE user_id = $1)
  ORDER BY sh.filed_at DESC
  LIMIT 5
)
SELECT * FROM user_favorites
UNION ALL
SELECT * FROM recent_cases
ORDER BY relevance_score DESC, last_activity_at DESC
LIMIT 10;
```

### Ranking & Relevance

```javascript
function calculateRelevanceScore(searchQuery, caseRecord) {
  let score = 0;

  // Exact match on visible ID (highest priority)
  if (caseRecord.case_id_visible.toLowerCase() === searchQuery.toLowerCase()) {
    return 1.0;
  }

  // Exact match on name
  if (caseRecord.name.toLowerCase() === searchQuery.toLowerCase()) {
    return 0.95;
  }

  // Prefix match on visible ID
  if (caseRecord.case_id_visible.toLowerCase().startsWith(searchQuery.toLowerCase())) {
    score += 0.4;
  }

  // Prefix match on name
  if (caseRecord.name.toLowerCase().startsWith(searchQuery.toLowerCase())) {
    score += 0.35;
  }

  // Partial match on name (word boundaries)
  const nameWords = caseRecord.name.toLowerCase().split(/\s+/);
  const matchingWords = nameWords.filter(word =>
    word.includes(searchQuery.toLowerCase())
  ).length;
  score += (matchingWords / nameWords.length) * 0.3;

  // Client name match
  if (caseRecord.client_name?.toLowerCase().includes(searchQuery.toLowerCase())) {
    score += 0.2;
  }

  // Boost recent activity
  const daysSinceActivity = (Date.now() - new Date(caseRecord.last_activity_at)) / (1000 * 60 * 60 * 24);
  const recencyBoost = Math.max(0, 1 - (daysSinceActivity / 90)) * 0.1;
  score += recencyBoost;

  return Math.min(score, 1.0);
}
```

### Performance Optimization

1. **Indexes:**
   ```sql
   CREATE INDEX idx_cases_name_trgm ON cases USING GIN (name gin_trgm_ops);
   CREATE INDEX idx_cases_visible_id ON cases (case_id_visible);
   CREATE INDEX idx_cases_status_activity ON cases (status, last_activity_at DESC);
   CREATE INDEX idx_cases_client_id ON cases (client_id);
   CREATE INDEX idx_cases_fts ON cases USING GIN (
     to_tsvector('english', name || ' ' || COALESCE(case_id_visible, '') || ' ' || COALESCE(client_name, ''))
   );
   ```

2. **Caching:**
   - Cache empty search results per user (5 min TTL)
   - Cache popular searches (15 min TTL)
   - Invalidate cache on case creation/update

3. **Query Optimization:**
   - Limit search to active cases by default
   - Use `LIMIT` to prevent large result sets
   - Combine indexes efficiently (visible ID + name in single query)

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` shows the expected case data structure:

### Case Data Structure
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/cases.ts`:

- **Lines 3-6:** `CaseOption` type
  ```typescript
  export type CaseOption = {
    id: string;
    label: string;  // This is the case name
  };
  ```

- **Lines 8-14:** Case mapping function
  - Shows fallback logic for label: `name || case_id_visible || Case ${id}`
  - Demonstrates how to handle missing fields

- **Lines 17-23:** `fetchCasesAll` function
  - Shows filtering by name and client_id
  - Simple structure, but your implementation needs search

### Case Suggestion Engine
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/caseSuggestionEngine.ts`:

- **Lines 20-38:** Case field extraction functions
  - `getCaseVisibleId` (line 21-23): Multiple field fallbacks
  - `getCaseTitle` (line 26-38): Multiple name field locations
  - Shows importance of flexible field access

- **Lines 149-162:** Case reference matching
  - Demonstrates visible ID matching in email content
  - Normalize and compare case IDs

### Recent Cases Storage
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/suggestionStorage.ts`:

- **Lines 7:** `RecentCase` type
  ```typescript
  export type RecentCase = {
    caseId: string;
    lastUsedAt: number;
    useCount: number
  };
  ```

- **Lines 178-184:** Recent cases tracking
  - Update last used timestamp
  - Increment use count
  - Shows what data to track for recent cases

---

## Dependencies

- **Requires:** BE-001 (Database Schema) - cases table must exist
- **Related:** BE-006 (Favorites API) - favorites shown in empty search
- **Related:** BE-007 (Suggestion Engine) - uses recent case data

---

## Notes

### Search UX Considerations

1. **Minimum Query Length:** Allow 1-char searches (e.g., "A" for "Acme")
2. **Debouncing:** Client should debounce keystrokes (300ms recommended)
3. **Empty Results:** Return helpful message and suggest broadening search
4. **Max Results:** Limit to 10 for dropdown, support "Show More" with pagination

### Handling Special Characters

```javascript
function sanitizeSearchQuery(query) {
  // Remove special SQL characters
  query = query.trim();

  // Remove or escape wildcards to prevent SQL injection
  query = query.replace(/[%_]/g, '\\$&');

  // Handle diacritics (optional - normalize to ASCII)
  query = query.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  return query;
}
```

### Case-Insensitive Search

- PostgreSQL: Use `LOWER()` or `ILIKE`
- MySQL: Use `COLLATE utf8mb4_unicode_ci`
- Consider storing normalized name in separate column for faster queries

### Testing Strategy

```javascript
// Test Cases to Implement

1. Exact Matches:
   - Search "2024-0123" → finds case with exact visible ID
   - Search "Acme Corp" → finds exact case name
   - Case-insensitive → "acme corp" finds "Acme Corp"

2. Partial Matches:
   - Search "Acme" → finds all Acme cases
   - Search "2024" → finds all 2024 cases
   - Search "Corp" → finds all Corp cases

3. Empty Search:
   - No query → returns favorites + recent
   - Favorites prioritized over recent
   - Max 10 results

4. Filtering:
   - status=active → only active cases
   - status=all → all cases
   - client_id filter → only that client's cases

5. Pagination:
   - limit=5 → returns 5 results
   - offset=10 → skips first 10 results
   - total count correct

6. Performance:
   - 1000 cases → search completes in < 200ms
   - 10,000 cases → search completes in < 500ms
   - Empty search → completes in < 50ms

7. Edge Cases:
   - Special characters → sanitized
   - Very long query → handled
   - No results → empty array returned
   - Invalid status → 400 Bad Request
```

### Security Considerations

- **Access Control:** Only return cases user has access to
- **SQL Injection:** Always use parameterized queries
- **Rate Limiting:** Max 200 searches/minute per user
- **Query Length:** Limit to 200 chars

### Monitoring & Observability

Track metrics:
- Search response time (P50, P95, P99)
- Cache hit rate
- Empty search rate
- Average query length
- Top searches (for optimization)

Alert on:
- P95 response time > 200ms
- Error rate > 2%
- Cache hit rate < 40%

### Future Enhancements

- **Fuzzy Matching:** Handle typos with Levenshtein distance
- **Synonym Support:** "Contract" matches "Agreement"
- **Autocomplete:** Return suggestions as user types
- **Search History:** Show user's recent searches
- **Smart Ranking:** ML-based relevance ranking
- **Multi-field Search:** "Acme 2024" searches both name and ID
- **Saved Searches:** Save common search queries
