# BE-007: Suggestion Intelligence Engine

**Story ID:** BE-007
**Story Points:** 8
**Epic Link:** Smart Suggestions
**Priority:** High
**Status:** Ready for Development

---

## Description

Implement an intelligent case suggestion API that analyzes email context (sender, subject, content) and user history to recommend the most relevant cases for filing. The engine should learn from user behavior over time and provide confidence scores for each suggestion.

This is a greenfield implementation - you are building from scratch, not modifying existing code.

**Key Intelligence Signals:**
1. **Sender History:** "You often file emails from john@acme.com to Case 2024-0123"
2. **Email Thread Context:** "This reply is part of a thread you filed before"
3. **Subject/Content Matching:** "Case name 'Q4 Planning' appears in email subject"
4. **Domain Patterns:** "Emails from @acme.com usually go to Case 2024-0123"
5. **Recent Activity:** "You recently filed 5 emails to this case"

---

## Acceptance Criteria

- [ ] `POST /api/v1/suggestions/cases` endpoint implemented
- [ ] Returns top 2-5 case suggestions with confidence scores
- [ ] Considers multiple signals: sender, domain, subject, thread, recency
- [ ] Auto-selects case if confidence > 70%
- [ ] Learns from user behavior (tracks successful filings)
- [ ] Response time < 200ms for 95% of requests
- [ ] HTTP status codes:
  - 200 OK - Suggestions returned (may be empty array)
  - 400 Bad Request - Invalid request
  - 401 Unauthorized - Missing/invalid token
  - 500 Internal Server Error - Server error
- [ ] Confidence algorithm validated with test data
- [ ] Handles edge cases (new sender, empty body, etc.)
- [ ] Integration tests cover suggestion scenarios
- [ ] API documentation with scoring explanation

---

## Technical Requirements

### API Specification

#### Endpoint
```
POST /api/v1/suggestions/cases
```

#### Request Body
```json
{
  "user_id": "user-123",
  "email_context": {
    "conversation_id": "AAQkADU1YjJj...",
    "subject": "Re: Q4 Planning - Budget Review",
    "body_snippet": "Here are the updated budget figures for Q4...",
    "from_email": "john@acme.com",
    "from_name": "John Doe",
    "to_emails": ["jane@company.com"],
    "attachment_names": ["Budget_Q4_2024.xlsx"],
    "date_sent": "2024-02-17T10:30:00Z"
  },
  "top_k": 2,
  "min_confidence": 10
}
```

**Fields:**
- `user_id` (string): Current user's ID
- `email_context` (object): Email metadata for analysis
- `top_k` (integer, optional): Max suggestions to return (default: 2, max: 5)
- `min_confidence` (integer, optional): Minimum confidence % to include (default: 10)

#### Response Body
```json
{
  "suggestions": [
    {
      "case_id": "case-123",
      "case_name": "Acme Corp - Q4 Planning",
      "case_id_visible": "2024-0123",
      "confidence_pct": 87,
      "score": 142.5,
      "reasons": [
        "Case name matches email subject",
        "You often attach emails from this sender to this case",
        "Recently used (5 emails filed in last 7 days)"
      ],
      "signals": {
        "subject_match": 98,
        "sender_history": 45,
        "recent_activity": 12
      }
    },
    {
      "case_id": "case-456",
      "case_name": "Acme Corp - Legal Review",
      "case_id_visible": "2024-0089",
      "confidence_pct": 52,
      "score": 78.3,
      "reasons": [
        "This domain often maps to this case",
        "Case name mentioned in email body"
      ],
      "signals": {
        "domain_history": 40,
        "body_mention": 38
      }
    }
  ],
  "auto_select_case_id": "case-123",
  "total_cases_analyzed": 247
}
```

**Response Fields:**
- `suggestions`: Array of case suggestions sorted by confidence (high to low)
- `auto_select_case_id`: Case ID to auto-select (if confidence >= 70%), null otherwise
- `total_cases_analyzed`: Total cases considered

### Scoring Algorithm

#### Scoring Components

```javascript
const SCORING_WEIGHTS = {
  // Strongest signals
  thread_context: 100,        // Same thread filed before
  case_id_reference: 95,      // Case ID in subject/body
  subject_exact_match: 98,    // Subject matches case name exactly

  // Strong signals
  subject_partial_match: 60,  // Subject contains case name
  sender_history_strong: 50,  // 10+ emails from sender to case
  body_case_mention: 45,      // Case name in email body

  // Medium signals
  domain_history: 40,         // Domain frequently maps to case
  sender_history_medium: 30,  // 3-9 emails from sender to case
  attachment_reference: 25,   // Attachment name mentions case

  // Weak signals
  recent_activity: 12,        // Case used recently
  sender_history_weak: 10,    // 1-2 emails from sender to case
};
```

#### Confidence Calculation

```javascript
function calculateConfidence(sortedScores, index) {
  const score = sortedScores[index] || 0;
  const topScore = sortedScores[0] || 0;
  const secondScore = sortedScores[1] || 0;

  // Base confidence: normalize score to 0-1 range (120 is typical good score)
  const baseConfidence = Math.min(score / 120, 1.0);

  // Separation bonus: reward clear separation from other suggestions
  const referenceScore = (index === 0) ? secondScore : topScore;
  const gap = Math.max(0, score - referenceScore);
  const separationBonus = Math.min(gap / 60, 1.0);

  // Weighted combination
  const confidence = (0.65 * baseConfidence) + (0.35 * separationBonus);

  return Math.round(confidence * 100); // Convert to percentage
}
```

#### Core Algorithm

```javascript
async function suggestCases(userId, emailContext, allCases, topK = 2) {
  const scores = {};
  const reasons = {};

  // 1. Thread Context (Strongest Signal)
  if (emailContext.conversation_id) {
    const threadCase = await getThreadMappedCase(userId, emailContext.conversation_id);
    if (threadCase) {
      addScore(scores, reasons, threadCase.case_id, 100,
        'Same email thread previously attached to this case');
    }
  }

  // 2. Case ID Reference in Content
  const normalizedSubject = normalize(emailContext.subject);
  const normalizedBody = normalize(emailContext.body_snippet);

  for (const caseData of allCases) {
    const caseId = caseData.id;
    const caseIdVisible = normalize(caseData.case_id_visible);

    if (normalizedSubject.includes(caseIdVisible) ||
        normalizedBody.includes(caseIdVisible)) {
      addScore(scores, reasons, caseId, 95,
        'Case reference found in the email');
    }
  }

  // 3. Case Name Matching
  for (const caseData of allCases) {
    const caseId = caseData.id;
    const caseName = normalize(caseData.name);
    const caseTokens = tokenize(caseName);

    // Exact match
    if (normalizedSubject === caseName) {
      addScore(scores, reasons, caseId, 98,
        'Email subject matches the case name');
      continue;
    }

    // Partial match in subject
    const subjectOverlap = calculateTokenOverlap(caseTokens, normalizedSubject);
    if (subjectOverlap.hits >= 2) {
      const matchQuality = subjectOverlap.hits / subjectOverlap.total;
      const points = 60 + (30 * matchQuality); // 60-90 points
      addScore(scores, reasons, caseId, points,
        'Case name matches the email subject');
    }

    // Case mention in body
    const bodyOverlap = calculateTokenOverlap(caseTokens, normalizedBody);
    if (bodyOverlap.hits >= 2) {
      const matchQuality = bodyOverlap.hits / bodyOverlap.total;
      const points = 35 + (25 * matchQuality); // 35-60 points
      addScore(scores, reasons, caseId, points,
        'Case name mentioned in the email body');
    }
  }

  // 4. Sender History
  if (emailContext.from_email) {
    const senderHistory = await getSenderFilingHistory(
      userId,
      emailContext.from_email,
      90 // days
    );

    for (const [caseId, count] of Object.entries(senderHistory)) {
      let points, reason;
      if (count >= 10) {
        points = 50;
        reason = 'You often attach emails from this sender to this case';
      } else if (count >= 3) {
        points = 30;
        reason = 'You sometimes attach emails from this sender to this case';
      } else {
        points = 10;
        reason = 'You previously attached email from this sender to this case';
      }
      addScore(scores, reasons, caseId, points * Math.log1p(count), reason);
    }
  }

  // 5. Domain History
  const domain = extractDomain(emailContext.from_email);
  if (domain) {
    const domainHistory = await getDomainFilingHistory(userId, domain, 90);

    for (const [caseId, count] of Object.entries(domainHistory)) {
      const points = 20 * Math.log1p(count); // Logarithmic scaling
      addScore(scores, reasons, caseId, points,
        'This domain often maps to this case');
    }
  }

  // 6. Recent Activity
  const recentCases = await getRecentCases(userId, 14); // Last 14 days
  for (const recentCase of recentCases) {
    const daysAgo = (Date.now() - recentCase.last_used_at) / (1000 * 60 * 60 * 24);
    const decay = Math.max(0, 1 - (daysAgo / 14));
    const points = 12 * decay;
    if (points > 0) {
      addScore(scores, reasons, recentCase.case_id, points, 'Recently used');
    }
  }

  // Sort by score and calculate confidence
  const sorted = Object.entries(scores)
    .map(([caseId, score]) => ({
      case_id: caseId,
      score,
      reasons: reasons[caseId] || []
    }))
    .sort((a, b) => b.score - a.score);

  const sortedScores = sorted.map(s => s.score);

  const suggestions = sorted
    .slice(0, topK)
    .map((s, idx) => ({
      ...s,
      confidence_pct: calculateConfidence(sortedScores, idx)
    }))
    .filter(s => s.confidence_pct >= minConfidence);

  const autoSelectCaseId = (suggestions[0]?.confidence_pct >= 70)
    ? suggestions[0].case_id
    : null;

  return { suggestions, autoSelectCaseId };
}
```

### Supporting Database Queries

#### Get Thread-Mapped Case
```sql
SELECT case_id, last_filed_at
FROM suggestion_history
WHERE user_id = $1
  AND conversation_id = $2
ORDER BY filed_at DESC
LIMIT 1;
```

#### Get Sender Filing History
```sql
SELECT case_id, COUNT(*) as filing_count, MAX(filed_at) as last_filed_at
FROM suggestion_history
WHERE user_id = $1
  AND sender_email = $2
  AND filed_at > NOW() - INTERVAL '90 days'
GROUP BY case_id
ORDER BY filing_count DESC, last_filed_at DESC;
```

#### Get Domain Filing History
```sql
SELECT case_id, COUNT(*) as filing_count
FROM suggestion_history
WHERE user_id = $1
  AND sender_domain = $2
  AND filed_at > NOW() - INTERVAL '90 days'
GROUP BY case_id
ORDER BY filing_count DESC;
```

#### Get Recent Cases
```sql
SELECT DISTINCT case_id, MAX(filed_at) as last_used_at
FROM suggestion_history
WHERE user_id = $1
  AND filed_at > NOW() - INTERVAL '14 days'
GROUP BY case_id
ORDER BY last_used_at DESC
LIMIT 10;
```

#### Record Successful Filing
```sql
INSERT INTO suggestion_history (
  id, user_id, case_id, sender_email, sender_domain,
  conversation_id, filed_at, source
) VALUES (
  $1, $2, $3, $4, $5, $6, NOW(), $7
);
```

### Text Normalization & Tokenization

```javascript
function normalize(text) {
  if (!text) return '';

  // Lowercase
  text = text.trim().toLowerCase();

  // Remove diacritics
  text = text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  // Strip Re:/Fw: prefixes for subjects
  text = text.replace(/^(re|fw|fwd):\s*/gi, '');

  // Collapse whitespace
  text = text.replace(/\s+/g, ' ');

  return text.trim();
}

function tokenize(text) {
  return normalize(text)
    .split(/\s+/)
    .filter(token => token.length >= 3); // Ignore short tokens
}

function calculateTokenOverlap(tokens1, text2) {
  const text2Normalized = ` ${normalize(text2)} `;
  let hits = 0;

  for (const token of tokens1) {
    if (text2Normalized.includes(` ${token} `)) {
      hits++;
    }
  }

  return { hits, total: tokens1.length };
}

function extractDomain(email) {
  const match = email.match(/@([^@]+)$/);
  return match ? match[1].toLowerCase() : null;
}
```

---

## Reference Implementation

The demo project at `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/` provides the client-side suggestion algorithm:

### Complete Suggestion Algorithm
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/caseSuggestionEngine.ts`:

- **Lines 101-290:** `suggestCasesLocal` function - THE REFERENCE ALGORITHM
  - Thread mapping (lines 139-142): 100 points
  - Case ID reference (lines 149-162): 95 points
  - Case title matching (lines 166-218): 60-98 points
  - Sender history (lines 221-226): 30 points with log scaling
  - Domain history (lines 229-234): 20 points with log scaling
  - Recent activity (lines 237-243): 12 points with decay

- **Lines 86-99:** `confidencePctFor` function
  - Confidence calculation formula (YOUR REFERENCE)
  - Base confidence + separation bonus
  - Weights: 65% base, 35% separation

- **Lines 40-84:** Text normalization utilities
  - `stripDiacritics` (lines 40-46)
  - `normText` and `normLoose` (lines 49-59)
  - `tokenizeLoose` (lines 61-66)
  - `tokenOverlapScore` (lines 68-76)

### Suggestion Storage & Tracking
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/suggestionStorage.ts`:

- **Lines 9-15:** `CaseSuggestState` structure
  - Shows what data to track for learning

- **Lines 144-192:** `recordSuccessfulAttach` function
  - Updates thread mapping
  - Increments sender history
  - Updates domain statistics
  - Tracks recent cases
  - **THIS IS CRITICAL FOR LEARNING**

- **Lines 194-202:** `getSuggestStats` function
  - Returns data for suggestion algorithm

### Content-Based Suggestions
See `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/caseSuggestionEngine.ts`:

- **Lines 293-422:** `suggestCasesByContent` function
  - Alternative algorithm without history
  - Pure content-based matching
  - Use case: "Select different case" button

---

## Dependencies

- **Requires:** BE-001 (Database Schema) - suggestion_history table
- **Requires:** BE-004 (Metadata Storage) - email metadata for matching
- **Related:** BE-002 (Filing API) - record filing in suggestion_history
- **Related:** BE-005 (Case Search) - get list of cases to analyze

---

## Notes

### Learning & Improvement

The engine learns from every successful filing:

1. **Record Filing Event:**
   ```javascript
   await recordFiling({
     userId: 'user-123',
     caseId: 'case-456',
     senderEmail: 'john@acme.com',
     senderDomain: 'acme.com',
     conversationId: 'AAQkADU...',
     source: 'manual' // or 'suggested' or 'auto'
   });
   ```

2. **Update Statistics:**
   - Increment sender → case mapping counter
   - Increment domain → case mapping counter
   - Update thread → case mapping
   - Add to recent cases list

3. **Prune Old Data:**
   - Keep last 90 days of sender history
   - Keep last 30 days of recent cases
   - Limit to 10 most common mappings per sender

### Algorithm Tuning

**Signal Priorities (from testing):**

1. **Thread Context (100 pts):** Most reliable - same thread should go to same case
2. **Case ID Reference (95 pts):** Explicit reference is strong signal
3. **Exact Subject Match (98 pts):** Very high confidence
4. **Sender History (10-50 pts):** Reliable but needs frequency threshold
5. **Domain History (20-40 pts):** Useful for organizations
6. **Recent Activity (12 pts):** Weak signal, but good tiebreaker

**Confidence Thresholds:**
- **70%+:** Auto-select (very confident)
- **30-69%:** Show as suggestion
- **10-29%:** Show if no better options
- **< 10%:** Don't show

### Testing Strategy

```javascript
// Test Scenarios

1. Thread Context:
   - Email in thread filed to Case A → suggest Case A with 100+ score
   - New thread → no thread bonus

2. Case Reference:
   - Subject "Re: Case 2024-0123" → suggest that case with 95+ score
   - Body mentions "2024-0123" → suggest that case

3. Name Matching:
   - Subject "Q4 Planning" + Case "Q4 Planning" → 98 score
   - Subject "Q4 Budget" + Case "Q4 Planning" → 60-90 score (token overlap)
   - Body mentions case name → 35-60 score

4. Sender History:
   - 15 emails from sender to Case A → 50 score
   - 5 emails from sender to Case A → 30 score
   - 1 email from sender to Case A → 10 score

5. Confidence Calculation:
   - Top score 150, second 50 → 87% confidence
   - Top score 80, second 75 → 42% confidence (close race)
   - Top score 30, second 5 → 33% confidence (low score)

6. Edge Cases:
   - New sender (no history) → content-based only
   - Empty subject → use sender/domain history
   - No matches → return empty suggestions []
```

### Performance Optimization

1. **Caching:**
   - Cache user's suggestion statistics (15 min TTL)
   - Cache recent cases (5 min TTL)
   - Cache case list (10 min TTL)

2. **Query Optimization:**
   - Limit history queries to 90 days
   - Use indexes on sender_email, sender_domain
   - Pre-aggregate statistics for frequent senders

3. **Async Processing:**
   - Update suggestion_history asynchronously after filing
   - Don't block filing API waiting for statistics update

### Monitoring & Observability

Track metrics:
- Suggestion acceptance rate (% of time user selects suggested case)
- Auto-select accuracy (% of time auto-selected case is kept)
- Average confidence score
- Signal distribution (which signals fire most often)
- Response time

Alert on:
- Suggestion acceptance rate < 60%
- P95 response time > 200ms
- Error rate > 2%

### Future Enhancements

- **Machine Learning:** Train model on user behavior
- **Collaborative Filtering:** "Users like you filed to Case X"
- **Time-Based Patterns:** "You usually file these emails on Mondays"
- **Project Phase Detection:** "Case is in closing phase, unlikely to receive new emails"
- **Sentiment Analysis:** Urgent emails → suggest different cases
- **Email Classification:** Automatically categorize email type (invoice, contract, correspondence)
