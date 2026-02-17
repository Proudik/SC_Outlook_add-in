# FE-003: Integrate Server-Side Case Suggestions

**Story ID:** FE-003
**Story Points:** 5
**Epic Link:** Smart Filing & Suggestions
**Status:** Ready for Development

## Description

Replace the current client-side case suggestion logic with server-side intelligent case suggestions powered by machine learning and historical filing patterns. The backend API will analyze email content (subject, body, sender, recipients) and return ranked case suggestions. The frontend must integrate this API, display suggestions with confidence scores, and allow users to accept or ignore suggestions.

Remove all local suggestion algorithms and replace with API-driven suggestions that improve over time as more emails are filed.

## Acceptance Criteria

1. **API Integration**
   - Call `POST /publicapi/v1/cases/suggest` with email metadata
   - Send: `{ subject, bodySnippet, fromEmail, toEmails, ccEmails, attachmentNames }`
   - Receive: `{ suggestions: Array<{ caseId, caseName, confidenceScore, reason }> }`
   - Send authenticated request with OAuth token
   - Handle pagination if backend returns many suggestions

2. **Suggestion Display**
   - Show top 3 suggestions in filing dialog
   - Display case name prominently
   - Show confidence score as percentage or visual indicator (stars/dots)
   - Show reason for suggestion (e.g., "Similar subject", "Same sender", "Recent history")
   - Highlight suggested case in case dropdown/picker
   - Allow user to click suggestion to auto-select that case

3. **Auto-Selection Logic**
   - If top suggestion has confidence > 80%, auto-select it
   - Show visual indicator that case was auto-selected
   - User can override auto-selection by manually picking different case
   - Don't auto-select if user has already manually selected a case
   - Remember user's manual overrides to improve future suggestions

4. **Loading & Empty States**
   - Show "Analyzing email..." spinner while fetching suggestions
   - If no suggestions returned, show "No suggestions available"
   - If suggestions API fails, fall back to showing full case list without suggestions
   - Never block filing if suggestions fail to load

5. **Remove Legacy Logic**
   - Delete `utils/caseSuggestionEngine.ts` (local suggestion logic)
   - Remove `hooks/useCaseSuggestions.ts` (old local suggestion hook)
   - Remove `utils/suggestionStorage.ts` (local suggestion cache)
   - Clean up any imports referencing deleted files

## Technical Requirements

### React Components

1. **CaseSuggestionsPanel.tsx** (new component)
   ```tsx
   interface CaseSuggestion {
     caseId: string;
     caseName: string;
     confidenceScore: number; // 0-100
     reason: string;
   }

   interface CaseSuggestionsPanelProps {
     suggestions: CaseSuggestion[];
     isLoading: boolean;
     onSelectSuggestion: (caseId: string) => void;
     selectedCaseId?: string;
   }
   ```
   - Display suggestions in card layout
   - Show confidence indicator (progress bar or star rating)
   - Show reason text below case name
   - Highlight selected suggestion
   - Click to select suggestion

2. **SuggestionConfidenceIndicator.tsx** (new component)
   ```tsx
   interface SuggestionConfidenceIndicatorProps {
     confidenceScore: number; // 0-100
     variant: 'stars' | 'bar' | 'percentage';
   }
   ```
   - Visual representation of confidence
   - Stars: 1-5 stars based on score
   - Bar: Progress bar with color gradient (red → yellow → green)
   - Percentage: "85% confident"

3. **Update MainWorkspace.tsx**
   - Fetch suggestions when email opens
   - Pass suggestions to filing dialog
   - Handle auto-selection if confidence high
   - Track if user manually overrode suggestion

4. **Update CaseSelector.tsx**
   - Integrate suggestions panel above case dropdown
   - Highlight suggested cases in dropdown with badge
   - Handle suggestion click to populate selector

### Services

1. **services/caseSuggestions.ts** (new service)
   ```typescript
   export interface EmailContext {
     subject: string;
     bodySnippet: string; // First 500 chars
     fromEmail: string;
     toEmails: string[];
     ccEmails: string[];
     attachmentNames: string[];
   }

   export interface CaseSuggestion {
     caseId: string;
     caseName: string;
     confidenceScore: number; // 0-100
     reason: string;
   }

   export async function getSuggestedCases(
     token: string,
     emailContext: EmailContext
   ): Promise<CaseSuggestion[]>;

   export async function recordSuggestionFeedback(
     token: string,
     suggestionId: string,
     accepted: boolean
   ): Promise<void>;
   ```

2. **Update services/singlecase.ts**
   - Add suggestion API methods
   - Handle suggestion response parsing
   - Normalize confidence scores (ensure 0-100 range)

### Hooks

1. **hooks/useCaseSuggestionsApi.ts** (new hook, replaces old useCaseSuggestions)
   ```typescript
   export function useCaseSuggestionsApi(
     token: string,
     emailContext: EmailContext | null,
     enabled: boolean
   ): {
     suggestions: CaseSuggestion[];
     isLoading: boolean;
     error: Error | null;
     refetch: () => Promise<void>;
   }
   ```
   - Fetch suggestions on mount if enabled
   - Cache suggestions for current email (by subject + from)
   - Debounce refetch (500ms) if email context changes
   - Auto-retry once on failure

2. **hooks/useAutoSelectSuggestion.ts** (new hook)
   ```typescript
   export function useAutoSelectSuggestion(
     suggestions: CaseSuggestion[],
     selectedCaseId: string,
     userManuallySelected: boolean,
     threshold: number = 80
   ): {
     autoSelectedCaseId: string | null;
     shouldAutoSelect: boolean;
   }
   ```
   - Auto-select top suggestion if confidence > threshold
   - Don't override user's manual selection
   - Return null if no high-confidence suggestion

### Office.js APIs

1. **Email Metadata Extraction**
   - Use `Office.context.mailbox.item.subject` for subject
   - Use `Office.context.mailbox.item.body.getAsync()` for body snippet (first 500 chars)
   - Use `Office.context.mailbox.item.from` for sender email
   - Use `Office.context.mailbox.item.to` / `cc` for recipients
   - Use `Office.context.mailbox.item.attachments` for attachment names

2. **Body Snippet Extraction**
   ```typescript
   async function getBodySnippet(): Promise<string> {
     return new Promise((resolve, reject) => {
       Office.context.mailbox.item.body.getAsync(
         Office.CoercionType.Text,
         { asyncContext: {} },
         (result) => {
           if (result.status === Office.AsyncResultStatus.Succeeded) {
             // Take first 500 chars only
             resolve(result.value.substring(0, 500));
           } else {
             reject(result.error);
           }
         }
       );
     });
   }
   ```

### API Integration Patterns

1. **Fetch Suggestions**
   ```typescript
   // In services/caseSuggestions.ts
   export async function getSuggestedCases(
     token: string,
     emailContext: EmailContext
   ): Promise<CaseSuggestion[]> {
     const response = await scRequest<any>(
       'POST',
       '/cases/suggest',
       token,
       {
         subject: emailContext.subject,
         body_snippet: emailContext.bodySnippet,
         from_email: emailContext.fromEmail,
         to_emails: emailContext.toEmails,
         cc_emails: emailContext.ccEmails,
         attachment_names: emailContext.attachmentNames,
       }
     );

     const suggestions = Array.isArray(response.suggestions)
       ? response.suggestions
       : [];

     return suggestions.map((s: any) => ({
       caseId: String(s.case_id || s.caseId),
       caseName: String(s.case_name || s.caseName || 'Unknown'),
       confidenceScore: Math.min(100, Math.max(0, Number(s.confidence_score || s.confidenceScore || 0))),
       reason: String(s.reason || 'Similar to previous emails'),
     }));
   }
   ```

2. **Auto-Selection with Override Tracking**
   ```typescript
   // In MainWorkspace.tsx
   const [userManuallySelected, setUserManuallySelected] = React.useState(false);

   const { autoSelectedCaseId } = useAutoSelectSuggestion(
     suggestions,
     selectedCaseId,
     userManuallySelected,
     80 // confidence threshold
   );

   React.useEffect(() => {
     if (autoSelectedCaseId && !userManuallySelected) {
       setSelectedCaseId(autoSelectedCaseId);
     }
   }, [autoSelectedCaseId, userManuallySelected]);

   const handleManualCaseSelection = (caseId: string) => {
     setSelectedCaseId(caseId);
     setUserManuallySelected(true);
   };
   ```

3. **Suggestion Feedback**
   ```typescript
   // After user files email
   const selectedSuggestion = suggestions.find(s => s.caseId === selectedCaseId);
   if (selectedSuggestion) {
     await recordSuggestionFeedback(
       token,
       selectedSuggestion.caseId,
       true // accepted
     );
   }
   ```

## Reference Implementation

Review the demo's local suggestion logic (then remove/replace):

1. **Current Suggestion Logic**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/caseSuggestionEngine.ts`
   - Shows local algorithm (keyword matching, sender history, etc.)
   - **DELETE THIS FILE** after implementing server-side suggestions

2. **Current Suggestion Hook**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/hooks/useCaseSuggestions.ts`
   - Shows how suggestions are currently triggered and used
   - Reuse hook pattern, but call API instead of local function
   - **REPLACE COMPLETELY** with new API-driven hook

3. **Email Metadata Extraction**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/MainWorkspace/MainWorkspace.tsx`
   - Shows how email data is extracted from Office.js
   - Reuse this extraction logic for API payload

4. **UI Integration**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/CaseSelector.tsx`
   - Current case picker UI
   - Add suggestions panel above this component

## Dependencies

- **Requires**: FE-001 (OAuth Integration) - need token for authenticated API calls
- **Requires**: Backend API endpoint `POST /publicapi/v1/cases/suggest` must be implemented
- **Relates to**: FE-002 (Duplicate Detection) - similar API integration pattern

## Notes

1. **Performance Considerations**
   - Cache suggestions per email (by subject + sender) for 5 minutes
   - Debounce suggestion fetching if email context changes rapidly
   - Show cached suggestions immediately, refresh in background
   - Cancel in-flight requests if user closes filing dialog

2. **Machine Learning Feedback Loop**
   - Record which suggestions users accept vs. ignore
   - Send feedback to backend to improve ML model
   - Track: suggested case ID, selected case ID, accepted (boolean)
   - Feedback API: `POST /publicapi/v1/cases/suggest/feedback`

3. **Edge Cases**
   - No suggestions returned: Show full case list without highlighting
   - API returns empty array: Treat as "no suggestions"
   - Multiple suggestions with same confidence: Show all, no auto-select
   - Suggestion API timeout: Fail gracefully, show case list

4. **User Experience**
   - Don't surprise users with auto-selection - show visual indicator
   - Make it obvious when a case is suggested vs. manually selected
   - Allow easy dismissal of suggestions
   - Show "Why this suggestion?" tooltip on hover

5. **Testing Strategy**
   - Mock suggestion API with various confidence scores
   - Test auto-selection threshold (80%, 90%, 100%)
   - Test user override behavior (manual selection after auto-select)
   - Test suggestion ranking (highest confidence first)
   - Test empty/error states

6. **Confidence Score Display**
   - 90-100%: Green, "Very confident"
   - 70-89%: Yellow, "Confident"
   - 50-69%: Orange, "Possible match"
   - < 50%: Don't show or show with "Low confidence" warning

7. **Localization**
   - Translate suggestion reasons (e.g., "Similar subject")
   - Keep case names as-is (no translation)
   - Localize confidence labels ("Very confident", etc.)

## Definition of Done

- [ ] Server-side suggestion API integrated
- [ ] `CaseSuggestionsPanel` component displays top 3 suggestions
- [ ] Confidence scores displayed visually (stars or progress bar)
- [ ] Suggestion reasons displayed below case names
- [ ] Auto-selection works for high-confidence suggestions (>80%)
- [ ] User can override auto-selected suggestion
- [ ] Loading state shown while fetching suggestions
- [ ] Error handling falls back to full case list
- [ ] Legacy local suggestion logic completely removed
- [ ] `utils/caseSuggestionEngine.ts` deleted
- [ ] `hooks/useCaseSuggestions.ts` replaced with new API hook
- [ ] Suggestion feedback recorded after filing
- [ ] Unit tests for suggestion service
- [ ] Integration tests for suggestion UI
- [ ] Performance tested with slow API responses
- [ ] Documentation updated with suggestion behavior
