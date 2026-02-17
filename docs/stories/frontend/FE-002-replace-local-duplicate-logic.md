# FE-002: Replace Local Duplicate Logic with Server-Side Duplicate Detection

**Story ID:** FE-002
**Story Points:** 5
**Epic Link:** Email Filing Core Functionality
**Status:** Ready for Development

## Description

Remove the current client-side duplicate detection logic and replace it with a server-side duplicate check API. The new implementation must call the backend API to check if an email (identified by `internetMessageId`) has already been filed to any case. Display server-provided duplicate information to users before they file, preventing accidental duplicate filings.

This story involves removing local caching logic and integrating a new API endpoint that provides authoritative duplicate detection across all users and mailboxes.

## Acceptance Criteria

1. **API Integration**
   - Call `POST /publicapi/v1/emails/check-duplicate` with `{ internetMessageId: string }`
   - Send authenticated request with OAuth token from FE-001
   - Handle response containing: `{ isDuplicate: boolean, caseId?: string, caseName?: string, filedBy?: string, filedAt?: string }`
   - Parse and validate API response before using data

2. **Duplicate Check Timing**
   - Check for duplicates when user opens filing dialog
   - Check when user switches between emails in conversation view
   - Do NOT check on every keystroke or rapid UI interaction
   - Debounce duplicate checks to avoid excessive API calls (500ms)

3. **User Experience**
   - Show loading spinner while checking for duplicates
   - Display prominent warning banner if duplicate detected
   - Banner includes: case name, filed date, filed by user
   - Allow user to file anyway (override) with confirmation dialog
   - Show success state ("No duplicates found") if email is new
   - Handle "already filed to THIS case" scenario gracefully (not an error)

4. **Error Handling**
   - If duplicate check API fails, allow filing to proceed (fail open)
   - Show warning: "Could not verify duplicates - proceed with caution"
   - Log API errors for debugging
   - Retry once on network timeout (5s timeout per attempt)
   - Never block filing due to duplicate check failures

5. **Remove Legacy Logic**
   - Delete local duplicate cache in `utils/filedCache.ts` references
   - Remove `internetMessageId` storage in local memory
   - Remove client-side duplicate tracking hooks
   - Clean up any localStorage or sessionStorage duplicate keys

## Technical Requirements

### React Components

1. **DuplicateWarningBanner.tsx** (new component)
   ```tsx
   interface DuplicateWarningProps {
     caseId: string;
     caseName: string;
     filedBy?: string;
     filedAt?: string;
     onFileAnyway: () => void;
     onCancel: () => void;
   }
   ```
   - Fluent UI `MessageBar` with warning severity
   - Show case info prominently
   - "File Anyway" and "Cancel" action buttons
   - Display timestamp in user's locale format

2. **Update MainWorkspace.tsx**
   - Integrate duplicate check before showing file button
   - Show `DuplicateWarningBanner` if duplicate detected
   - Handle "File Anyway" override flow
   - Clear duplicate state when switching emails

3. **Update AttachmentsStep.tsx**
   - Remove any duplicate checking logic
   - Rely on parent component (MainWorkspace) for duplicate status
   - Display duplicate warning if passed as prop

### Services

1. **services/duplicateCheck.ts** (new service)
   ```typescript
   export interface DuplicateCheckResult {
     isDuplicate: boolean;
     caseId?: string;
     caseName?: string;
     filedBy?: string;
     filedAt?: string;
   }

   export async function checkEmailDuplicate(
     token: string,
     internetMessageId: string
   ): Promise<DuplicateCheckResult>;

   export async function checkEmailDuplicateWithRetry(
     token: string,
     internetMessageId: string,
     maxRetries?: number
   ): Promise<DuplicateCheckResult | null>;
   ```

2. **Update services/singlecase.ts**
   - Add `checkDuplicate` method
   - Use authenticated `scRequest` helper
   - Handle 404 (email not found) as "not duplicate"
   - Handle 200 with duplicate data

### Hooks

1. **hooks/useDuplicateCheck.ts** (new hook)
   ```typescript
   export function useDuplicateCheck(
     token: string,
     internetMessageId: string,
     enabled: boolean
   ): {
     isChecking: boolean;
     isDuplicate: boolean;
     duplicateInfo: DuplicateCheckResult | null;
     error: Error | null;
     refetch: () => Promise<void>;
   }
   ```
   - Automatically check on mount if enabled
   - Debounce refetch calls (500ms)
   - Cache result for same `internetMessageId`
   - Clear cache when `internetMessageId` changes

### Office.js APIs

1. **Get internetMessageId**
   - Use `Office.context.mailbox.item.internetMessageId` if available (read mode)
   - For compose mode, use Graph API after send to retrieve it
   - Fallback: Use `conversationId` + `subject` hashing for temporary ID (pre-send)

2. **Error Handling**
   - If `internetMessageId` is unavailable, skip duplicate check
   - Show info message: "Duplicate check unavailable in compose mode"
   - Still allow filing to proceed

### API Integration Patterns

1. **Duplicate Check Request**
   ```typescript
   // In services/duplicateCheck.ts
   export async function checkEmailDuplicate(
     token: string,
     internetMessageId: string
   ): Promise<DuplicateCheckResult> {
     const response = await scRequest<any>(
       'POST',
       '/emails/check-duplicate',
       token,
       { internetMessageId }
     );

     return {
       isDuplicate: Boolean(response.isDuplicate),
       caseId: response.caseId ? String(response.caseId) : undefined,
       caseName: response.caseName ? String(response.caseName) : undefined,
       filedBy: response.filedBy ? String(response.filedBy) : undefined,
       filedAt: response.filedAt ? String(response.filedAt) : undefined,
     };
   }
   ```

2. **Debounced Check**
   ```typescript
   // In hooks/useDuplicateCheck.ts
   const checkDebounced = React.useMemo(
     () => debounce(async () => {
       setIsChecking(true);
       try {
         const result = await checkEmailDuplicateWithRetry(token, internetMessageId, 2);
         setDuplicateInfo(result);
         setIsDuplicate(result?.isDuplicate ?? false);
       } catch (err) {
         setError(err as Error);
       } finally {
         setIsChecking(false);
       }
     }, 500),
     [token, internetMessageId]
   );

   React.useEffect(() => {
     if (enabled && internetMessageId) {
       checkDebounced();
     }
   }, [enabled, internetMessageId, checkDebounced]);
   ```

3. **Fail-Open Error Handling**
   ```typescript
   // In MainWorkspace.tsx
   const { isDuplicate, duplicateInfo, error } = useDuplicateCheck(
     token,
     internetMessageId,
     true
   );

   if (error) {
     // Show warning but allow filing
     return (
       <MessageBar severity="warning">
         Could not verify duplicates. Proceed with caution.
       </MessageBar>
     );
   }
   ```

## Reference Implementation

Review the demo's duplicate detection for patterns (then remove/replace):

1. **Current Duplicate Logic**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/filedCache.ts`
   - Shows how duplicates are currently tracked locally
   - **DELETE THIS FILE** after implementing server-side checks

2. **internetMessageId Handling**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/graphMail.ts`
   - `getInternetMessageIdViaGraph()` function
   - Shows how to retrieve stable message identifier
   - Reuse this pattern for new duplicate check

3. **UI Patterns**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/MainWorkspace/MainWorkspace.tsx`
   - Current filing UI flow
   - Where to integrate duplicate warning banner

**Important**: Remove all local duplicate tracking. The server is now the source of truth for duplicate detection.

## Dependencies

- **Requires**: FE-001 (OAuth Integration) - need token for authenticated API calls
- **Requires**: Backend API endpoint `POST /publicapi/v1/emails/check-duplicate` must be implemented
- **Blocks**: FE-004 (Idempotent Filing API) - duplicate detection is prerequisite

## Notes

1. **Performance Considerations**
   - Cache duplicate check results per `internetMessageId` for 5 minutes
   - Debounce checks to avoid API spam
   - Show cached result immediately, refresh in background
   - Cancel in-flight requests if user switches emails

2. **Edge Cases**
   - Email filed to multiple cases: Show all cases in warning
   - Email filed to same case user is selecting: Show info, not warning
   - internetMessageId changes after send: Recheck duplicate status
   - Duplicate check API returns 500: Fail open, allow filing

3. **User Experience**
   - Don't block filing workflow with duplicate checks
   - Make duplicate warning prominent but not obtrusive
   - Allow power users to file duplicates intentionally (with confirmation)
   - Show filing history in duplicate warning (who filed, when)

4. **Testing Strategy**
   - Mock API responses for duplicate and non-duplicate scenarios
   - Test debounce behavior (rapid email switching)
   - Test retry logic on network failures
   - Test fail-open behavior when API is down
   - Test caching prevents redundant API calls

5. **Migration Path**
   - Deploy server-side API first
   - Update frontend to use new API
   - Verify duplicate detection working correctly
   - Remove old local cache logic
   - Clean up localStorage keys from user devices (one-time migration)

6. **Localization**
   - Duplicate warning messages must be localizable
   - Date/time formatting must respect user's locale
   - "Filed by" user names should be displayed as-is (no translation)

## Definition of Done

- [ ] Server-side duplicate check API integrated
- [ ] `DuplicateWarningBanner` component displays duplicate info
- [ ] Duplicate check debounced to avoid excessive API calls
- [ ] Loading states shown during duplicate check
- [ ] Error handling allows filing if duplicate check fails
- [ ] Legacy local duplicate cache logic completely removed
- [ ] `utils/filedCache.ts` deleted or gutted
- [ ] "File Anyway" override flow working with confirmation
- [ ] Unit tests for duplicate check service
- [ ] Integration tests for duplicate warning UI
- [ ] Performance tested with rapid email switching
- [ ] Documentation updated with duplicate check behavior
