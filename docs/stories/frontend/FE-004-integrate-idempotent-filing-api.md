# FE-004: Integrate Idempotent Filing API

**Story ID:** FE-004
**Story Points:** 8
**Epic Link:** Email Filing Core Functionality
**Status:** Ready for Development

## Description

Build from scratch the email filing integration with the new idempotent backend API. When users file an email to a case, the frontend must collect all email metadata, attachments, and context, then send a single idempotent POST request to `/publicapi/v1/cases/{caseId}/emails`. The API will handle duplicate detection, attachment uploads, and email record creation atomically.

This replaces the current mock filing logic with a real, production-ready filing flow that supports retries, partial failures, and cross-mailbox duplicate prevention.

## Acceptance Criteria

1. **Email Metadata Collection**
   - Extract subject, from, to, cc, bcc from Office.js
   - Extract body (plain text and/or HTML) up to 100KB
   - Extract `internetMessageId` as idempotency key
   - Extract `conversationId` for threading
   - Extract received date/time
   - Extract sender display name

2. **Attachment Handling**
   - Get all attachments via `Office.context.mailbox.item.attachments`
   - For each attachment: get name, size, contentType, isInline
   - Download attachment content as Base64 via `item.getAttachmentContentAsync()`
   - Support attachments up to 25MB each (Office.js limit)
   - Show progress indicator while downloading attachments
   - Handle attachment download failures gracefully

3. **Idempotent Filing Request**
   - Send POST to `/publicapi/v1/cases/{caseId}/emails` with all data
   - Include `internetMessageId` for idempotency
   - Include attachments as Base64 in request body
   - Include user preference flags (e.g., includeBodySnippet, includeAttachments)
   - Retry on network failure (up to 3 attempts with exponential backoff)
   - Handle 409 Conflict (already filed) as success

4. **Response Handling**
   - 201 Created: Email filed successfully, show success message
   - 409 Conflict: Email already filed to this case, show "Already filed" message
   - 400 Bad Request: Show validation error to user
   - 401 Unauthorized: Clear auth and redirect to login
   - 413 Payload Too Large: Show error about attachment sizes
   - 500 Server Error: Show retry option

5. **Category Application**
   - After successful filing, apply "SC: Filed" category to email
   - Use Office.js categories API (`item.categories.addAsync()`)
   - Fallback to Graph API if Office.js fails
   - Handle category application failure without blocking success message

6. **Success Feedback**
   - Show success toast/banner: "Email filed to [Case Name]"
   - Update UI to reflect filed state (disable file button, show checkmark)
   - Optionally show link to view email in SingleCase web app
   - Close filing dialog after success

## Technical Requirements

### React Components

1. **Update MainWorkspace.tsx**
   - Integrate filing API call on "File Email" button click
   - Show loading spinner during filing
   - Handle all response scenarios
   - Show success/error messages
   - Apply category after successful filing

2. **FilingProgressModal.tsx** (new component)
   ```tsx
   interface FilingProgressModalProps {
     isOpen: boolean;
     step: 'collecting' | 'uploading_attachments' | 'filing' | 'applying_category' | 'done';
     progress: number; // 0-100
     error: Error | null;
     onClose: () => void;
   }
   ```
   - Show progress bar for each filing step
   - Display current step label
   - Show attachment upload progress (X of Y)
   - Handle errors with retry option

3. **Update AttachmentsStep.tsx**
   - Allow user to toggle which attachments to include
   - Show attachment size warnings (e.g., "Large attachment: 23MB")
   - Display total payload size estimate
   - Warn if total > 25MB (API limit)

### Services

1. **services/emailFiling.ts** (new service)
   ```typescript
   export interface EmailFilingPayload {
     internetMessageId: string;
     conversationId?: string;
     subject: string;
     fromEmail: string;
     fromName: string;
     toEmails: string[];
     ccEmails: string[];
     bccEmails: string[];
     bodyText?: string;
     bodyHtml?: string;
     receivedDateTime: string; // ISO 8601
     attachments: AttachmentPayload[];
   }

   export interface AttachmentPayload {
     name: string;
     contentType: string;
     size: number;
     contentBase64: string;
     isInline: boolean;
   }

   export interface FilingResult {
     success: boolean;
     emailId?: string;
     message: string;
     alreadyFiled?: boolean;
   }

   export async function fileEmailToCase(
     token: string,
     caseId: string,
     payload: EmailFilingPayload
   ): Promise<FilingResult>;

   export async function fileEmailToCaseWithRetry(
     token: string,
     caseId: string,
     payload: EmailFilingPayload,
     maxRetries?: number
   ): Promise<FilingResult>;
   ```

2. **services/attachmentDownload.ts** (new service)
   ```typescript
   export async function downloadAttachment(
     attachmentId: string
   ): Promise<{ contentBase64: string }>;

   export async function downloadAllAttachments(
     attachments: Office.AttachmentDetails[],
     onProgress?: (current: number, total: number) => void
   ): Promise<AttachmentPayload[]>;
   ```

3. **Update services/singlecase.ts**
   - Add `fileEmail` method
   - Handle large payload (chunked upload if needed)
   - Add request timeout (30s for small, 120s for large)

### Office.js APIs

1. **Email Metadata Extraction**
   ```typescript
   async function extractEmailMetadata(): Promise<Partial<EmailFilingPayload>> {
     const item = Office.context.mailbox.item;

     return {
       subject: item.subject || '',
       fromEmail: item.from?.emailAddress || '',
       fromName: item.from?.displayName || '',
       toEmails: item.to?.map(r => r.emailAddress) || [],
       ccEmails: item.cc?.map(r => r.emailAddress) || [],
       bccEmails: item.bcc?.map(r => r.emailAddress) || [],
       conversationId: item.conversationId || undefined,
       receivedDateTime: item.dateTimeCreated?.toISOString() || new Date().toISOString(),
     };
   }
   ```

2. **Body Extraction**
   ```typescript
   async function extractBody(): Promise<{ bodyText?: string; bodyHtml?: string }> {
     return new Promise((resolve) => {
       Office.context.mailbox.item.body.getAsync(
         Office.CoercionType.Html,
         (result) => {
           if (result.status === Office.AsyncResultStatus.Succeeded) {
             resolve({ bodyHtml: result.value });
           } else {
             // Fallback to text
             Office.context.mailbox.item.body.getAsync(
               Office.CoercionType.Text,
               (textResult) => {
                 resolve({
                   bodyText: textResult.status === Office.AsyncResultStatus.Succeeded
                     ? textResult.value
                     : undefined
                 });
               }
             );
           }
         }
       );
     });
   }
   ```

3. **Attachment Download**
   ```typescript
   async function downloadAttachment(attachmentId: string): Promise<string> {
     return new Promise((resolve, reject) => {
       Office.context.mailbox.item.getAttachmentContentAsync(
         attachmentId,
         (result) => {
           if (result.status === Office.AsyncResultStatus.Succeeded) {
             resolve(result.value.content); // Base64 string
           } else {
             reject(result.error);
           }
         }
       );
     });
   }
   ```

4. **Category Application**
   ```typescript
   async function applyFiledCategory(): Promise<void> {
     return new Promise((resolve, reject) => {
       Office.context.mailbox.item.categories.addAsync(
         ['SC: Filed'],
         (result) => {
           if (result.status === Office.AsyncResultStatus.Succeeded) {
             resolve();
           } else {
             // Fallback to Graph API
             applyFiledCategoryViaGraph().then(resolve).catch(reject);
           }
         }
       );
     });
   }
   ```

### API Integration Patterns

1. **Complete Filing Flow**
   ```typescript
   // In MainWorkspace.tsx
   const handleFileEmail = async () => {
     setFilingStep('collecting');
     setIsFilingModalOpen(true);

     try {
       // Step 1: Collect metadata
       const metadata = await extractEmailMetadata();
       const body = await extractBody();
       const internetMessageId = await getInternetMessageId();

       // Step 2: Download attachments
       setFilingStep('uploading_attachments');
       const attachmentPayloads = await downloadAllAttachments(
         Office.context.mailbox.item.attachments,
         (current, total) => setFilingProgress((current / total) * 50) // 0-50%
       );

       // Step 3: File email
       setFilingStep('filing');
       const result = await fileEmailToCaseWithRetry(
         token,
         selectedCaseId,
         {
           internetMessageId,
           ...metadata,
           ...body,
           attachments: includeAttachments ? attachmentPayloads : [],
         },
         3 // max retries
       );

       if (!result.success) {
         throw new Error(result.message);
       }

       // Step 4: Apply category
       setFilingStep('applying_category');
       await applyFiledCategory();

       // Step 5: Success
       setFilingStep('done');
       setFilingProgress(100);
       showSuccessToast(`Email filed to ${selectedCaseName}`);

     } catch (error) {
       setFilingError(error);
       showErrorToast('Failed to file email');
     } finally {
       setIsFilingModalOpen(false);
     }
   };
   ```

2. **Retry with Exponential Backoff**
   ```typescript
   export async function fileEmailToCaseWithRetry(
     token: string,
     caseId: string,
     payload: EmailFilingPayload,
     maxRetries: number = 3
   ): Promise<FilingResult> {
     let lastError: Error | null = null;

     for (let attempt = 0; attempt < maxRetries; attempt++) {
       try {
         return await fileEmailToCase(token, caseId, payload);
       } catch (error) {
         lastError = error as Error;

         // Don't retry on client errors (4xx except 429)
         if (error.status >= 400 && error.status < 500 && error.status !== 429) {
           throw error;
         }

         // Wait before retry (exponential backoff)
         const delayMs = Math.min(1000 * Math.pow(2, attempt), 10000);
         await new Promise(resolve => setTimeout(resolve, delayMs));
       }
     }

     throw lastError || new Error('Filing failed after retries');
   }
   ```

3. **Idempotency Handling**
   ```typescript
   // API returns 409 if already filed
   const response = await scRequest<any>(
     'POST',
     `/cases/${caseId}/emails`,
     token,
     payload
   );

   if (response.status === 409) {
     return {
       success: true,
       alreadyFiled: true,
       message: 'Email already filed to this case',
     };
   }
   ```

## Reference Implementation

Review the demo for patterns to adapt:

1. **Current Mock Filing**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecase.ts`
   - `submitEmailToCase()` function (currently mocked)
   - **REPLACE** with real API integration

2. **Attachment Handling**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/AttachmentsPicker.tsx`
   - Shows UI for selecting attachments
   - Reuse UI patterns, add download logic

3. **Category Application**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/graphMail.ts`
   - `applyFiledCategoryToCurrentEmailOfficeJs()` function
   - Reuse this logic after successful filing

4. **Email Metadata**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/MainWorkspace/MainWorkspace.tsx`
   - Shows how email data is extracted
   - Expand to collect all required fields

## Dependencies

- **Requires**: FE-001 (OAuth Integration) - need token for authenticated API calls
- **Requires**: FE-002 (Duplicate Detection) - should check before filing
- **Requires**: Backend API endpoint `POST /publicapi/v1/cases/{caseId}/emails` must be implemented
- **Relates to**: FE-005 (Error Handling) - error scenarios must be handled gracefully

## Notes

1. **Attachment Size Limits**
   - Office.js: 25MB per attachment
   - SingleCase API: Configure server-side limit (recommend 50MB total)
   - Warn users before uploading large attachments
   - Consider chunked upload for files > 10MB (future enhancement)

2. **Idempotency Key**
   - `internetMessageId` is unique per email across all mailboxes
   - Use this as idempotency key to prevent duplicate filings
   - If unavailable (compose mode), generate temporary key from `conversationId` + `subject` hash
   - Backend should deduplicate based on `internetMessageId`

3. **Partial Failures**
   - Email filed but category application failed: Still show success, log warning
   - Attachments partially uploaded: Retry entire request (idempotent)
   - Network timeout during filing: Retry with same `internetMessageId`

4. **Performance**
   - Attachment download is slow: Show progress bar
   - Large emails (>1MB): Show estimated upload time
   - Consider compression for body content (future enhancement)

5. **Testing Strategy**
   - Mock Office.js APIs for attachment download
   - Test with various attachment types (PDF, images, Office docs)
   - Test with inline attachments (should be filtered out)
   - Test retry logic with simulated network failures
   - Test idempotency (file same email twice)
   - Test 409 Conflict handling

6. **User Settings Integration**
   - `includeBodySnippet`: Control whether to send full body or snippet
   - `includeAttachments`: Control whether to upload attachments
   - `rememberLastCase`: Auto-select last used case
   - Load settings from `settingsStorage.ts`

7. **Offline Handling**
   - If offline, queue filing request for later (future: FE-010)
   - Show "Offline - will file when connected" message
   - For now, just show error "No internet connection"

## Definition of Done

- [ ] Email metadata collected from Office.js (subject, from, to, cc, body)
- [ ] Attachments downloaded as Base64 via Office.js
- [ ] Idempotent POST request to `/cases/{caseId}/emails` working
- [ ] `internetMessageId` used as idempotency key
- [ ] 409 Conflict (already filed) handled as success
- [ ] Retry logic with exponential backoff (up to 3 attempts)
- [ ] "SC: Filed" category applied after successful filing
- [ ] Progress modal shows filing steps and progress
- [ ] Success toast displayed after filing
- [ ] Error handling for all failure scenarios (4xx, 5xx, network)
- [ ] Large attachment warnings (>10MB)
- [ ] Unit tests for filing service
- [ ] Integration tests for full filing flow
- [ ] Performance tested with large attachments (20MB+)
- [ ] Documentation updated with filing API spec
