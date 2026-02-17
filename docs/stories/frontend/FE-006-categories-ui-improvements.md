# FE-006: Categories UI Improvements

**Story ID:** FE-006
**Story Points:** 3
**Epic Link:** User Experience & Polish
**Status:** Ready for Development

## Description

Enhance the email category management UI to make filed/unfiled status more visible and actionable. Add visual indicators for category states, provide quick actions to manage categories, and ensure consistent category application across different Outlook platforms (Windows, Mac, Web, Mobile). Improve error handling for category operations and add fallback mechanisms when Office.js category APIs fail.

This story focuses on polishing the category experience that supports the filing workflow (FE-004).

## Acceptance Criteria

1. **Category Status Display**
   - Show badge/pill indicating current email's category status
   - "Filed" badge: Green, shows when "SC: Filed" category applied
   - "Unfiled" badge: Orange, shows when "SC: Unfiled" category applied
   - "No Category" state: Gray, shows when neither category present
   - Badge updates in real-time when category changes

2. **Category Quick Actions**
   - "Remove Category" button appears when email has a category
   - Confirmation dialog before removing "SC: Filed" category
   - Show category change animation/feedback
   - Disable category actions during operations

3. **Master Category Management**
   - Automatically create master categories on first use
   - "SC: Filed" category: Green (Preset4)
   - "SC: Unfiled" category: Orange (Preset7)
   - Verify master categories exist before applying
   - Handle category creation failures gracefully

4. **Cross-Platform Consistency**
   - Use Office.js categories API as primary method
   - Fallback to Graph API when Office.js unavailable (Mac/Web)
   - Verify category was applied successfully
   - Handle platform-specific category limitations

5. **Error Handling**
   - Show clear error if category application fails
   - Retry once on transient failures
   - Don't block filing workflow if category fails (log warning)
   - Provide manual "Apply Category" button if auto-apply fails

6. **Category Sync**
   - Read current categories when email opens
   - Update UI if categories change externally (other devices)
   - Handle race conditions (multiple devices applying categories)

## Technical Requirements

### React Components

1. **CategoryBadge.tsx** (new component)
   ```tsx
   interface CategoryBadgeProps {
     status: 'filed' | 'unfiled' | 'none';
     onClick?: () => void;
   }
   ```
   - Display pill-shaped badge with appropriate color
   - Show category icon (checkmark for filed, warning for unfiled)
   - Click to open category actions menu (optional)
   - Fluent UI `Badge` or custom styled component

2. **CategoryActionsMenu.tsx** (new component)
   ```tsx
   interface CategoryActionsMenuProps {
     currentStatus: 'filed' | 'unfiled' | 'none';
     onApplyFiled: () => Promise<void>;
     onApplyUnfiled: () => Promise<void>;
     onRemoveCategory: () => Promise<void>;
   }
   ```
   - Dropdown menu with category actions
   - "Mark as Filed" option
   - "Mark as Unfiled" option
   - "Remove Category" option
   - Disable actions during async operations

3. **RemoveCategoryConfirmDialog.tsx** (new component)
   ```tsx
   interface RemoveCategoryConfirmDialogProps {
     isOpen: boolean;
     categoryName: string;
     onConfirm: () => void;
     onCancel: () => void;
   }
   ```
   - Modal dialog asking for confirmation
   - Warn: "This email will no longer be marked as filed"
   - "Remove" and "Cancel" buttons

4. **Update MainWorkspace.tsx**
   - Display `CategoryBadge` near email subject
   - Read categories on email load
   - Update badge when filing completes
   - Show category actions menu

5. **Update Header.tsx** (if exists)
   - Integrate `CategoryBadge` into header area
   - Position near case selector or email info

### Services

1. **Update services/graphMail.ts**
   - Add `getCurrentEmailCategories()` function (Office.js)
   - Add `removeCategoryFromCurrentEmail()` function
   - Add `ensureMasterCategoriesExist()` function
   - Improve error handling for category operations

2. **services/categoryManager.ts** (new service)
   ```typescript
   export type CategoryStatus = 'filed' | 'unfiled' | 'none';

   export async function getCurrentCategoryStatus(): Promise<CategoryStatus>;

   export async function applyCategoryWithFallback(
     category: 'filed' | 'unfiled'
   ): Promise<{ success: boolean; method: 'office-js' | 'graph' | 'failed' }>;

   export async function removeCategoryWithConfirmation(
     categoryName: string
   ): Promise<boolean>;

   export async function ensureMasterCategories(): Promise<void>;

   export async function verifyCategoryApplied(
     categoryName: string
   ): Promise<boolean>;
   ```

### Hooks

1. **hooks/useEmailCategory.ts** (new hook)
   ```typescript
   export function useEmailCategory(): {
     status: CategoryStatus;
     isLoading: boolean;
     error: Error | null;
     applyFiled: () => Promise<void>;
     applyUnfiled: () => Promise<void>;
     removeCategory: () => Promise<void>;
     refresh: () => Promise<void>;
   }
   ```
   - Read current category on mount
   - Provide actions to change category
   - Auto-refresh after changes
   - Handle loading and error states

2. **hooks/useCategorySync.ts** (new hook)
   ```typescript
   export function useCategorySync(
     enabled: boolean,
     intervalMs?: number
   ): {
     status: CategoryStatus;
     lastSyncTime: Date | null;
   }
   ```
   - Poll for category changes every 5 seconds (optional)
   - Compare with last known status
   - Trigger UI update if changed

### Office.js APIs

1. **Read Current Categories (Office.js)**
   ```typescript
   async function getCurrentCategories(): Promise<string[]> {
     return new Promise((resolve, reject) => {
       const item = Office.context.mailbox.item;
       if (!item?.categories) {
         return resolve([]);
       }

       item.categories.getAsync((result) => {
         if (result.status === Office.AsyncResultStatus.Succeeded) {
           resolve(result.value || []);
         } else {
           reject(result.error);
         }
       });
     });
   }
   ```

2. **Remove Category (Office.js)**
   ```typescript
   async function removeCategory(categoryName: string): Promise<void> {
     return new Promise((resolve, reject) => {
       const item = Office.context.mailbox.item;
       if (!item?.categories) {
         return reject(new Error('Categories API unavailable'));
       }

       item.categories.removeAsync([categoryName], (result) => {
         if (result.status === Office.AsyncResultStatus.Succeeded) {
           resolve();
         } else {
           reject(result.error);
         }
       });
     });
   }
   ```

3. **Ensure Master Category Exists**
   ```typescript
   async function ensureMasterCategory(
     displayName: string,
     color: Office.MailboxEnums.CategoryColor
   ): Promise<void> {
     return new Promise((resolve) => {
       try {
         const mc = Office.context.mailbox.masterCategories;
         if (!mc?.addAsync) return resolve(); // Not supported on this platform

         mc.addAsync([{ displayName, color }], () => {
           resolve(); // Ignore errors (category may already exist)
         });
       } catch {
         resolve();
       }
     });
   }
   ```

4. **Fallback to Graph API**
   - Use pattern from `services/graphMail.ts`
   - If Office.js fails, try Graph API
   - Get access token via `OfficeRuntime.auth.getAccessToken()`
   - PATCH `/me/messages/{messageId}` with categories array

## API Integration Patterns

1. **Apply Category with Fallback**
   ```typescript
   // In services/categoryManager.ts
   export async function applyCategoryWithFallback(
     category: 'filed' | 'unfiled'
   ): Promise<{ success: boolean; method: string }> {
     const categoryName = category === 'filed' ? 'SC: Filed' : 'SC: Unfiled';
     const color = category === 'filed'
       ? Office.MailboxEnums.CategoryColor.Preset4
       : Office.MailboxEnums.CategoryColor.Preset7;

     // Ensure master category exists
     await ensureMasterCategory(categoryName, color);

     // Try Office.js first
     try {
       await applyFiledCategoryToCurrentEmailOfficeJs();
       return { success: true, method: 'office-js' };
     } catch (officeJsError) {
       console.warn('Office.js categories failed, trying Graph API', officeJsError);

       // Fallback to Graph API
       try {
         const token = await getGraphAccessToken();
         await applyFiledCategoryToCurrentEmail(token);
         return { success: true, method: 'graph' };
       } catch (graphError) {
         console.error('Both Office.js and Graph API failed', graphError);
         return { success: false, method: 'failed' };
       }
     }
   }
   ```

2. **Verify Category Applied**
   ```typescript
   export async function verifyCategoryApplied(
     categoryName: string
   ): Promise<boolean> {
     // Wait 500ms for category to propagate
     await new Promise(resolve => setTimeout(resolve, 500));

     const categories = await getCurrentCategories();
     return categories.includes(categoryName);
   }
   ```

3. **Remove Category with Confirmation**
   ```typescript
   // In MainWorkspace.tsx
   const handleRemoveCategory = async () => {
     const confirmed = await showConfirmDialog(
       'Remove filing category?',
       'This email will no longer be marked as filed in Outlook.'
     );

     if (!confirmed) return;

     try {
       await removeCategory('SC: Filed');
       showSuccess('Category removed');
       await refreshCategoryStatus();
     } catch (error) {
       showError('Failed to remove category');
     }
   };
   ```

## Reference Implementation

Review the demo's category management:

1. **Current Category Logic**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/graphMail.ts`
   - `applyFiledCategoryToCurrentEmailOfficeJs()` function
   - `getCurrentEmailCategoriesGraph()` function
   - **ENHANCE** with better error handling and UI integration

2. **Master Category Creation**: Same file
   - Pattern for creating master categories
   - Reuse and improve this logic

3. **UI Integration Points**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/MainWorkspace/MainWorkspace.tsx`
   - Where to add `CategoryBadge` component
   - Where to trigger category application

## Dependencies

- **Requires**: FE-004 (Idempotent Filing) - categories applied after filing
- **Requires**: FE-005 (Error Handling) - use toast system for category errors
- **Enhances**: User experience for filed email tracking

## Notes

1. **Platform Differences**
   - Windows: Office.js categories work well
   - Mac: Office.js categories may be limited, use Graph fallback
   - Web: Office.js categories supported, Graph fallback recommended
   - Mobile: Limited category support, gracefully degrade

2. **Category Colors**
   - Filed (Green): Preset4 in Office.js, "preset4" in Graph API
   - Unfiled (Orange): Preset7 in Office.js, "preset7" in Graph API
   - Keep consistent across platforms

3. **Race Conditions**
   - User files email on Device A
   - User opens same email on Device B before sync
   - Device B should detect category change
   - Use polling or event listeners to detect external changes

4. **Testing Strategy**
   - Test on Windows Desktop (Office.js categories)
   - Test on Mac (Graph API fallback)
   - Test on Web (both methods)
   - Test master category creation (first time use)
   - Test category removal with confirmation
   - Test category sync across devices (manual test)

5. **Performance**
   - Category reads should be fast (<100ms)
   - Don't block filing workflow with category operations
   - Cache category status for current email
   - Debounce category polling (5s interval)

6. **User Experience**
   - Make category status always visible (prominent badge)
   - Provide quick access to category actions
   - Show visual feedback when categories change (animation)
   - Allow manual category management (power users)

## Definition of Done

- [ ] `CategoryBadge` component displays current category status
- [ ] Badge shows "Filed" (green), "Unfiled" (orange), or "None" (gray)
- [ ] `CategoryActionsMenu` provides quick actions
- [ ] "Remove Category" action requires confirmation
- [ ] Master categories created automatically on first use
- [ ] Office.js category API integrated
- [ ] Graph API fallback implemented for Mac/Web
- [ ] Category status updates in real-time after filing
- [ ] Error handling for category failures (doesn't block filing)
- [ ] Category sync detects external changes (optional polling)
- [ ] Manual "Apply Category" button for failed auto-apply
- [ ] Unit tests for category service functions
- [ ] Integration tests for category UI
- [ ] Tested on Windows, Mac, and Web platforms
- [ ] Documentation updated with category management guide
