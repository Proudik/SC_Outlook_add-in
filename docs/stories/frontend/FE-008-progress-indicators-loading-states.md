# FE-008: Progress Indicators & Loading States

**Story ID:** FE-008
**Story Points:** 3
**Epic Link:** User Experience & Polish
**Status:** Ready for Development

## Description

Implement comprehensive loading states and progress indicators throughout the add-in to provide clear visual feedback during async operations. Replace generic spinners with contextual progress indicators that show what's happening and estimated completion. Ensure consistent loading UX across all features (filing, suggestions, duplicate checks, case loading).

This story focuses on creating a polished, professional user experience that keeps users informed during waits.

## Acceptance Criteria

1. **Skeleton Loaders**
   - Case list loading: Show skeleton cards instead of spinner
   - Suggestions loading: Show skeleton suggestion cards
   - Email metadata loading: Show skeleton text blocks
   - Skeleton loaders animate (shimmer effect)
   - Smooth transition from skeleton to real content

2. **Progress Bars**
   - File upload: Show progress bar (0-100%)
   - Attachment download: Show "Downloading 2 of 5 attachments"
   - Large operations: Show determinate progress when possible
   - Indeterminate progress bar for unknown duration tasks

3. **Inline Spinners**
   - Button loading: Show spinner inside button, disable during load
   - Icon button loading: Replace icon with small spinner
   - Input loading: Show spinner at right edge of input
   - Consistent spinner size and color across components

4. **Loading Messages**
   - Show contextual message explaining what's happening
   - "Loading cases...", "Checking for duplicates...", "Filing email..."
   - Update message as operation progresses
   - Show estimated time for long operations (>3s)

5. **Optimistic UI Updates**
   - Update UI immediately on user action (e.g., select case)
   - Show loading indicator in background
   - Revert on failure with error message
   - Don't block UI for fast operations (<200ms)

6. **Loading State Hierarchy**
   - Full-screen overlay: For blocking operations (initial load, critical errors)
   - Modal/dialog loading: For async operations within modal
   - Component-level loading: For isolated component updates
   - Never show multiple conflicting loading indicators

## Technical Requirements

### React Components

1. **SkeletonLoader.tsx** (new component)
   ```tsx
   interface SkeletonLoaderProps {
     variant: 'text' | 'card' | 'list' | 'avatar' | 'rect';
     width?: string | number;
     height?: string | number;
     count?: number; // For multiple skeleton items
     animate?: boolean;
   }
   ```
   - Animated shimmer effect
   - Different variants for different content types
   - Use Fluent UI `Skeleton` or build custom

2. **ProgressBar.tsx** (new component)
   ```tsx
   interface ProgressBarProps {
     value: number; // 0-100 or null for indeterminate
     label?: string;
     showPercentage?: boolean;
     color?: 'primary' | 'success' | 'warning';
   }
   ```
   - Determinate progress bar (known progress)
   - Indeterminate progress bar (unknown progress)
   - Optional label and percentage display
   - Smooth animation on value changes

3. **LoadingButton.tsx** (new component)
   ```tsx
   interface LoadingButtonProps extends ButtonProps {
     isLoading: boolean;
     loadingText?: string;
     spinner?: React.ReactNode;
   }
   ```
   - Show spinner inside button when loading
   - Disable button during loading
   - Optional different text during loading
   - Preserve button width during loading

4. **LoadingOverlay.tsx** (enhance existing component from FE-005)
   - Add optional progress bar
   - Add optional loading message
   - Add optional substep indicator (e.g., "Step 2 of 4")

5. **CaseListSkeleton.tsx** (new component)
   - Skeleton loader specifically for case list
   - Shows 5-10 skeleton case cards
   - Matches actual case card dimensions

6. **SuggestionsSkeleton.tsx** (new component)
   - Skeleton loader for case suggestions
   - Shows 3 skeleton suggestion cards
   - Matches actual suggestion card dimensions

### Hooks

1. **hooks/useLoadingState.ts** (new hook)
   ```typescript
   export function useLoadingState(initialState?: boolean): {
     isLoading: boolean;
     startLoading: () => void;
     stopLoading: () => void;
     withLoading: <T>(fn: () => Promise<T>) => Promise<T>;
   }
   ```
   - Manage loading state for async operations
   - `withLoading` wrapper automatically sets loading state
   - Prevents double-loading (ignore if already loading)

2. **hooks/useProgress.ts** (new hook)
   ```typescript
   export function useProgress(): {
     progress: number; // 0-100 or null
     setProgress: (value: number | null) => void;
     reset: () => void;
     increment: (amount: number) => void;
   }
   ```
   - Track progress for multi-step operations
   - Smooth progress transitions

3. **hooks/useOptimisticUpdate.ts** (new hook)
   ```typescript
   export function useOptimisticUpdate<T>(
     initialValue: T,
     updateFn: (value: T) => Promise<void>
   ): {
     value: T;
     optimisticUpdate: (newValue: T) => Promise<void>;
     isUpdating: boolean;
     error: Error | null;
   }
   ```
   - Update UI immediately (optimistic)
   - Revert on API failure
   - Show loading indicator during update

### Loading State Patterns

1. **Initial Data Loading**
   ```typescript
   // In MainWorkspace.tsx
   const [cases, setCases] = React.useState<CaseOption[]>([]);
   const [isLoadingCases, setIsLoadingCases] = React.useState(true);

   React.useEffect(() => {
     setIsLoadingCases(true);
     listCases(token)
       .then(setCases)
       .catch(error => showError('Failed to load cases'))
       .finally(() => setIsLoadingCases(false));
   }, [token]);

   return (
     <>
       {isLoadingCases ? (
         <CaseListSkeleton count={8} />
       ) : (
         <CaseList cases={cases} />
       )}
     </>
   );
   ```

2. **Button Loading State**
   ```typescript
   const [isFilingprogress, setIsFilingprogress] = React.useState(false);

   const handleFileEmail = async () => {
     setIsFilingprogress(true);
     try {
       await fileEmailToCase(token, caseId, payload);
       showSuccess('Email filed!');
     } catch (error) {
       showError('Filing failed');
     } finally {
       setIsFilingprogress(false);
     }
   };

   return (
     <LoadingButton
       isLoading={isFilingprogress}
       loadingText="Filing..."
       onClick={handleFileEmail}
     >
       File Email
     </LoadingButton>
   );
   ```

3. **Progress Bar for Uploads**
   ```typescript
   const [uploadProgress, setUploadProgress] = React.useState<number | null>(null);

   const uploadAttachments = async (attachments: Attachment[]) => {
     setUploadProgress(0);

     for (let i = 0; i < attachments.length; i++) {
       await uploadAttachment(attachments[i]);
       setUploadProgress(((i + 1) / attachments.length) * 100);
     }

     setUploadProgress(null);
   };

   return (
     <>
       {uploadProgress !== null && (
         <ProgressBar
           value={uploadProgress}
           label={`Uploading attachments...`}
           showPercentage
         />
       )}
     </>
   );
   ```

4. **Optimistic Favorite Toggle**
   ```typescript
   const { optimisticUpdate, isUpdating } = useOptimisticUpdate(
     isFavorite,
     async (newValue) => {
       await toggleFavoriteCase(token, caseId, newValue);
     }
   );

   return (
     <IconButton
       icon={isFavorite ? <StarFilled /> : <Star />}
       onClick={() => optimisticUpdate(!isFavorite)}
       disabled={isUpdating}
     />
   );
   ```

5. **Multi-Step Progress**
   ```typescript
   const { progress, setProgress } = useProgress();

   const handleFileWithSteps = async () => {
     setProgress(0);

     // Step 1: Collect metadata (25%)
     await collectMetadata();
     setProgress(25);

     // Step 2: Download attachments (50%)
     await downloadAttachments();
     setProgress(50);

     // Step 3: Upload to server (75%)
     await uploadToServer();
     setProgress(75);

     // Step 4: Apply category (100%)
     await applyCategory();
     setProgress(100);

     setTimeout(() => setProgress(null), 500); // Clear after animation
   };

   return (
     <Modal>
       <ProgressBar value={progress} />
       <LoadingMessage>
         {progress < 25 && 'Collecting email data...'}
         {progress >= 25 && progress < 50 && 'Downloading attachments...'}
         {progress >= 50 && progress < 75 && 'Uploading to SingleCase...'}
         {progress >= 75 && 'Finalizing...'}
       </LoadingMessage>
     </Modal>
   );
   ```

## Reference Implementation

Review the demo for existing loading patterns:

1. **Current Loading States**: Look throughout demo components
   - Identify inconsistent loading patterns
   - Standardize with new components

2. **Case List Loading**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/MainWorkspace/MainWorkspace.tsx`
   - Replace simple spinner with skeleton loader

3. **Button States**: Check all buttons in demo
   - Replace with `LoadingButton` component

## Dependencies

- **Requires**: FE-005 (Error Handling) - integrate with loading states
- **Used by**: All feature stories (FE-001 through FE-010)
- **Requires**: Fluent UI components for base styling

## Notes

1. **Performance**
   - Skeleton loaders should not impact render performance
   - Smooth animations (60fps)
   - Lazy load skeleton components if large
   - Use CSS animations over JS animations

2. **Loading State Best Practices**
   - Show loading indicator only if operation takes >200ms (avoid flashing)
   - Use skeleton loaders for content that loads on mount
   - Use spinners for user-triggered actions
   - Use progress bars for operations with known duration
   - Never show multiple overlapping loading indicators

3. **User Experience**
   - Be specific with loading messages ("Loading cases" not "Loading...")
   - Show progress when possible (builds trust)
   - Don't block UI unnecessarily (use optimistic updates)
   - Show "Still working..." message for operations >10s

4. **Testing Strategy**
   - Test loading states with fast network (edge case)
   - Test with slow network (3G throttling)
   - Test rapid user actions (clicking multiple buttons)
   - Test skeleton loader transitions
   - Test progress bar accuracy

5. **Accessibility**
   - Use ARIA `role="progressbar"` for progress bars
   - Use ARIA `aria-busy="true"` for loading regions
   - Use ARIA live regions to announce loading state changes
   - Ensure loading indicators are visible in high contrast mode
   - Don't rely only on color to indicate loading

6. **Skeleton Loader Design**
   - Match skeleton dimensions to real content
   - Use subtle shimmer animation (not distracting)
   - Gray background, lighter shimmer overlay
   - Rounded corners to match actual UI
   - Fade in real content smoothly

## Definition of Done

- [ ] `SkeletonLoader` component with multiple variants
- [ ] `ProgressBar` component (determinate and indeterminate)
- [ ] `LoadingButton` component with inline spinner
- [ ] `CaseListSkeleton` for case list loading
- [ ] `SuggestionsSkeleton` for suggestions loading
- [ ] `LoadingOverlay` enhanced with progress support
- [ ] All async operations show appropriate loading state
- [ ] Loading messages are contextual and clear
- [ ] Progress bars used for file uploads
- [ ] Optimistic updates implemented for favorites
- [ ] No loading state flashing (<200ms operations)
- [ ] Smooth transitions from loading to content
- [ ] All loading states accessible (ARIA attributes)
- [ ] Unit tests for loading hooks
- [ ] Visual tests for skeleton loaders
- [ ] Tested on slow network (3G)
- [ ] Documentation updated with loading patterns
