# FE-005: Error Handling & User Feedback System

**Story ID:** FE-005
**Story Points:** 5
**Epic Link:** User Experience & Polish
**Status:** Ready for Development

## Description

Build a comprehensive, centralized error handling and user feedback system for the new Outlook add-in. All API errors, Office.js errors, and user actions must provide clear, actionable feedback through consistent UI patterns (toasts, banners, modals). Implement graceful degradation for network failures, proper error logging, and user-friendly error messages that avoid technical jargon.

This is foundational UX infrastructure that will be used across all other stories (FE-001 through FE-010).

## Acceptance Criteria

1. **Toast Notifications**
   - Success toasts: Green, auto-dismiss after 3 seconds, show checkmark icon
   - Error toasts: Red, require manual dismiss, show error icon
   - Warning toasts: Yellow, auto-dismiss after 5 seconds, show warning icon
   - Info toasts: Blue, auto-dismiss after 3 seconds, show info icon
   - Max 3 toasts visible at once (stack vertically)
   - Toasts positioned at top-right of add-in window

2. **Error Messages**
   - User-friendly messages (no stack traces or technical codes visible)
   - Actionable suggestions (e.g., "Check your internet connection", "Try again")
   - Show "Report Problem" button for unexpected errors
   - Display error ID for support reference
   - Categorize errors: Network, Auth, Validation, Server, Unknown

3. **Loading States**
   - Global loading overlay for full-screen operations
   - Inline spinners for component-level loading
   - Skeleton loaders for data fetching (case list, suggestions)
   - Progress bars for long operations (attachment uploads)
   - Disable buttons during async operations

4. **Offline Detection**
   - Detect when add-in loses internet connection
   - Show persistent banner: "You're offline. Some features unavailable."
   - Automatically hide banner when connection restored
   - Disable features that require internet (filing, suggestions)
   - Allow viewing cached data offline

5. **Error Recovery**
   - Retry buttons for transient errors (network, timeout)
   - "Refresh" button to reload data after error
   - "Sign in again" button for auth errors
   - Clear instructions on how to recover from each error type

6. **Error Logging**
   - Log all errors to browser console (development)
   - Send error reports to backend API (production)
   - Include: error message, stack trace, user context, timestamp
   - Respect user privacy (no PII in logs without consent)

## Technical Requirements

### React Components

1. **ToastProvider.tsx** (new context provider)
   ```tsx
   interface Toast {
     id: string;
     type: 'success' | 'error' | 'warning' | 'info';
     message: string;
     duration?: number; // ms, 0 = no auto-dismiss
     action?: {
       label: string;
       onClick: () => void;
     };
   }

   interface ToastContextValue {
     toasts: Toast[];
     showToast: (toast: Omit<Toast, 'id'>) => void;
     dismissToast: (id: string) => void;
     showSuccess: (message: string) => void;
     showError: (message: string, action?: Toast['action']) => void;
     showWarning: (message: string) => void;
     showInfo: (message: string) => void;
   }

   export function ToastProvider({ children }: { children: React.ReactNode });
   export function useToast(): ToastContextValue;
   ```

2. **ToastContainer.tsx** (new component)
   - Render active toasts in top-right corner
   - Stack toasts vertically with animation
   - Auto-dismiss after duration
   - Click to dismiss manually
   - Use Fluent UI `MessageBar` or custom styled component

3. **ErrorBoundary.tsx** (new component)
   ```tsx
   interface ErrorBoundaryProps {
     children: React.ReactNode;
     fallback?: React.ReactNode;
     onError?: (error: Error, errorInfo: React.ErrorInfo) => void;
   }
   ```
   - Catch React rendering errors
   - Display user-friendly error page
   - Log error to console and backend
   - Provide "Reload Add-in" button

4. **OfflineBanner.tsx** (new component)
   - Show banner at top of add-in when offline
   - Auto-hide when connection restored
   - Use `navigator.onLine` and `online`/`offline` events

5. **LoadingOverlay.tsx** (new component)
   ```tsx
   interface LoadingOverlayProps {
     isLoading: boolean;
     message?: string;
     progress?: number; // 0-100
   }
   ```
   - Full-screen overlay with spinner
   - Optional loading message
   - Optional progress bar

6. **ErrorDisplay.tsx** (new component)
   ```tsx
   interface ErrorDisplayProps {
     error: AppError;
     onRetry?: () => void;
     onDismiss?: () => void;
   }
   ```
   - Display formatted error message
   - Show appropriate icon (based on error type)
   - Show retry button if applicable
   - Show error ID for support

### Services

1. **services/errorHandler.ts** (new service)
   ```typescript
   export enum ErrorType {
     Network = 'network',
     Auth = 'auth',
     Validation = 'validation',
     ServerError = 'server_error',
     OfficeJs = 'office_js',
     Unknown = 'unknown',
   }

   export interface AppError {
     id: string; // unique error ID
     type: ErrorType;
     message: string; // user-friendly message
     technicalMessage?: string; // developer message
     statusCode?: number;
     canRetry: boolean;
     originalError?: Error;
   }

   export function createAppError(
     error: unknown,
     type?: ErrorType
   ): AppError;

   export function formatErrorMessage(error: AppError): string;

   export function isNetworkError(error: unknown): boolean;
   export function isAuthError(error: unknown): boolean;
   ```

2. **services/errorReporting.ts** (new service)
   ```typescript
   export interface ErrorReport {
     errorId: string;
     message: string;
     stack?: string;
     userAgent: string;
     timestamp: string;
     context: {
       route?: string;
       action?: string;
       userId?: string;
     };
   }

   export async function reportError(
     error: AppError,
     context?: ErrorReport['context']
   ): Promise<void>;

   export function logError(error: AppError): void;
   ```

3. **services/networkMonitor.ts** (new service)
   ```typescript
   export type NetworkStatus = 'online' | 'offline';

   export function getCurrentNetworkStatus(): NetworkStatus;

   export function subscribeToNetworkChanges(
     callback: (status: NetworkStatus) => void
   ): () => void;

   export function waitForOnline(timeoutMs?: number): Promise<void>;
   ```

4. **Update services/singlecase.ts**
   - Wrap all API calls with error handling
   - Convert fetch errors to `AppError`
   - Add request timeout handling
   - Add retry logic for network errors

### Hooks

1. **hooks/useNetworkStatus.ts** (new hook)
   ```typescript
   export function useNetworkStatus(): {
     isOnline: boolean;
     isOffline: boolean;
   }
   ```

2. **hooks/useErrorHandler.ts** (new hook)
   ```typescript
   export function useErrorHandler(): {
     handleError: (error: unknown, context?: string) => void;
     clearError: () => void;
     error: AppError | null;
   }
   ```

3. **hooks/useRetryableRequest.ts** (new hook)
   ```typescript
   export function useRetryableRequest<T>(
     requestFn: () => Promise<T>,
     options?: {
       maxRetries?: number;
       retryDelay?: number;
       shouldRetry?: (error: unknown) => boolean;
     }
   ): {
     execute: () => Promise<T>;
     isLoading: boolean;
     error: AppError | null;
     retry: () => Promise<T>;
   }
   ```

### Error Message Templates

1. **Network Errors**
   - "Unable to connect. Check your internet connection and try again."
   - "Request timed out. Please try again."
   - "Connection lost. Retrying..."

2. **Auth Errors**
   - "Your session has expired. Please sign in again."
   - "Authentication failed. Please check your credentials."
   - "You don't have permission to perform this action."

3. **Validation Errors**
   - "Please select a case before filing."
   - "Email subject cannot be empty."
   - "Attachment is too large (max 25MB)."

4. **Server Errors**
   - "Something went wrong on our end. Please try again later."
   - "The service is temporarily unavailable. (Error ID: {errorId})"
   - "Failed to save email. Please contact support with error ID: {errorId}"

5. **Office.js Errors**
   - "Unable to access email data. Please reopen the add-in."
   - "Failed to apply category. The email was filed successfully."
   - "Attachment download failed. Try again or file without attachments."

## API Integration Patterns

1. **Centralized Error Handling**
   ```typescript
   // In services/singlecase.ts
   async function scRequest<T>(
     method: string,
     path: string,
     token: string,
     body?: any
   ): Promise<T> {
     try {
       const response = await fetch(url, { method, headers, body });

       if (!response.ok) {
         throw createAppError(
           new Error(`HTTP ${response.status}`),
           response.status === 401 ? ErrorType.Auth :
           response.status >= 500 ? ErrorType.ServerError :
           ErrorType.Validation
         );
       }

       return await response.json();
     } catch (error) {
       if (error instanceof TypeError) {
         // Network error
         throw createAppError(error, ErrorType.Network);
       }
       throw createAppError(error);
     }
   }
   ```

2. **Component Error Handling**
   ```typescript
   // In MainWorkspace.tsx
   const { showError, showSuccess } = useToast();
   const { handleError } = useErrorHandler();

   const handleFileEmail = async () => {
     try {
       setIsLoading(true);
       await fileEmailToCase(token, caseId, payload);
       showSuccess('Email filed successfully!');
     } catch (error) {
       const appError = createAppError(error);
       handleError(appError, 'file_email');

       if (appError.canRetry) {
         showError(appError.message, {
           label: 'Retry',
           onClick: handleFileEmail,
         });
       } else {
         showError(appError.message);
       }
     } finally {
       setIsLoading(false);
     }
   };
   ```

3. **Offline Detection**
   ```typescript
   // In App.tsx
   const { isOffline } = useNetworkStatus();

   return (
     <>
       {isOffline && <OfflineBanner />}
       <MainContent />
     </>
   );
   ```

4. **Error Boundary Integration**
   ```typescript
   // In App.tsx
   <ErrorBoundary
     fallback={
       <ErrorDisplay
         error={createAppError(new Error('Something went wrong'))}
         onRetry={() => window.location.reload()}
       />
     }
     onError={(error, errorInfo) => {
       const appError = createAppError(error);
       reportError(appError, { route: 'app' });
       console.error('React error:', error, errorInfo);
     }}
   >
     <AppContent />
   </ErrorBoundary>
   ```

## Reference Implementation

Review the demo for patterns to expand upon:

1. **Current Error Handling**: Look for `try/catch` blocks throughout demo
   - **IMPROVE** by centralizing error handling
   - Add consistent user feedback for all errors

2. **Toast/Message Patterns**: Check if demo uses any notification system
   - Build comprehensive toast system from scratch

3. **Loading States**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/MainWorkspace/MainWorkspace.tsx`
   - Identify existing loading states
   - Standardize with new `LoadingOverlay` component

## Dependencies

- **Used by**: All frontend stories (FE-001 through FE-010)
- **Requires**: Fluent UI components for consistent styling
- **Optional**: Backend error reporting endpoint (can log to console initially)

## Notes

1. **User-Friendly Error Messages**
   - Avoid: "Failed to fetch", "Network request failed", "500 Internal Server Error"
   - Use: "Unable to connect", "Something went wrong", "Please try again"
   - Always provide next steps or recovery actions

2. **Error Reporting Privacy**
   - Don't send PII (email addresses, names) in error reports
   - Anonymize user identifiers
   - Get user consent before sending reports (opt-in on first error)

3. **Testing Strategy**
   - Simulate network errors (offline mode, slow 3G)
   - Test all error types (auth, validation, server)
   - Test toast stacking (show 5+ toasts rapidly)
   - Test error boundary (throw error in component)
   - Test retry logic (fail → retry → success)

4. **Performance**
   - Toast animations should be smooth (60fps)
   - Error logging should not block UI
   - Error reporting should be async (fire-and-forget)
   - Don't spam error reports (deduplicate similar errors)

5. **Accessibility**
   - Error messages must be announced to screen readers
   - Use ARIA live regions for toasts
   - Ensure error displays are keyboard navigable
   - High contrast mode support for error colors

6. **Localization**
   - All error messages must be translatable
   - Store error message templates in i18n files (future)
   - Technical error details can remain in English

## Definition of Done

- [ ] `ToastProvider` context and hooks created
- [ ] `ToastContainer` component displays stacked toasts
- [ ] Success, error, warning, info toast variants implemented
- [ ] `ErrorBoundary` catches and displays React errors
- [ ] `OfflineBanner` shows/hides based on network status
- [ ] `LoadingOverlay` component for full-screen loading
- [ ] `AppError` type and creation function implemented
- [ ] Error reporting service sends logs to backend
- [ ] Network monitor detects online/offline changes
- [ ] All error message templates defined
- [ ] User-friendly error messages for all error types
- [ ] Retry buttons work for retryable errors
- [ ] Error IDs generated and displayed for support
- [ ] Unit tests for error handling utilities
- [ ] Integration tests for toast display
- [ ] Accessibility tested with screen reader
- [ ] Documentation updated with error handling patterns
