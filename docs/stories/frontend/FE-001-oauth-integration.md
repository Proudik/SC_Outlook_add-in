# FE-001: OAuth Integration

**Story ID:** FE-001
**Story Points:** 8
**Epic Link:** Authentication & Authorization
**Status:** Ready for Development

## Description

Build from scratch a complete OAuth 2.0 authentication flow for the new React + TypeScript Outlook add-in. Users must authenticate with SingleCase API using OAuth (not public tokens), with support for both dialog-based and SSO flows. The implementation must handle token storage, refresh, expiration, and cross-runtime synchronization (taskpane vs commands runtime).

This is foundational work - no existing OAuth infrastructure exists in the new codebase. All authentication logic must be built from the ground up.

## Acceptance Criteria

1. **OAuth Dialog Flow**
   - User can initiate login via a branded "Sign In" button
   - Office.js `displayDialogAsync` opens OAuth provider's authorization page
   - Dialog handles OAuth callback with authorization code
   - Dialog sends token back to taskpane via `messageParent`
   - Taskpane receives and stores access token securely
   - Dialog closes automatically after successful authentication

2. **Token Management**
   - Access tokens stored in `sessionStorage` (taskpane runtime)
   - Tokens mirrored to `OfficeRuntime.storage` (commands runtime access)
   - Refresh tokens stored securely if provided by OAuth flow
   - Token expiry tracked with issued-at timestamp
   - Automatic token refresh before expiration (if refresh token available)
   - Manual re-authentication trigger when tokens are invalid/expired

3. **Token Validation**
   - Validate token format before storage
   - Test API call to `/publicapi/v1/cases` after receiving token
   - Display clear error messages for failed authentication
   - Retry logic for transient network failures during auth

4. **User Experience**
   - Loading spinner during authentication process
   - Success confirmation message after login
   - Error handling with user-friendly messages
   - "Sign Out" option that clears all stored tokens
   - Persist user email alongside token for display purposes

5. **Cross-Runtime Token Sync**
   - Taskpane runtime can read/write tokens
   - Commands runtime can read tokens via `OfficeRuntime.storage`
   - Token changes in taskpane propagate to commands runtime
   - Logout clears tokens from both runtimes

## Technical Requirements

### React Components

1. **AuthScreen.tsx** (new component)
   ```tsx
   interface AuthScreenProps {
     onAuthSuccess: (token: string, email: string) => void;
     onAuthError: (error: Error) => void;
   }
   ```
   - Display SingleCase branding
   - "Sign In with SingleCase" button
   - Loading state during authentication
   - Error display area
   - Support for dev mode (skip OAuth for local testing)

2. **OAuthDialog.tsx** (new component for dialog.html)
   - Runs in `dialog.html` window
   - Receives OAuth callback redirect
   - Parses authorization code from URL query params
   - Exchanges code for access token via API call
   - Sends token to parent window via `Office.context.ui.messageParent`

3. **Update App.tsx**
   - Check authentication state on mount
   - Show `AuthScreen` if no valid token
   - Show main workspace if authenticated
   - Handle logout and clear all auth state

### Services

1. **services/oauth.ts** (new service)
   ```typescript
   export async function initiateOAuthFlow(): Promise<void>;
   export async function handleOAuthCallback(code: string): Promise<{ token: string; email: string }>;
   export async function exchangeCodeForToken(code: string): Promise<{ access_token: string; refresh_token?: string; expires_in: number }>;
   export async function refreshAccessToken(refreshToken: string): Promise<{ access_token: string; expires_in: number }>;
   export function getOAuthAuthorizationUrl(): string;
   ```

2. **Update services/auth.ts**
   - Replace public token logic with OAuth token storage
   - Add `getAccessToken(): Promise<string>` that auto-refreshes if needed
   - Add `getRefreshToken(): string | null`
   - Add `setOAuthTokens(accessToken: string, refreshToken?: string, expiresIn?: number): Promise<void>`
   - Add `isTokenExpired(): boolean`
   - Update `clearAuth()` to clear OAuth tokens and refresh tokens

3. **services/tokenRefresh.ts** (new service)
   ```typescript
   export async function setupTokenRefreshInterval(): Promise<void>;
   export async function attemptTokenRefresh(): Promise<boolean>;
   export function clearRefreshInterval(): void;
   ```
   - Background refresh 5 minutes before expiration
   - Retry logic for failed refreshes
   - Clear interval on logout

### Office.js APIs

1. **Dialog API**
   - Use `Office.context.ui.displayDialogAsync()` for OAuth dialog
   - Configure dialog size: `{ height: 70, width: 45, displayInIframe: false }`
   - Handle `DialogMessageReceived` event for token message
   - Handle `DialogEventReceived` for dialog closure
   - Close dialog after successful auth or on user cancellation

2. **Storage**
   - Taskpane: Use `sessionStorage` for tokens (current session only)
   - Commands: Use `OfficeRuntime.storage.setItem/getItem` for cross-runtime access
   - Fallback to `Office.context.roamingSettings` if `OfficeRuntime.storage` unavailable

### Configuration

1. **Environment Variables** (add to .env)
   ```
   OAUTH_CLIENT_ID=your-oauth-client-id
   OAUTH_CLIENT_SECRET=your-oauth-client-secret
   OAUTH_AUTHORIZATION_URL=https://singlecase.auth.endpoint/authorize
   OAUTH_TOKEN_URL=https://singlecase.auth.endpoint/token
   OAUTH_REDIRECT_URI=https://localhost:3000/dialog.html
   OAUTH_SCOPES=cases:read cases:write clients:read
   ```

2. **Webpack Config**
   - Inject OAuth env vars via `webpack.DefinePlugin`
   - Update `dialog.html` entry point in webpack config

### API Integration Patterns

1. **Token Exchange Flow**
   ```typescript
   // In OAuthDialog.tsx after receiving callback
   const urlParams = new URLSearchParams(window.location.search);
   const code = urlParams.get('code');
   const { access_token, refresh_token, expires_in } = await exchangeCodeForToken(code);
   Office.context.ui.messageParent(JSON.stringify({
     type: 'auth_success',
     token: access_token,
     refreshToken: refresh_token,
     expiresIn: expires_in
   }));
   ```

2. **Automatic Token Refresh**
   ```typescript
   // In services/auth.ts
   export async function getAccessToken(): Promise<string> {
     const { token, issuedAt, expiresIn } = getAuth();
     const now = Date.now();
     const expiresAt = issuedAt + (expiresIn * 1000);

     // Refresh 5 minutes before expiry
     if (now > expiresAt - (5 * 60 * 1000)) {
       const refreshToken = getRefreshToken();
       if (refreshToken) {
         const newTokens = await refreshAccessToken(refreshToken);
         await setOAuthTokens(newTokens.access_token, refreshToken, newTokens.expires_in);
         return newTokens.access_token;
       }
       throw new Error('Token expired and no refresh token available');
     }

     return token;
   }
   ```

3. **Auth Error Handling**
   ```typescript
   // In services/singlecase.ts
   async function scRequest<T>(method: string, path: string, body?: any): Promise<T> {
     const token = await getAccessToken(); // Auto-refresh if needed
     const response = await fetch(url, {
       headers: {
         'Authentication': token,
         'Content-Type': 'application/json'
       }
     });

     if (response.status === 401) {
       // Token invalid, clear auth and force re-login
       await clearAuth();
       throw new Error('Authentication failed. Please sign in again.');
     }

     return response.json();
   }
   ```

## Reference Implementation

Review the demo's current auth implementation for patterns to adapt:

1. **Dialog Pattern**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/dialog/DialogAuth.tsx`
   - Shows how to handle dialog messaging
   - Demonstrates `messageParent` for token passing
   - Example of dialog size configuration

2. **Token Storage**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/auth.ts`
   - Current public token storage pattern (replace with OAuth)
   - Cross-runtime storage synchronization logic
   - Session expiry checking

3. **Auth Screen**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/AuthScreen.tsx`
   - UI/UX patterns for login screen
   - Error display patterns
   - Loading states

**Important**: The demo uses public token authentication. DO NOT copy this approach. Build OAuth flow from scratch, but use the demo's dialog management and storage patterns as reference.

## Dependencies

- **Blocks**: All other frontend stories (FE-002 through FE-010) require authentication
- **Requires**: OAuth provider configuration (client ID, secret, endpoints)
- **Requires**: Backend API must support OAuth token validation

## Notes

1. **Security Considerations**
   - Never log tokens to console in production builds
   - Use secure storage APIs (OfficeRuntime.storage preferred over localStorage)
   - Implement token rotation to minimize exposure window
   - Clear tokens on logout and on auth errors

2. **Testing Strategy**
   - Mock OAuth flow in development environment
   - Test token expiration scenarios
   - Test cross-runtime token access (taskpane → commands)
   - Test dialog closure edge cases (user closes dialog prematurely)
   - Test network failures during token exchange

3. **Error Scenarios**
   - User denies OAuth permission: Show friendly error, allow retry
   - Token exchange fails: Show technical error, log details
   - Dialog API unavailable: Fallback to inline login form (future enhancement)
   - Refresh token expired: Force full re-authentication

4. **Performance**
   - Token validation should not block UI rendering
   - Background token refresh should be silent (no UI interruption)
   - Dialog opening should feel instant (no loading delay)

5. **Accessibility**
   - Sign In button must have proper ARIA label
   - Error messages must be announced to screen readers
   - Dialog must be keyboard navigable
   - Focus management when returning from dialog

6. **Browser Compatibility**
   - Test in Outlook Web (all browsers)
   - Test in Outlook Desktop (Windows, Mac)
   - Test in Outlook Mobile (iOS, Android)
   - Note: Dialog API behavior varies by platform

## Definition of Done

- [ ] User can authenticate via OAuth dialog
- [ ] Access token stored securely in sessionStorage and OfficeRuntime.storage
- [ ] Refresh token stored and used for automatic token renewal
- [ ] Token expiration handled gracefully with auto-refresh
- [ ] Commands runtime can access tokens for on-send handler
- [ ] Sign out clears all tokens and returns to AuthScreen
- [ ] All error scenarios handled with user-friendly messages
- [ ] Code reviewed for security vulnerabilities
- [ ] Unit tests written for token management functions
- [ ] Integration tests verify full auth flow
- [ ] Documentation updated with OAuth setup instructions
