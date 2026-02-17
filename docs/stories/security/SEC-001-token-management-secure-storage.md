# SEC-001: Token Management & Secure Storage

**Story ID:** SEC-001
**Story Points:** 5
**Epic Link:** Security & Compliance
**Status:** Ready for Development

## Description

Implement secure token management and storage infrastructure for OAuth 2.0 access tokens, refresh tokens, and sensitive authentication data. This story establishes production-grade security patterns for token lifecycle management, including secure storage, encryption at rest, token rotation, and secure deletion practices.

The implementation must comply with OAuth 2.0 security best practices (RFC 6749, RFC 8252), OWASP Token Storage guidelines, and enterprise security requirements for Office Add-ins.

## Acceptance Criteria

1. **Secure Token Storage**
   - Use Office.js Roaming Settings for cross-device token sync (if applicable)
   - Implement browser SessionStorage for temporary tokens (in-memory preferred)
   - Never store tokens in LocalStorage (vulnerable to XSS)
   - Never log tokens or include in console output
   - Encrypt sensitive tokens before storage (if required by compliance)
   - Clear all tokens on logout or session expiration

2. **Token Lifecycle Management**
   - Access tokens expire after configurable TTL (default: 1 hour)
   - Refresh tokens valid for configurable period (default: 30 days)
   - Automatic token refresh 5 minutes before expiration
   - Graceful handling of expired tokens (auto-refresh or re-auth)
   - Token revocation on logout (call revoke endpoint if available)
   - Secure token cleanup on add-in shutdown

3. **Token Encryption**
   - Encrypt access tokens before storing in Roaming Settings
   - Use Web Crypto API for encryption (AES-GCM 256-bit)
   - Generate unique encryption key per session (ephemeral key)
   - Never store encryption keys in code or config
   - Use PKCE (Proof Key for Code Exchange) for OAuth flow
   - Implement secure key derivation (PBKDF2 or similar)

4. **Token Transmission Security**
   - Always use HTTPS for token transmission
   - Include tokens in Authorization header (not URL query params)
   - Use Bearer token format: `Authorization: Bearer <token>`
   - Never include tokens in URL query strings or POST body (unless OAuth standard)
   - Validate token format before transmission
   - Add request signing for critical operations (optional)

5. **Token Validation & Verification**
   - Validate token format (JWT structure, base64 encoding)
   - Verify token signature (if JWT with public key available)
   - Check token expiration before use (exp claim)
   - Validate issuer (iss claim) matches expected OAuth provider
   - Validate audience (aud claim) matches application
   - Reject tokens with suspicious claims or tampering

6. **Security Monitoring & Logging**
   - Log token issuance events (without token values)
   - Log token refresh attempts and failures
   - Log token validation failures
   - Detect and alert on suspicious token activity (multiple failures)
   - Track token age and rotation compliance
   - Monitor for token reuse attacks

7. **Error Handling**
   - 401 Unauthorized: Invalid or expired token (trigger re-auth)
   - 403 Forbidden: Valid token, insufficient permissions
   - Handle token refresh failures gracefully (re-authenticate user)
   - Display user-friendly error messages (hide technical details)
   - Implement exponential backoff for token refresh retries
   - Clear invalid tokens immediately

8. **OAuth 2.0 Security Patterns**
   - Use Authorization Code Flow with PKCE (not Implicit Flow)
   - Generate cryptographically random PKCE code verifier (43-128 chars)
   - Use SHA-256 for PKCE code challenge
   - Implement state parameter for CSRF protection (random 32-byte value)
   - Validate state parameter in OAuth callback
   - Use secure redirect URI (HTTPS, no wildcards)

## Technical Requirements

### Token Storage Architecture

1. **Storage Layer Abstraction**
   ```typescript
   // src/services/tokenStorage.ts

   export interface TokenData {
     accessToken: string;
     refreshToken?: string;
     expiresAt: number; // Unix timestamp
     tokenType: 'Bearer';
     scope?: string;
     issuedAt: number;
   }

   export interface SecureStorage {
     storeToken(token: TokenData): Promise<void>;
     retrieveToken(): Promise<TokenData | null>;
     clearToken(): Promise<void>;
     hasValidToken(): Promise<boolean>;
   }

   class OfficeSecureStorage implements SecureStorage {
     private readonly STORAGE_KEY = 'sc_auth_token_encrypted';
     private readonly IV_KEY = 'sc_auth_iv';

     async storeToken(token: TokenData): Promise<void> {
       try {
         // Encrypt token data
         const encrypted = await this.encryptToken(token);

         // Store in Office.js Roaming Settings
         await Office.context.roamingSettings.set(this.STORAGE_KEY, encrypted.ciphertext);
         await Office.context.roamingSettings.set(this.IV_KEY, encrypted.iv);
         await Office.context.roamingSettings.saveAsync();

         // Also keep in session storage for quick access
         sessionStorage.setItem('sc_token_cache', JSON.stringify(token));

         console.log('[TokenStorage] Token stored securely');
       } catch (error) {
         console.error('[TokenStorage] Failed to store token', error);
         throw new Error('Failed to store authentication token');
       }
     }

     async retrieveToken(): Promise<TokenData | null> {
       try {
         // Try session storage first (performance)
         const cached = sessionStorage.getItem('sc_token_cache');
         if (cached) {
           const token = JSON.parse(cached) as TokenData;
           if (this.isTokenValid(token)) {
             return token;
           }
           // Token expired, clear cache
           sessionStorage.removeItem('sc_token_cache');
         }

         // Retrieve from Office.js Roaming Settings
         const ciphertext = Office.context.roamingSettings.get(this.STORAGE_KEY);
         const iv = Office.context.roamingSettings.get(this.IV_KEY);

         if (!ciphertext || !iv) {
           return null;
         }

         // Decrypt token
         const token = await this.decryptToken({ ciphertext, iv });

         // Validate token
         if (!this.isTokenValid(token)) {
           await this.clearToken();
           return null;
         }

         // Update session cache
         sessionStorage.setItem('sc_token_cache', JSON.stringify(token));

         return token;
       } catch (error) {
         console.error('[TokenStorage] Failed to retrieve token', error);
         return null;
       }
     }

     async clearToken(): Promise<void> {
       try {
         // Clear from Office.js Roaming Settings
         Office.context.roamingSettings.remove(this.STORAGE_KEY);
         Office.context.roamingSettings.remove(this.IV_KEY);
         await Office.context.roamingSettings.saveAsync();

         // Clear session storage
         sessionStorage.removeItem('sc_token_cache');

         console.log('[TokenStorage] Token cleared');
       } catch (error) {
         console.error('[TokenStorage] Failed to clear token', error);
       }
     }

     async hasValidToken(): Promise<boolean> {
       const token = await this.retrieveToken();
       return token !== null && this.isTokenValid(token);
     }

     private isTokenValid(token: TokenData): boolean {
       const now = Date.now();
       const bufferMs = 5 * 60 * 1000; // 5 minutes buffer
       return token.expiresAt > now + bufferMs;
     }

     private async encryptToken(token: TokenData): Promise<{ ciphertext: string; iv: string }> {
       // Use Web Crypto API for AES-GCM encryption
       const encoder = new TextEncoder();
       const data = encoder.encode(JSON.stringify(token));

       // Generate encryption key (ephemeral, derived from session)
       const key = await this.getOrCreateEncryptionKey();

       // Generate random IV
       const iv = crypto.getRandomValues(new Uint8Array(12));

       // Encrypt
       const encrypted = await crypto.subtle.encrypt(
         { name: 'AES-GCM', iv },
         key,
         data
       );

       return {
         ciphertext: this.arrayBufferToBase64(encrypted),
         iv: this.arrayBufferToBase64(iv),
       };
     }

     private async decryptToken(encrypted: { ciphertext: string; iv: string }): Promise<TokenData> {
       const key = await this.getOrCreateEncryptionKey();

       const ciphertext = this.base64ToArrayBuffer(encrypted.ciphertext);
       const iv = this.base64ToArrayBuffer(encrypted.iv);

       const decrypted = await crypto.subtle.decrypt(
         { name: 'AES-GCM', iv },
         key,
         ciphertext
       );

       const decoder = new TextDecoder();
       const json = decoder.decode(decrypted);
       return JSON.parse(json) as TokenData;
     }

     private async getOrCreateEncryptionKey(): Promise<CryptoKey> {
       // In production, derive from secure source (user session, hardware token, etc.)
       // For now, use session-based key
       const keyData = sessionStorage.getItem('sc_encryption_key');

       if (keyData) {
         const rawKey = this.base64ToArrayBuffer(keyData);
         return crypto.subtle.importKey(
           'raw',
           rawKey,
           { name: 'AES-GCM' },
           false,
           ['encrypt', 'decrypt']
         );
       }

       // Generate new key
       const key = await crypto.subtle.generateKey(
         { name: 'AES-GCM', length: 256 },
         true,
         ['encrypt', 'decrypt']
       );

       // Store in session (ephemeral)
       const exportedKey = await crypto.subtle.exportKey('raw', key);
       sessionStorage.setItem('sc_encryption_key', this.arrayBufferToBase64(exportedKey));

       return key;
     }

     private arrayBufferToBase64(buffer: ArrayBuffer): string {
       const bytes = new Uint8Array(buffer);
       let binary = '';
       for (let i = 0; i < bytes.length; i++) {
         binary += String.fromCharCode(bytes[i]);
       }
       return btoa(binary);
     }

     private base64ToArrayBuffer(base64: string): ArrayBuffer {
       const binary = atob(base64);
       const bytes = new Uint8Array(binary.length);
       for (let i = 0; i < binary.length; i++) {
         bytes[i] = binary.charCodeAt(i);
       }
       return bytes.buffer;
     }
   }

   // Export singleton
   export const tokenStorage: SecureStorage = new OfficeSecureStorage();
   ```

2. **Token Refresh Service**
   ```typescript
   // src/services/tokenRefresh.ts

   import { tokenStorage, TokenData } from './tokenStorage';
   import { msalConfig } from '../config/authConfig';

   class TokenRefreshService {
     private refreshTimer: NodeJS.Timeout | null = null;
     private isRefreshing = false;

     async initialize(): Promise<void> {
       // Check if we have a valid token
       const hasToken = await tokenStorage.hasValidToken();
       if (hasToken) {
         this.scheduleTokenRefresh();
       }
     }

     async refreshTokenIfNeeded(): Promise<boolean> {
       if (this.isRefreshing) {
         console.log('[TokenRefresh] Refresh already in progress');
         return false;
       }

       try {
         this.isRefreshing = true;

         const token = await tokenStorage.retrieveToken();
         if (!token) {
           console.log('[TokenRefresh] No token to refresh');
           return false;
         }

         // Check if refresh is needed (within 5 minutes of expiration)
         const now = Date.now();
         const bufferMs = 5 * 60 * 1000;

         if (token.expiresAt > now + bufferMs) {
           console.log('[TokenRefresh] Token still valid, no refresh needed');
           return true;
         }

         if (!token.refreshToken) {
           console.log('[TokenRefresh] No refresh token available, re-authentication required');
           await tokenStorage.clearToken();
           return false;
         }

         // Attempt token refresh
         console.log('[TokenRefresh] Refreshing token');
         const newToken = await this.performTokenRefresh(token.refreshToken);

         // Store new token
         await tokenStorage.storeToken(newToken);

         // Schedule next refresh
         this.scheduleTokenRefresh();

         console.log('[TokenRefresh] Token refreshed successfully');
         return true;

       } catch (error) {
         console.error('[TokenRefresh] Failed to refresh token', error);
         await tokenStorage.clearToken();
         return false;
       } finally {
         this.isRefreshing = false;
       }
     }

     private async performTokenRefresh(refreshToken: string): Promise<TokenData> {
       // Call OAuth token endpoint
       const response = await fetch(msalConfig.auth.authority + '/oauth2/v2.0/token', {
         method: 'POST',
         headers: {
           'Content-Type': 'application/x-www-form-urlencoded',
         },
         body: new URLSearchParams({
           client_id: msalConfig.auth.clientId,
           grant_type: 'refresh_token',
           refresh_token: refreshToken,
           scope: msalConfig.auth.scopes.join(' '),
         }),
       });

       if (!response.ok) {
         throw new Error(`Token refresh failed: ${response.status}`);
       }

       const data = await response.json();

       return {
         accessToken: data.access_token,
         refreshToken: data.refresh_token || refreshToken, // Some providers don't return new refresh token
         expiresAt: Date.now() + (data.expires_in * 1000),
         tokenType: 'Bearer',
         scope: data.scope,
         issuedAt: Date.now(),
       };
     }

     private scheduleTokenRefresh(): void {
       // Clear existing timer
       if (this.refreshTimer) {
         clearTimeout(this.refreshTimer);
       }

       // Schedule refresh for 5 minutes before expiration
       tokenStorage.retrieveToken().then(token => {
         if (!token) return;

         const now = Date.now();
         const bufferMs = 5 * 60 * 1000; // 5 minutes
         const refreshAt = token.expiresAt - bufferMs;
         const delayMs = Math.max(0, refreshAt - now);

         console.log(`[TokenRefresh] Next refresh scheduled in ${Math.round(delayMs / 1000)}s`);

         this.refreshTimer = setTimeout(() => {
           this.refreshTokenIfNeeded();
         }, delayMs);
       });
     }

     cleanup(): void {
       if (this.refreshTimer) {
         clearTimeout(this.refreshTimer);
         this.refreshTimer = null;
       }
     }
   }

   export const tokenRefreshService = new TokenRefreshService();
   ```

3. **OAuth PKCE Implementation**
   ```typescript
   // src/services/pkce.ts

   export class PKCEService {
     /**
      * Generate cryptographically random code verifier (43-128 characters)
      * RFC 7636 Section 4.1
      */
     generateCodeVerifier(): string {
       const array = new Uint8Array(32);
       crypto.getRandomValues(array);
       return this.base64URLEncode(array);
     }

     /**
      * Generate code challenge from verifier using SHA-256
      * RFC 7636 Section 4.2
      */
     async generateCodeChallenge(verifier: string): Promise<string> {
       const encoder = new TextEncoder();
       const data = encoder.encode(verifier);
       const hash = await crypto.subtle.digest('SHA-256', data);
       return this.base64URLEncode(new Uint8Array(hash));
     }

     /**
      * Generate random state parameter for CSRF protection
      * Minimum 32 bytes of entropy
      */
     generateState(): string {
       const array = new Uint8Array(32);
       crypto.getRandomValues(array);
       return this.base64URLEncode(array);
     }

     /**
      * Base64URL encoding (without padding)
      * RFC 4648 Section 5
      */
     private base64URLEncode(buffer: Uint8Array): string {
       let binary = '';
       for (let i = 0; i < buffer.length; i++) {
         binary += String.fromCharCode(buffer[i]);
       }
       return btoa(binary)
         .replace(/\+/g, '-')
         .replace(/\//g, '_')
         .replace(/=/g, '');
     }
   }

   export const pkceService = new PKCEService();
   ```

4. **Token Validation Service**
   ```typescript
   // src/services/tokenValidation.ts

   import { TokenData } from './tokenStorage';

   export class TokenValidationService {
     /**
      * Validate token structure and claims
      */
     validateToken(token: TokenData): { valid: boolean; error?: string } {
       // Check required fields
       if (!token.accessToken || !token.tokenType) {
         return { valid: false, error: 'Missing required token fields' };
       }

       // Check token type
       if (token.tokenType !== 'Bearer') {
         return { valid: false, error: 'Invalid token type' };
       }

       // Check expiration
       if (token.expiresAt <= Date.now()) {
         return { valid: false, error: 'Token expired' };
       }

       // Check issuance time
       if (token.issuedAt > Date.now()) {
         return { valid: false, error: 'Token issued in future' };
       }

       // If JWT, validate structure
       if (this.isJWT(token.accessToken)) {
         const jwtValidation = this.validateJWT(token.accessToken);
         if (!jwtValidation.valid) {
           return jwtValidation;
         }
       }

       return { valid: true };
     }

     /**
      * Check if token is JWT format
      */
     private isJWT(token: string): boolean {
       return token.split('.').length === 3;
     }

     /**
      * Validate JWT structure and basic claims
      */
     private validateJWT(token: string): { valid: boolean; error?: string } {
       try {
         const parts = token.split('.');
         if (parts.length !== 3) {
           return { valid: false, error: 'Invalid JWT structure' };
         }

         // Decode payload (don't verify signature - requires public key)
         const payload = JSON.parse(atob(parts[1]));

         // Check expiration claim
         if (payload.exp && payload.exp * 1000 <= Date.now()) {
           return { valid: false, error: 'JWT expired' };
         }

         // Check not-before claim
         if (payload.nbf && payload.nbf * 1000 > Date.now()) {
           return { valid: false, error: 'JWT not yet valid' };
         }

         // Check issued-at claim
         if (payload.iat && payload.iat * 1000 > Date.now()) {
           return { valid: false, error: 'JWT issued in future' };
         }

         return { valid: true };
       } catch (error) {
         return { valid: false, error: 'Failed to parse JWT' };
       }
     }
   }

   export const tokenValidationService = new TokenValidationService();
   ```

### Environment Variables

```env
# .env

# OAuth Configuration
OAUTH_CLIENT_ID=your-client-id
OAUTH_AUTHORITY=https://login.microsoftonline.com/common
OAUTH_REDIRECT_URI=https://localhost:3000/auth-callback
OAUTH_SCOPES=User.Read,Files.Read

# Token Configuration
TOKEN_EXPIRY_SECONDS=3600
REFRESH_TOKEN_EXPIRY_DAYS=30
TOKEN_REFRESH_BUFFER_MINUTES=5

# Security
ENABLE_TOKEN_ENCRYPTION=true
USE_PKCE=true
```

## OWASP Top 10 Coverage

### A02:2021 - Cryptographic Failures
- **Mitigation**: Encrypt tokens at rest using AES-GCM 256-bit
- **Mitigation**: Use HTTPS for all token transmission
- **Mitigation**: Use Web Crypto API (not deprecated crypto libraries)

### A05:2021 - Security Misconfiguration
- **Mitigation**: Never store tokens in LocalStorage (XSS vulnerability)
- **Mitigation**: Use secure OAuth 2.0 configuration (PKCE, state parameter)
- **Mitigation**: Validate all security configurations at startup

### A07:2021 - Identification and Authentication Failures
- **Mitigation**: Implement automatic token refresh
- **Mitigation**: Use refresh tokens for long-lived sessions
- **Mitigation**: Revoke tokens on logout
- **Mitigation**: Validate token expiration and signature

### A09:2021 - Security Logging and Monitoring Failures
- **Mitigation**: Log token issuance and refresh events (without token values)
- **Mitigation**: Monitor token validation failures
- **Mitigation**: Alert on suspicious token activity

## OAuth 2.0 Security Best Practices

1. **RFC 6749 - OAuth 2.0 Authorization Framework**
   - Use Authorization Code Flow (not Implicit Flow)
   - Include state parameter for CSRF protection
   - Validate redirect URI strictly

2. **RFC 8252 - OAuth 2.0 for Native Apps**
   - Use PKCE (Proof Key for Code Exchange)
   - Generate cryptographically random code verifier
   - Use SHA-256 for code challenge

3. **RFC 6819 - OAuth 2.0 Threat Model**
   - Protect against token leakage
   - Implement token binding (if supported)
   - Use short-lived access tokens

4. **OAuth 2.0 Security Best Current Practice**
   - Always use PKCE (even for confidential clients)
   - Rotate refresh tokens on each use
   - Implement token revocation

## Testing Strategy

1. **Unit Tests**
   - Test token encryption/decryption
   - Test token validation logic
   - Test PKCE code generation
   - Test token expiration checks

2. **Integration Tests**
   - Test token storage and retrieval
   - Test token refresh flow
   - Test token cleanup on logout
   - Test OAuth PKCE flow end-to-end

3. **Security Tests**
   - Verify tokens never appear in console logs
   - Verify tokens not stored in LocalStorage
   - Verify tokens encrypted at rest
   - Verify PKCE parameters generated correctly
   - Verify state parameter validated
   - Test token reuse prevention

4. **Performance Tests**
   - Token encryption/decryption performance (<10ms)
   - Token refresh latency (<500ms)
   - Storage operations performance

## Definition of Done

- [ ] Token storage service implemented with encryption
- [ ] Office.js Roaming Settings integration working
- [ ] Session storage caching implemented
- [ ] Token refresh service with automatic scheduling
- [ ] PKCE implementation (code verifier, challenge, state)
- [ ] Token validation service with JWT support
- [ ] Secure token cleanup on logout
- [ ] No tokens in console logs or LocalStorage
- [ ] Web Crypto API used for all cryptographic operations
- [ ] Token refresh 5 minutes before expiration
- [ ] OAuth 2.0 Authorization Code Flow with PKCE implemented
- [ ] State parameter CSRF protection implemented
- [ ] Unit tests for all token operations (>90% coverage)
- [ ] Security tests passing (no token leakage)
- [ ] Documentation for token management
- [ ] Code review by security expert
- [ ] Penetration testing for token security

## Dependencies

- Office.js API (Roaming Settings)
- Web Crypto API (browser native)
- Azure MSAL library (optional, for Microsoft OAuth)

## Security References

- [RFC 6749 - OAuth 2.0 Authorization Framework](https://datatracker.ietf.org/doc/html/rfc6749)
- [RFC 8252 - OAuth 2.0 for Native Apps](https://datatracker.ietf.org/doc/html/rfc8252)
- [RFC 7636 - PKCE for OAuth 2.0](https://datatracker.ietf.org/doc/html/rfc7636)
- [RFC 6819 - OAuth 2.0 Threat Model](https://datatracker.ietf.org/doc/html/rfc6819)
- [OWASP Token Storage Cheat Sheet](https://cheatsheetseries.owasp.org/cheatsheets/HTML5_Security_Cheat_Sheet.html#local-storage)
- [Web Crypto API](https://developer.mozilla.org/en-US/docs/Web/API/Web_Crypto_API)

## Notes

1. **Production Considerations**
   - Consider using Azure Key Vault for key management
   - Implement token binding (if supported by OAuth provider)
   - Add hardware security module (HSM) support for enterprise
   - Implement certificate pinning for API calls

2. **Compliance Requirements**
   - GDPR: Ensure tokens can be deleted on user request
   - SOC 2: Implement audit logging for token access
   - HIPAA: Use FIPS 140-2 validated cryptographic modules (if required)

3. **Future Enhancements**
   - Implement token binding (RFC 8471)
   - Add biometric authentication for token access
   - Implement certificate-based authentication
   - Add support for multiple concurrent tokens (multi-workspace)
