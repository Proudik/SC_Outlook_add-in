# FE-010: Offline Mode & Sync Strategy

**Story ID:** FE-010
**Story Points:** 2
**Epic Link:** User Experience & Polish
**Status:** Ready for Development

## Description

Implement basic offline mode functionality and data caching to improve user experience when the add-in loses internet connectivity. Cache case lists, favorites, and user preferences locally so users can browse previously loaded data offline. Show clear offline indicators and gracefully handle offline scenarios without showing confusing errors.

This is a foundational story - advanced features like offline email queuing and background sync are out of scope.

## Acceptance Criteria

1. **Offline Detection**
   - Detect when add-in loses internet connection
   - Use `navigator.onLine` API and `online`/`offline` events
   - Verify connectivity with periodic ping to proxy health endpoint
   - Show persistent banner: "You're offline" at top of add-in

2. **Data Caching**
   - Cache case list in localStorage (last fetched list)
   - Cache favorites in localStorage
   - Cache user preferences (settings)
   - Cache workspace info (workspace ID, name, host)
   - Set cache expiration (24 hours)
   - Clear cache on logout

3. **Offline UI Behavior**
   - Show cached cases when offline
   - Display "Last updated: 2 hours ago" timestamp
   - Disable filing button when offline (with explanation tooltip)
   - Disable suggestions when offline (show cached suggestions if available)
   - Allow viewing cached data (read-only mode)

4. **Graceful Degradation**
   - API failures due to offline: Show friendly "You're offline" message (not generic errors)
   - Don't spam error toasts when offline
   - Queue actions that require internet (out of scope for this story, just disable for now)
   - Automatically retry failed requests when coming back online

5. **Online Restoration**
   - Detect when connection restored
   - Hide offline banner
   - Automatically refresh stale cached data
   - Re-enable disabled features
   - Show success toast: "Back online"

6. **Cache Management**
   - Add "Refresh" button to manually refresh cached data
   - Clear cache on logout
   - Clear cache if corrupted or invalid
   - Limit cache size (max 5MB)

## Technical Requirements

### Services

1. **services/offlineManager.ts** (new service)
   ```typescript
   export interface OfflineState {
     isOnline: boolean;
     lastOnlineAt: Date | null;
     isConnectivityCheckPending: boolean;
   }

   export function getCurrentOfflineState(): OfflineState;

   export function subscribeToConnectivityChanges(
     callback: (state: OfflineState) => void
   ): () => void; // Returns unsubscribe function

   export async function checkConnectivity(): Promise<boolean>;

   export function scheduleConnectivityCheck(intervalMs?: number): void;

   export function clearConnectivityCheck(): void;
   ```

2. **services/cacheManager.ts** (new service)
   ```typescript
   export interface CacheEntry<T> {
     data: T;
     cachedAt: number; // timestamp
     expiresAt: number; // timestamp
   }

   export function setCache<T>(
     key: string,
     data: T,
     expirationMs?: number
   ): void;

   export function getCache<T>(key: string): T | null;

   export function isCacheValid(key: string): boolean;

   export function clearCache(key?: string): void; // Clear specific key or all

   export function getCacheTimestamp(key: string): Date | null;
   ```

3. **Update services/singlecase.ts**
   - Wrap API calls with offline detection
   - Return cached data if offline
   - Add cache-first strategy for read operations

### Hooks

1. **hooks/useOfflineMode.ts** (new hook)
   ```typescript
   export function useOfflineMode(): {
     isOnline: boolean;
     isOffline: boolean;
     lastOnlineAt: Date | null;
     checkConnectivity: () => Promise<boolean>;
   }
   ```

2. **hooks/useCachedData.ts** (new hook)
   ```typescript
   export function useCachedData<T>(
     cacheKey: string,
     fetchFn: () => Promise<T>,
     options?: {
       cacheExpirationMs?: number;
       refetchOnMount?: boolean;
       refetchOnOnline?: boolean;
     }
   ): {
     data: T | null;
     isLoading: boolean;
     isCached: boolean;
     cachedAt: Date | null;
     error: Error | null;
     refetch: () => Promise<void>;
   }
   ```

3. **Update hooks/useNetworkStatus.ts** (from FE-005)
   - Enhance with offline detection logic
   - Add connectivity checking
   - Add online restoration detection

### React Components

1. **Update OfflineBanner.tsx** (from FE-005)
   - Show "You're offline" message
   - Show "Last online: 5 minutes ago"
   - Show "Retry" button to check connectivity
   - Auto-hide when connection restored

2. **CacheStatusIndicator.tsx** (new component)
   ```tsx
   interface CacheStatusIndicatorProps {
     cachedAt: Date | null;
     onRefresh?: () => void;
   }
   ```
   - Show "Last updated: X minutes ago"
   - Show "Refresh" button
   - Use relative time formatting

3. **Update MainWorkspace.tsx**
   - Use cached cases when offline
   - Show cache status indicator
   - Disable filing when offline
   - Show "Offline - can't file emails" tooltip on disabled file button

### Storage Keys

1. **Cache Keys (localStorage)**
   ```typescript
   const CACHE_KEYS = {
     CASES: 'sc_cache_cases',
     FAVORITES: 'sc_cache_favorites',
     SUGGESTIONS: 'sc_cache_suggestions',
     USER_PREFERENCES: 'sc_cache_preferences',
     WORKSPACE_INFO: 'sc_cache_workspace',
   };
   ```

2. **Cache Expiration**
   - Cases: 1 hour
   - Favorites: 24 hours
   - User preferences: Never (until changed)
   - Workspace info: Never (until logout)

## API Integration Patterns

1. **Cache-First API Call**
   ```typescript
   // In services/singlecase.ts
   export async function listCases(
     token: string,
     scope: CaseScope = 'my'
   ): Promise<CaseOption[]> {
     const cacheKey = `${CACHE_KEYS.CASES}_${scope}`;

     // Try cache first if offline
     if (!navigator.onLine) {
       const cached = getCache<CaseOption[]>(cacheKey);
       if (cached) {
         console.log('[Offline] Returning cached cases');
         return cached;
       }
       throw new Error('No cached cases available offline');
     }

     // Fetch from API
     try {
       const cases = await scRequest<CaseOption[]>('GET', '/cases', token, { scope });

       // Update cache
       setCache(cacheKey, cases, 60 * 60 * 1000); // 1 hour

       return cases;
     } catch (error) {
       // Fallback to cache on network error
       if (isNetworkError(error)) {
         const cached = getCache<CaseOption[]>(cacheKey);
         if (cached) {
           console.log('[Network Error] Returning cached cases');
           return cached;
         }
       }
       throw error;
     }
   }
   ```

2. **Offline Detection with Ping**
   ```typescript
   // In services/offlineManager.ts
   export async function checkConnectivity(): Promise<boolean> {
     // First check navigator.onLine (fast, but not reliable)
     if (!navigator.onLine) {
       return false;
     }

     // Verify with actual network request (reliable)
     try {
       const controller = new AbortController();
       const timeoutId = setTimeout(() => controller.abort(), 3000); // 3s timeout

       const response = await fetch('/singlecase/health', {
         method: 'HEAD',
         signal: controller.signal,
         cache: 'no-store',
       });

       clearTimeout(timeoutId);
       return response.ok;
     } catch {
       return false;
     }
   }
   ```

3. **Auto-Refresh on Online**
   ```typescript
   // In MainWorkspace.tsx
   const { isOnline } = useOfflineMode();
   const { data: cases, refetch } = useCachedData(
     CACHE_KEYS.CASES,
     () => listCases(token),
     { refetchOnOnline: true }
   );

   React.useEffect(() => {
     if (isOnline) {
       // Coming back online - refresh data
       refetch().then(() => {
         showInfo('Data refreshed');
       });
     }
   }, [isOnline]);
   ```

4. **Disable Features When Offline**
   ```typescript
   const { isOffline } = useOfflineMode();

   return (
     <Tooltip
       content={isOffline ? "Can't file emails while offline" : "File email to case"}
     >
       <Button
         onClick={handleFileEmail}
         disabled={isOffline || isLoading}
       >
         File Email
       </Button>
     </Tooltip>
   );
   ```

## Reference Implementation

Review the demo for patterns:

1. **Storage Usage**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/storage.ts`
   - Current storage utilities
   - Enhance with cache management

2. **Network Detection**: Check if demo has any offline handling
   - Likely none, so build from scratch

3. **Settings Storage**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/utils/settingsStorage.ts`
   - Shows how preferences are stored
   - Reuse pattern for cache storage

## Dependencies

- **Requires**: FE-005 (Error Handling) - offline banner component
- **Enhances**: All data-fetching stories (FE-001 through FE-004)
- **Future**: Offline email queuing (out of scope for this story)

## Notes

1. **Out of Scope**
   - Offline email filing queue (defer to future story)
   - Background sync (defer to future story)
   - Service workers (not needed for basic caching)
   - IndexedDB (localStorage is sufficient for now)

2. **Browser Compatibility**
   - `navigator.onLine` is widely supported
   - `online`/`offline` events are reliable in modern browsers
   - localStorage is available in all Office.js contexts

3. **Cache Size Limits**
   - localStorage limit: ~5-10MB per origin
   - Monitor cache size, clear old entries if approaching limit
   - Prioritize important data (workspace info, favorites > suggestions)

4. **Testing Strategy**
   - Test with browser offline mode (DevTools)
   - Test with slow 3G throttling
   - Test offline → online transition
   - Test cache expiration (mock time)
   - Test cache corruption handling
   - Test cache size limits

5. **User Experience**
   - Make offline mode obvious (persistent banner)
   - Don't confuse users with "network error" messages when offline
   - Allow browsing cached data (better than blocking everything)
   - Show cache age to manage expectations

6. **Performance**
   - Cache reads should be <10ms (localStorage is fast)
   - Don't block UI with connectivity checks
   - Use background connectivity polling (every 30s)
   - Debounce online/offline events (prevent flapping)

## Definition of Done

- [ ] Offline detection using `navigator.onLine` and connectivity ping
- [ ] Persistent offline banner shown when offline
- [ ] Cases cached in localStorage with 1-hour expiration
- [ ] Favorites cached in localStorage with 24-hour expiration
- [ ] User preferences cached permanently
- [ ] Cached data displayed when offline
- [ ] "Last updated" timestamp shown for cached data
- [ ] Filing button disabled when offline (with tooltip)
- [ ] Suggestions disabled when offline
- [ ] Auto-refresh when coming back online
- [ ] "Back online" toast shown on restoration
- [ ] "Refresh" button manually refreshes cached data
- [ ] Cache cleared on logout
- [ ] Cache-first strategy for API calls
- [ ] Network errors fall back to cache
- [ ] Unit tests for cache manager
- [ ] Integration tests for offline mode
- [ ] Tested with browser offline mode
- [ ] Documentation updated with offline behavior
