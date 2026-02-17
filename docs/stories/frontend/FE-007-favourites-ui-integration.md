# FE-007: Favourites UI Integration

**Story ID:** FE-007
**Story Points:** 3
**Epic Link:** User Experience & Polish
**Status:** Ready for Development

## Description

Add favorites/starred cases functionality to the case selector UI, allowing users to mark frequently-used cases as favorites for quick access. Integrate with the backend favorites API, provide visual indicators for favorite cases, and offer a dedicated "Favorites" filter view. Sync favorites across devices and persist user preferences.

This improves productivity for users who repeatedly file emails to the same set of cases.

## Acceptance Criteria

1. **Favorite Toggle Action**
   - Star icon next to each case in case selector/dropdown
   - Click star to toggle favorite status (filled = favorite, outline = not favorite)
   - Star changes immediately (optimistic update)
   - Sync favorite status to backend API
   - Show loading indicator during sync
   - Rollback on API failure with error message

2. **Favorites Filter View**
   - "Favorites" tab/filter in case selector
   - Show only favorited cases when filter active
   - Display count: "Favorites (5)"
   - Empty state: "No favorite cases yet. Star cases to add them here."
   - Switch between "All Cases", "My Cases", "Favorites" views

3. **Visual Indicators**
   - Favorite cases show filled star icon
   - Non-favorite cases show outline star icon
   - Favorite cases listed at top of "All Cases" view (optional)
   - Highlight favorite cases with subtle background color (optional)

4. **Persistence & Sync**
   - Favorites stored in backend (user-specific)
   - Load favorites on app startup
   - Sync favorites when toggled
   - Handle concurrent updates (last write wins)
   - Cache favorites locally for offline access

5. **Quick Access**
   - Recent cases (last 3 filed) shown above favorites
   - "Pin" case to favorites from success toast after filing
   - Keyboard shortcut to toggle favorite (e.g., "F" key)
   - Search still works within favorites view

6. **Error Handling**
   - If favorites API fails to load, show all cases
   - If toggle API fails, revert UI change
   - Show error toast: "Failed to update favorite"
   - Retry once on network failure

## Technical Requirements

### React Components

1. **FavoriteToggle.tsx** (new component)
   ```tsx
   interface FavoriteToggleProps {
     caseId: string;
     isFavorite: boolean;
     onToggle: (caseId: string, isFavorite: boolean) => Promise<void>;
     size?: 'small' | 'medium' | 'large';
   }
   ```
   - Star icon button (filled or outline)
   - Click to toggle favorite status
   - Show loading spinner during API call
   - Disable during loading
   - Use Fluent UI `Star` and `StarFilled` icons

2. **FavoritesList.tsx** (new component)
   ```tsx
   interface FavoritesListProps {
     cases: CaseOption[];
     favorites: Set<string>; // case IDs
     onSelectCase: (caseId: string) => void;
     onToggleFavorite: (caseId: string, isFavorite: boolean) => Promise<void>;
   }
   ```
   - Display list of favorite cases
   - Include `FavoriteToggle` for each case
   - Click case to select it
   - Show empty state if no favorites

3. **Update CaseSelector.tsx**
   - Add filter tabs: "All", "My Cases", "Favorites"
   - Add `FavoriteToggle` icon to each case item
   - Filter cases based on selected tab
   - Sort favorites to top in "All" view (optional)

4. **CaseSelectorTabs.tsx** (new component)
   ```tsx
   interface CaseSelectorTabsProps {
     activeTab: 'all' | 'my' | 'favorites';
     favoritesCount: number;
     onTabChange: (tab: 'all' | 'my' | 'favorites') => void;
   }
   ```
   - Three tabs for case filtering
   - Show count for favorites tab
   - Fluent UI `TabList` component

5. **Update MainWorkspace.tsx**
   - Load favorites on mount
   - Pass favorites state to CaseSelector
   - Handle favorite toggle actions
   - Show "Add to Favorites" in success toast

### Services

1. **services/favorites.ts** (new service)
   ```typescript
   export interface FavoriteCase {
     caseId: string;
     addedAt: string; // ISO timestamp
   }

   export async function getFavoriteCases(
     token: string
   ): Promise<string[]>; // Returns array of case IDs

   export async function addFavoriteCase(
     token: string,
     caseId: string
   ): Promise<void>;

   export async function removeFavoriteCase(
     token: string,
     caseId: string
   ): Promise<void>;

   export async function toggleFavoriteCase(
     token: string,
     caseId: string,
     isFavorite: boolean
   ): Promise<void>;
   ```

2. **Update services/singlecase.ts**
   - Add favorites API methods
   - Endpoints:
     - `GET /publicapi/v1/users/me/favorites/cases`
     - `POST /publicapi/v1/users/me/favorites/cases/{caseId}`
     - `DELETE /publicapi/v1/users/me/favorites/cases/{caseId}`

### Hooks

1. **hooks/useFavorites.ts** (new hook)
   ```typescript
   export function useFavorites(token: string): {
     favorites: Set<string>; // case IDs
     isLoading: boolean;
     error: Error | null;
     toggleFavorite: (caseId: string, isFavorite: boolean) => Promise<void>;
     isFavorite: (caseId: string) => boolean;
     refresh: () => Promise<void>;
   }
   ```
   - Load favorites on mount
   - Provide toggle function with optimistic update
   - Return favorites as Set for O(1) lookup
   - Handle loading and error states

2. **hooks/useCaseFilter.ts** (new hook)
   ```typescript
   export function useCaseFilter(
     cases: CaseOption[],
     favorites: Set<string>,
     filterMode: 'all' | 'my' | 'favorites'
   ): {
     filteredCases: CaseOption[];
     sortedCases: CaseOption[];
   }
   ```
   - Filter cases based on active tab
   - Sort favorites to top in "all" mode
   - Return filtered and sorted case list

### Storage

1. **utils/favoritesCache.ts** (new utility)
   ```typescript
   export function cacheFavorites(favorites: string[]): void;
   export function getCachedFavorites(): string[] | null;
   export function clearFavoritesCache(): void;
   ```
   - Store favorites in localStorage as backup
   - Use when API fails to load
   - Clear on logout

## API Integration Patterns

1. **Load Favorites**
   ```typescript
   // In hooks/useFavorites.ts
   React.useEffect(() => {
     if (!token) return;

     setIsLoading(true);
     getFavoriteCases(token)
       .then(favoriteIds => {
         setFavorites(new Set(favoriteIds));
         cacheFavorites(favoriteIds); // Cache for offline
       })
       .catch(error => {
         console.error('Failed to load favorites', error);
         // Fallback to cached favorites
         const cached = getCachedFavorites();
         if (cached) {
           setFavorites(new Set(cached));
         }
         setError(error);
       })
       .finally(() => setIsLoading(false));
   }, [token]);
   ```

2. **Toggle Favorite with Optimistic Update**
   ```typescript
   const toggleFavorite = async (caseId: string, isFavorite: boolean) => {
     // Optimistic update
     setFavorites(prev => {
       const next = new Set(prev);
       if (isFavorite) {
         next.add(caseId);
       } else {
         next.delete(caseId);
       }
       return next;
     });

     try {
       if (isFavorite) {
         await addFavoriteCase(token, caseId);
       } else {
         await removeFavoriteCase(token, caseId);
       }

       // Update cache
       cacheFavorites(Array.from(favorites));

     } catch (error) {
       // Rollback on error
       setFavorites(prev => {
         const next = new Set(prev);
         if (isFavorite) {
           next.delete(caseId); // Rollback add
         } else {
           next.add(caseId); // Rollback remove
         }
         return next;
       });

       showError('Failed to update favorite');
       throw error;
     }
   };
   ```

3. **Filter Cases by Favorites**
   ```typescript
   // In hooks/useCaseFilter.ts
   export function useCaseFilter(
     cases: CaseOption[],
     favorites: Set<string>,
     filterMode: 'all' | 'my' | 'favorites'
   ) {
     const filteredCases = React.useMemo(() => {
       switch (filterMode) {
         case 'favorites':
           return cases.filter(c => favorites.has(c.id));
         case 'my':
           return cases.filter(c => c.isMine);
         case 'all':
         default:
           return cases;
       }
     }, [cases, favorites, filterMode]);

     const sortedCases = React.useMemo(() => {
       if (filterMode === 'all') {
         // Sort favorites to top
         return [...filteredCases].sort((a, b) => {
           const aFav = favorites.has(a.id);
           const bFav = favorites.has(b.id);
           if (aFav && !bFav) return -1;
           if (!aFav && bFav) return 1;
           return 0;
         });
       }
       return filteredCases;
     }, [filteredCases, favorites, filterMode]);

     return { filteredCases, sortedCases };
   }
   ```

4. **Quick Add to Favorites from Success Toast**
   ```typescript
   // After successful filing
   showSuccess(`Email filed to ${caseName}`, {
     action: {
       label: 'Add to Favorites',
       onClick: async () => {
         await toggleFavorite(caseId, true);
         showInfo('Case added to favorites');
       }
     }
   });
   ```

## Reference Implementation

Review the demo for case selector patterns:

1. **Case Selector UI**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/taskpane/components/CaseSelector.tsx`
   - Current case picker UI
   - **ENHANCE** by adding favorite toggle icons

2. **Case Scope Filtering**: `/Users/Martin/SingleCase/SC_Outlook_addin v2 copy/src/services/singlecase.ts`
   - `listCases(token, scope)` function with "my" scope
   - Add "favorites" scope support

3. **API Service Pattern**: Same file
   - Follow existing API service patterns for favorites endpoints

## Dependencies

- **Requires**: FE-001 (OAuth Integration) - need token for favorites API
- **Requires**: FE-005 (Error Handling) - use toast system for feedback
- **Requires**: Backend API endpoints for favorites CRUD operations
- **Enhances**: Case selection workflow

## Notes

1. **Backend API Requirements**
   - `GET /users/me/favorites/cases` - Returns list of favorite case IDs
   - `POST /users/me/favorites/cases/{caseId}` - Add case to favorites
   - `DELETE /users/me/favorites/cases/{caseId}` - Remove case from favorites
   - All endpoints return 200 OK or 4xx/5xx errors

2. **Performance**
   - Load favorites once on app startup, cache in memory
   - Favorites lookup must be O(1) - use Set, not Array
   - Don't reload entire case list when toggling favorite
   - Debounce favorite toggles (prevent rapid clicking)

3. **User Experience**
   - Star icon should feel responsive (immediate visual feedback)
   - Show success feedback only for explicit "add to favorites" actions
   - Don't interrupt filing workflow with favorite prompts
   - Allow keyboard shortcuts for power users

4. **Testing Strategy**
   - Test toggle favorite (add and remove)
   - Test optimistic update with API failure (rollback)
   - Test favorites filter view
   - Test empty favorites state
   - Test sorting (favorites at top)
   - Test offline mode (use cached favorites)

5. **Future Enhancements**
   - Drag to reorder favorites (custom sort order)
   - Folder-like grouping for favorites
   - Share favorite lists with team members
   - Auto-favorite cases filed more than N times

6. **Accessibility**
   - Star button must have ARIA label: "Add to favorites" / "Remove from favorites"
   - Announce favorite status changes to screen readers
   - Keyboard navigation for favorites list
   - High contrast mode support for star icons

## Definition of Done

- [ ] `FavoriteToggle` component displays star icon
- [ ] Click star to toggle favorite status
- [ ] Favorites synced to backend API
- [ ] Optimistic update with rollback on error
- [ ] "Favorites" filter tab in case selector
- [ ] Favorites count displayed in tab
- [ ] Empty state shown when no favorites
- [ ] Favorites sorted to top in "All Cases" view
- [ ] Favorites cached locally for offline use
- [ ] "Add to Favorites" action in success toast
- [ ] Error handling for favorites API failures
- [ ] Loading indicators during favorites sync
- [ ] Unit tests for favorites service
- [ ] Integration tests for favorites UI
- [ ] Tested with 0, 5, and 50+ favorites
- [ ] Documentation updated with favorites feature
