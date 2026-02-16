# Content-Based Suggestions - Debug & Fix Guide

## Problem
Content-based suggestions are being generated but the UI cards are not visible.

## What I've Done

### 1. Added Comprehensive Console Logging

**File: `src/taskpane/components/MainWorkspace/MainWorkspace.tsx`**
- The handler still has a setTimeout (line 2690-2714) which causes issues
- **YOU NEED TO MANUALLY FIX THIS** (see patch below)

**File: `src/taskpane/components/CaseSelector.tsx`**
- Added detailed logging in `contentSuggestionRows` processing (lines 203-251)
- Added useEffect to log render state (lines 318-327)
- These logs will show:
  - When content suggestions are received
  - How many map to valid cases
  - What's being rendered

### 2. Key Issues to Check

#### Issue #1: setTimeout Causes Stale Closure
The current handler uses `setTimeout(() => { ... }, 300)` which can cause:
- State updates happening after component unmounts
- Stale closure capturing old values
- UI rendering before suggestions are ready

**Solution:** Remove setTimeout, run synchronously

#### Issue #2: Case Mapping Failure
If `contentSuggestionRows` is empty even when `contentSuggestions` has data, it means:
- The suggested cases don't exist in `allCaseRows` (visibleCases)
- This happens if the case filter (favourites vs all) excludes suggested cases

**Solution:** Check console logs for "Case X not found in allCaseRows" warnings

#### Issue #3: UI Not Re-rendering
If state updates but UI doesn't change:
- CaseSelector might not be re-rendering
- Props might not be updating correctly

**Solution:** Check the useEffect logs to confirm re-renders

## Manual Fix Required

### Replace the handler in MainWorkspace.tsx

**Find this code around line 2670-2726:**

```typescript
if (intent === "pick_another_case") {
  // ... current implementation with setTimeout
}
```

**Replace with:**

```typescript
if (intent === "pick_another_case") {
  // Check if we should trigger content-based suggestions
  const wasAutoSelected = selectedSource === "suggested" || selectedSource === "remembered";
  const hasContent = subjectText.trim() || suggestBodySnippet.trim();

  console.log("[pick_another_case] 🔵 Button clicked", {
    wasAutoSelected,
    hasContent,
    selectedSource,
    subjectLength: subjectText.length,
    bodyLength: suggestBodySnippet.length,
    visibleCasesCount: visibleCases.length,
  });

  if (composeMode) {
    setSelectedCaseId("");
    setSelectedSource("manual");

    // Trigger content-based suggestions if case was auto-selected and we have content
    if (wasAutoSelected && hasContent) {
      console.log("[pick_another_case] ✅ Triggering content analysis");

      try {
        const result = suggestCasesByContent({
          subject: subjectText,
          bodySnippet: suggestBodySnippet,
          cases: visibleCases,
          topK: 5,
        });

        console.log("[pick_another_case] 📊 Analysis complete:", {
          foundCount: result.suggestions.length,
          suggestions: result.suggestions.map(s => ({
            caseId: s.caseId,
            pct: s.confidencePct,
            score: s.score,
            reasons: s.reasons,
          })),
        });

        setContentBasedSuggestions(result.suggestions);

        if (result.suggestions.length === 0) {
          console.log("[pick_another_case] ❌ No suggestions found");
          setPrompt({ itemId: activeItemId, kind: "unfiled", text: "Žádné návrhy podle obsahu. Vyberte spis ručně." });
        } else {
          console.log("[pick_another_case] ✅ Set", result.suggestions.length, "suggestions in state");
          setPrompt({ itemId: activeItemId, kind: "unfiled", text: "Vyberte spis podle obsahu emailu." });
        }
      } catch (error) {
        console.error("[pick_another_case] ❌ Analysis failed:", error);
        setContentBasedSuggestions([]);
        setPrompt({ itemId: activeItemId, kind: "unfiled", text: "Vyberte spis." });
      }
    } else {
      console.log("[pick_another_case] ⏭️ Skipping analysis:", {
        reason: !wasAutoSelected ? "not auto-selected" : "no content",
      });
      setContentBasedSuggestions([]);
      setPrompt({ itemId: activeItemId, kind: "unfiled", text: "Vyberte spis." });
    }
  }

  setViewMode("pickCase");
  setPickStep("case");
  setChatStep("compose_choose_case");
  setQuickActions([]);
  return;
}
```

**Key changes:**
- ❌ Removed `setTimeout` - analysis runs synchronously
- ❌ Removed `setIsLoadingContentSuggestions` - instant results
- ✅ Added comprehensive logging at every step
- ✅ Suggestions update immediately before `setViewMode("pickCase")`

## Testing Steps

1. **Open Console** (Safari Web Inspector or Browser DevTools)

2. **Test Scenario:**
   - Compose email to someone you've emailed before
   - Wait for auto-suggestion to appear (e.g., "Contract ABC")
   - Write subject/body with different keywords (e.g., "Invoice for Project XYZ")
   - Click "Vybrat jiný spis"

3. **Check Console Logs:**

**Expected log sequence:**
```
[pick_another_case] 🔵 Button clicked {wasAutoSelected: true, hasContent: true, ...}
[pick_another_case] ✅ Triggering content analysis
[pick_another_case] 📊 Analysis complete: {foundCount: 3, suggestions: [...]}
[pick_another_case] ✅ Set 3 suggestions in state
[CaseSelector] 🔍 Processing content suggestions: {contentSuggestionsCount: 3, ...}
[CaseSelector] ✅ Mapped suggestion 0: {caseId: "...", label: "...", pct: 85}
[CaseSelector] ✅ Mapped suggestion 1: ...
[CaseSelector] ✅ Mapped suggestion 2: ...
[CaseSelector] 📋 Content suggestion rows ready: 3
[CaseSelector] 🎨 Render state: {contentSuggestionRowsCount: 3, ...}
```

4. **Check UI:**
   - Should see heading: "📄 Návrhy podle obsahu emailu"
   - Should see 1-3 blue suggestion cards with confidence badges
   - Each card should have:
     - Case title
     - Client name
     - Reason text (e.g., "Subject matches case name")
     - Confidence % badge

## Troubleshooting

### Problem: Console shows "foundCount: 0"
**Cause:** No cases match the email content
**Check:**
- Is subject/body generic?
- Do case titles contain any of the keywords?
**Fix:** Try with more specific content

### Problem: Console shows "Case X not found in allCaseRows"
**Cause:** Suggested case filtered out by scope (favourites vs all)
**Check:** Are you on "Favourites" tab but suggestion is in "All"?
**Fix:** Switch to "All" cases tab

### Problem: Logs show suggestions but no UI cards
**Cause:** Render logic issue or CSS hiding cards
**Check:**
- Look for "[CaseSelector] 📋 Content suggestion rows ready: X" where X > 0
- Look for "[CaseSelector] 🎨 Render state" logs
**Fix:** If logs show data but no rendering, check CSS or React DevTools

### Problem: "No suggestions yet" always shows
**Cause:** Either no suggestions OR topSuggestion exists (history) hiding content
**Check:** Console for contentSuggestionRowsCount
**Fix:** Verify content suggestions are actually in state

## Next Steps

1. **Apply the manual fix above** (remove setTimeout)
2. **Rebuild:** `npm run build`
3. **Test** with console open
4. **Share console logs** if still not working

The logs will tell us exactly where the flow breaks!
