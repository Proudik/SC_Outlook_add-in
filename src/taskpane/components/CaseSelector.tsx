import * as React from "react";
import { CaseOption } from "../../services/singlecase";
import { AddinSettings } from "./SettingsModal";
import type { CaseSuggestion } from "../../utils/caseSuggestionEngine";
import "./CaseSelector.css";

type Props = {
  title?: string;

  scope: AddinSettings["caseListScope"];
  onScopeChange: (scope: AddinSettings["caseListScope"]) => void;

  selectedCaseId: string;
  onSelectCaseId: (caseId: string) => void;

  cases: CaseOption[];
  isLoadingCases: boolean;

  clientNamesById?: Record<string, string>;

  suggestions?: CaseSuggestion[];

  // When set, the matching suggestion card gets a green "recommended" outline
  // instead of the blue selected highlight. Driven by suggestedCaseId state in
  // MainWorkspace when rememberLastCase is OFF.
  suggestedCaseId?: string;

  // Content-based suggestions (triggered when user clicks "Vybrat jiný spis")
  contentSuggestions?: CaseSuggestion[];
  isLoadingContentSuggestions?: boolean;
};

type Row =
  | { kind: "client"; clientLabel: string; count: number }
  | { kind: "case"; clientLabel: string; caseId: string; label: string; clientId?: string };

function getCaseName(c: any): string {
  return String(c?.title || c?.name || "");
}

function getCaseVisibleId(c: any): string {
  return String(c?.case_id_visible || c?.caseIdVisible || "");
}

function buildCaseLabel(c: any): string {
  const visible = getCaseVisibleId(c);
  const name = getCaseName(c);
  return visible ? `${visible} · ${name}` : name;
}

function getClientKey(c: any): { key: string; label: string; id?: string } {
  const directName =
    (typeof c?.client === "string" && c.client.trim()) ||
    (typeof c?.client_name === "string" && c.client_name.trim()) ||
    (typeof c?.clientName === "string" && c.clientName.trim()) ||
    (typeof c?.client_title === "string" && c.client_title.trim());

  const nestedName =
    (typeof c?.client?.name === "string" && c.client.name.trim()) ||
    (typeof c?.client?.title === "string" && c.client.title.trim());

  const name = directName || nestedName;

  const idRaw =
    c?.client_id ?? c?.clientId ?? c?.client?.id ?? c?.client?.client_id ?? c?.client?.clientId;

  const id = idRaw !== undefined && idRaw !== null ? String(idRaw) : "";

  if (name) return { key: `name:${name}`, label: name, id: id || undefined };
  if (id) return { key: `id:${id}`, label: `Client ${id}`, id };
  return { key: "other", label: "Other" };
}

function buildClientLabel(
  key: { key: string; label: string; id?: string },
  map?: Record<string, string>
): string {
  if (key.key.startsWith("name:")) return key.label;
  if (key.id && map?.[key.id]) return map[key.id];
  return key.label;
}

function buildRows(cases: CaseOption[], clientNamesById?: Record<string, string>): Row[] {
  const groups = new Map<
    string,
    { key: string; label: string; id?: string; items: CaseOption[] }
  >();

  for (const c of cases) {
    const anyC = c as any;
    const ck = getClientKey(anyC);

    const existing = groups.get(ck.key);
    if (existing) existing.items.push(c);
    else groups.set(ck.key, { key: ck.key, label: ck.label, id: ck.id, items: [c] });
  }

  const sortedKeys = Array.from(groups.keys()).sort((a, b) => {
    if (a === "other") return 1;
    if (b === "other") return -1;
    return a.localeCompare(b, "cs");
  });

  const rows: Row[] = [];

  for (const key of sortedKeys) {
    const group = groups.get(key);
    if (!group) continue;

    const clientLabel = buildClientLabel(
      { key: group.key, label: group.label, id: group.id },
      clientNamesById
    );

    const caseRows: Row[] = group.items
      .map((c) => {
        const anyC = c as any;
        const caseId = String(anyC.id || "");
        const label = buildCaseLabel(anyC);

        return {
          kind: "case",
          clientLabel,
          clientId: group.id,
          caseId,
          label,
        } as Row;
      })
      .sort((a, b) => {
        if (a.kind !== "case" || b.kind !== "case") return 0;
        return a.label.localeCompare(b.label, "cs", { sensitivity: "base" });
      });

    if (caseRows.length === 0) continue;

    rows.push({ kind: "client", clientLabel, count: caseRows.length });
    rows.push(...caseRows);
  }

  return rows;
}

function confidencePillClassName(pct: number): string {
  if (pct >= 90) return "case-selector-confidence-pill case-selector-confidence-pill--high";
  if (pct >= 80) return "case-selector-confidence-pill case-selector-confidence-pill--mid";
  if (pct >= 70) return "case-selector-confidence-pill case-selector-confidence-pill--low";
  return "case-selector-confidence-pill case-selector-confidence-pill--none";
}

function chipClassName(active: boolean): string {
  return ["case-selector-chip", active ? "case-selector-chip--active" : ""]
    .filter(Boolean)
    .join(" ");
}

function safeLower(s: string): string {
  return String(s || "")
    .toLowerCase()
    .trim();
}

export default function CaseSelector({
  title = "Case",
  scope,
  onScopeChange,
  selectedCaseId,
  onSelectCaseId,
  cases,
  isLoadingCases,
  clientNamesById,
  suggestions,
  suggestedCaseId,
  contentSuggestions,
  isLoadingContentSuggestions,
}: Props) {
  const [query, setQuery] = React.useState("");
  const inputRef = React.useRef<HTMLInputElement>(null);

  const rows = React.useMemo(() => buildRows(cases, clientNamesById), [cases, clientNamesById]);
  const allCaseRows = React.useMemo(
    () => rows.filter((r) => r.kind === "case") as Array<Extract<Row, { kind: "case" }>>,
    [rows]
  );

  const suggestionRows = React.useMemo(() => {
    const s = (suggestions || []).slice(0, 10);
    if (!s.length) return [];

    return s
      .map((sug) => {
        const found = allCaseRows.find((r) => r.caseId === sug.caseId);
        if (!found) return null;

        return {
          caseId: sug.caseId,
          confidencePct: typeof sug.confidencePct === "number" ? sug.confidencePct : 0,
          reasons: sug.reasons || [],
          label: found.label,
          clientLabel: found.clientLabel,
        };
      })
      .filter(Boolean) as Array<{
      caseId: string;
      confidencePct: number;
      reasons: string[];
      label: string;
      clientLabel: string;
    }>;
  }, [suggestions, allCaseRows]);

  // Process content-based suggestions (triggered by "Vybrat jiný spis")
  const contentSuggestionRows = React.useMemo(() => {
    const s = (contentSuggestions || []).slice(0, 10);
    if (!s.length) return [];

    return s
      .map((sug) => {
        const found = allCaseRows.find((r) => r.caseId === sug.caseId);
        if (!found) return null;

        return {
          caseId: sug.caseId,
          confidencePct: typeof sug.confidencePct === "number" ? sug.confidencePct : 0,
          reasons: sug.reasons || [],
          label: found.label,
          clientLabel: found.clientLabel,
        };
      })
      .filter(Boolean) as Array<{
      caseId: string;
      confidencePct: number;
      reasons: string[];
      label: string;
      clientLabel: string;
    }>;
  }, [contentSuggestions, allCaseRows]);

  const topSuggestion = suggestionRows.length ? suggestionRows[0] : null;

  // show a second suggestion under the top one (when it exists and is not garbage)
  const SECONDARY_MIN_PCT = 1; // set to 50/60 if you want stricter
  const secondSuggestion = React.useMemo(() => {
    const s = suggestionRows[1];
    if (!s) return null;
    if (s.caseId === topSuggestion?.caseId) return null;
    if ((s.confidencePct || 0) < SECONDARY_MIN_PCT) return null;
    return s;
  }, [suggestionRows, topSuggestion?.caseId]);

  // Relevance threshold for the "other suggestions" list
  const RELEVANT_PCT = 60;

  const excludedIds = React.useMemo(() => {
    const set = new Set<string>();
    if (topSuggestion?.caseId) set.add(topSuggestion.caseId);
    if (secondSuggestion?.caseId) set.add(secondSuggestion.caseId);
    return set;
  }, [topSuggestion?.caseId, secondSuggestion?.caseId]);

  const otherRelevantSuggestions = React.useMemo(() => {
    if (!topSuggestion) return [];
    return suggestionRows
      .slice(1)
      .filter((s) => !excludedIds.has(s.caseId))
      .filter((s) => s.confidencePct >= RELEVANT_PCT);
  }, [suggestionRows, topSuggestion, excludedIds]);

  const fallbackAllCases = React.useMemo(() => {
    return allCaseRows.filter((c) => !excludedIds.has(c.caseId));
  }, [allCaseRows, excludedIds]);

  // "Select other" list base:
  // If we have relevant extra suggestions, show them first. Otherwise show all cases.
  const expandedListBase =
    otherRelevantSuggestions.length > 0 ? otherRelevantSuggestions : fallbackAllCases;

  const expandedList = React.useMemo(() => {
    const q = safeLower(query);
    // Only show results when user has typed something
    if (!q) return [];

    const filtered = expandedListBase.filter((x: any) => {
      const label = safeLower(String(x.label || ""));
      const clientLabel = safeLower(String(x.clientLabel || ""));
      return label.includes(q) || clientLabel.includes(q);
    });

    // Limit results for performance
    return filtered.slice(0, 20);
  }, [expandedListBase, query]);

  const selectedCaseDetails = React.useMemo(() => {
    if (!selectedCaseId) return null;
    const anyFound = (cases as any[]).find((c) => String((c as any)?.id || "") === selectedCaseId);
    if (!anyFound) return null;

    const clientKey = getClientKey(anyFound);
    const clientLabel = buildClientLabel(clientKey, clientNamesById);

    return {
      caseId: selectedCaseId,
      label: buildCaseLabel(anyFound),
      clientLabel,
    };
  }, [cases, selectedCaseId, clientNamesById]);

  // Demo favourites: take first 5 cases
  const favouriteCases = React.useMemo(() => {
    return allCaseRows.slice(0, 5);
  }, [allCaseRows]);

  const pick = (caseId: string) => {
    onSelectCaseId(caseId);
    setQuery("");
  };

  React.useEffect(() => {
    setQuery("");
  }, [selectedCaseId, scope]);

  // Log content suggestions only when they change (not on every render)
  React.useEffect(() => {
    if (contentSuggestions && contentSuggestions.length > 0) {
      console.log("[CaseSelector] 📄 Content suggestions received:", {
        count: contentSuggestions.length,
        mapped: contentSuggestionRows.length,
        suggestions: contentSuggestions.map((s) => ({
          caseId: s.caseId,
          pct: s.confidencePct,
        })),
      });
    }
  }, [contentSuggestions?.length]); // Only trigger when count changes, not the array itself

  const hasAnyCases = allCaseRows.length > 0;

  // Autofocus the search input when component mounts or cases load
  React.useEffect(() => {
    if (hasAnyCases && inputRef.current) {
      inputRef.current.focus();
    }
  }, [hasAnyCases]);

  const renderSuggestionButton = (
    s: {
      caseId: string;
      confidencePct: number;
      reasons: string[];
      label: string;
      clientLabel: string;
    },
    variant: "primary" | "secondary"
  ) => {
    const isSelected = selectedCaseId === s.caseId;
    // Recommended: system suggests this case but user has NOT explicitly selected it
    const isRecommended = !isSelected && suggestedCaseId === s.caseId;

    return (
      <button
        type="button"
        className={[
          "case-selector-suggested-btn",
          variant === "secondary" ? "case-selector-suggested-btn--secondary" : "",
          isSelected ? "case-selector-suggested-btn--selected" : "",
          isRecommended ? "case-selector-suggested-btn--recommended" : "",
        ]
          .filter(Boolean)
          .join(" ")}
        onClick={() => pick(s.caseId)}
      >
        <div className="case-selector-suggested-topline">
          <div className="case-selector-suggested-primary">
            {isSelected && <span className="case-selector-checkmark">✓ </span>}
            {s.label}
          </div>
          <div className={confidencePillClassName(s.confidencePct)}>{s.confidencePct}%</div>
        </div>

        <div className="case-selector-suggested-secondary">{s.clientLabel}</div>
        {s.reasons?.length ? (
          <div className="case-selector-suggested-reason">{s.reasons[0]}</div>
        ) : null}
        {isSelected && <div className="case-selector-selected-badge">Selected</div>}
      </button>
    );
  };

  return (
    <div className="case-selector-wrap">
      <div className="case-selector-top-row">
        <div className="case-selector-title">{title}</div>

        <div className="case-selector-chips">
          <button
            className={chipClassName(scope === "favourites")}
            onClick={() => onScopeChange("favourites")}
            type="button"
          >
            Favourites
          </button>

          <button
            className={chipClassName(scope === "all")}
            onClick={() => onScopeChange("all")}
            type="button"
          >
            All
          </button>
        </div>
      </div>

      {isLoadingCases ? (
        <div className="case-selector-loading">Loading…</div>
      ) : !hasAnyCases ? (
        <div className="case-selector-empty">No cases available.</div>
      ) : scope === "favourites" ? (
        // FAVOURITES MODE: Show only favourite cases list
        <div className="case-selector-favourites-list">
          {favouriteCases.map((fav) => {
            const isSelected = selectedCaseId === fav.caseId;
            return (
              <button
                key={`fav-${fav.caseId}`}
                type="button"
                className={[
                  "case-selector-other-item",
                  isSelected ? "case-selector-other-item--selected" : "",
                ]
                  .filter(Boolean)
                  .join(" ")}
                onClick={() => pick(fav.caseId)}
              >
                <div className="case-selector-other-left">
                  <div className="case-selector-other-primary">
                    {isSelected && <span className="case-selector-checkmark">✓ </span>}
                    {fav.label}
                  </div>
                  <div className="case-selector-other-secondary">{fav.clientLabel}</div>
                </div>
                {isSelected && <div className="case-selector-selected-badge">Selected</div>}
              </button>
            );
          })}
        </div>
      ) : (
        <div className="case-selector-suggested-card">
          {/* SELECTED CASE - Always show prominently when a case is selected */}
          {selectedCaseDetails && (
            <div className="case-selector-selected-section">

              <div className="case-selector-selected-card">
                <div className="case-selector-selected-card-content">
                  <div className="case-selector-checkmark-icon">✓</div>
                  <div className="case-selector-selected-info">
                    <div className="case-selector-selected-primary">{selectedCaseDetails.label}</div>
                    <div className="case-selector-selected-secondary">{selectedCaseDetails.clientLabel}</div>
                  </div>
                </div>
                <div className="case-selector-selected-badge-large">SELECTED</div>
              </div>
            </div>
          )}
          {/* Loading state for content-based suggestions */}
          {isLoadingContentSuggestions ? (
            <div className="case-selector-loading" style={{ marginBottom: 12 }}>
              Hledám odpovídající spisy…
            </div>
          ) : null}

          {/* Content-based suggestions (triggered by "Vybrat jiný spis") */}
          {!isLoadingContentSuggestions && contentSuggestionRows.length > 0 ? (
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: "#616161", marginBottom: 8 }}>
                📄 Návrhy podle obsahu emailu
              </div>
              {contentSuggestionRows.slice(0, 3).map((s, idx) => (
                <div key={`content-${s.caseId}`} style={{ marginBottom: 4 }}>
                  {renderSuggestionButton(s, idx === 0 ? "primary" : "secondary")}
                </div>
              ))}
            </div>
          ) : null}

          {/* History-based suggestions (show label only if content suggestions exist) */}
          {contentSuggestionRows.length > 0 && topSuggestion ? (
            <div style={{ fontSize: 12, fontWeight: 600, color: "#616161", marginBottom: 8 }}>
              🕒 Návrhy podle historie
            </div>
          ) : null}

          {/* Only show suggestions if they're not already selected (to avoid duplication with Selected Case section) */}
          {topSuggestion && topSuggestion.caseId !== selectedCaseId ? (
            <>
              {renderSuggestionButton(topSuggestion, "primary")}
              {secondSuggestion && secondSuggestion.caseId !== selectedCaseId ? renderSuggestionButton(secondSuggestion, "secondary") : null}
            </>
          ) : !isLoadingContentSuggestions && contentSuggestionRows.length === 0 && !selectedCaseId ? (
            <div className="case-selector-empty" style={{ marginBottom: 8 }}>
              I couldn’t find a good match yet. Try searching for a case below.
            </div>
          ) : null}

          {/* Always show search input */}
          <div className="case-selector-other-list">
            <input
              ref={inputRef}
              className="case-selector-input case-selector-input--search"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Escape") {
                  setQuery("");
                  e.currentTarget.blur();
                } else if (e.key === "Enter" && expandedList.length > 0) {
                  pick(expandedList[0].caseId);
                }
              }}
              placeholder="Search other cases or clients"
            />

            {/* Only show results when user has typed something */}
            {query.trim().length > 0 && (
              <>
                {expandedList.map((x: any) => {
                  const caseId = String(x.caseId || "");
                  const label = String(x.label || "");
                  const clientLabel = String(x.clientLabel || "");
                  const pct = typeof x.confidencePct === "number" ? x.confidencePct : 0;
                  const isSelected = selectedCaseId === caseId;

                  return (
                    <button
                      key={`o-${caseId}`}
                      type="button"
                      className={[
                        "case-selector-other-item",
                        isSelected ? "case-selector-other-item--selected" : "",
                      ]
                        .filter(Boolean)
                        .join(" ")}
                      onClick={() => pick(caseId)}
                    >
                      <div className="case-selector-other-left">
                        <div className="case-selector-other-primary">
                          {isSelected && <span className="case-selector-checkmark">✓ </span>}
                          {label}
                        </div>
                        <div className="case-selector-other-secondary">{clientLabel}</div>
                      </div>

                      {pct > 0 ? (
                        <div className={confidencePillClassName(pct)}>{pct}%</div>
                      ) : isSelected ? (
                        <div className="case-selector-selected-badge">Selected</div>
                      ) : (
                        <div className="case-selector-other-spacer" />
                      )}
                    </button>
                  );
                })}

                {expandedList.length === 0 && (
                  <div className="case-selector-empty">No matches found.</div>
                )}
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
