import * as React from "react";
import type { CaseOption } from "../services/singlecase";
import { suggestCasesLocal } from "../utils/caseSuggestionEngine";
import type { CaseSuggestion } from "../utils/caseSuggestionEngine";

type AttachmentLike = { name: string; isInline?: boolean };

type SelectedSource = "" | "remembered" | "last_case" | "suggested" | "manual";

export function useCaseSuggestions(params: {
  enabled: boolean;
  emailItemId: string;
  conversationKey?: string;
  subject: string;
  bodySnippet: string;
  fromEmail: string;
  attachments: AttachmentLike[];
  cases: CaseOption[];
  selectedCaseId: string;
  selectedSource?: SelectedSource;
  onAutoSelectCaseId?: (caseId: string) => void;
  topK?: number;
}): { suggestions: CaseSuggestion[]; autoSelectCaseId: string } {
  const {
    enabled,
    emailItemId,
    conversationKey,
    subject,
    bodySnippet,
    fromEmail,
    attachments,
    cases,
    selectedCaseId,
    selectedSource,
    onAutoSelectCaseId,
    topK,
  } = params;

  const [suggestions, setSuggestions] = React.useState<CaseSuggestion[]>([]);
  const [autoSelectCaseId, setAutoSelectCaseId] = React.useState("");

  // Keep latest callback without making the effect depend on it
  const onAutoSelectRef = React.useRef<typeof onAutoSelectCaseId>();
  React.useEffect(() => {
    onAutoSelectRef.current = onAutoSelectCaseId;
  }, [onAutoSelectCaseId]);

  // Stable keys so we do not depend on array identity
  const attachmentsKey = React.useMemo(() => {
    return (attachments || [])
      .filter((a) => !a?.isInline)
      .map((a) => String(a?.name || ""))
      .join("|");
  }, [attachments]);

  const casesKey = React.useMemo(() => {
    return (cases || []).map((c: any) => String(c?.id ?? "")).join("|");
  }, [cases]);

  // Prevent infinite auto-select loops
  const lastAutoPickRef = React.useRef<{ emailItemId: string; caseId: string }>({
    emailItemId: "",
    caseId: "",
  });

  React.useEffect(() => {
    if (!enabled || !emailItemId || !cases?.length) {
      setSuggestions([]);
      setAutoSelectCaseId("");
      return;
    }

    const attachmentNames = (attachments || [])
      .filter((a) => !a.isInline)
      .map((a) => a.name);

    const res = suggestCasesLocal({
      conversationKey: conversationKey || "",
      subject,
      bodySnippet,
      attachmentNames,
      fromEmail,
      cases,
      topK: topK ?? 2,
    });

    console.log("[SC suggest] cases:", cases.length);
console.log("[SC suggest] subject:", subject);
console.log("[SC suggest] bodySnippet:", bodySnippet);
console.log("[SC suggest] suggestions:", res.suggestions.map(s => ({ caseId: s.caseId, pct: s.confidencePct, score: s.score })));


const hasIKH = (cases || []).some((c: any) => {
  const t = String(
    c?.title ||
      c?.name ||
      c?.label ||
      c?.caseTitle ||
      c?.case_name_visible ||
      c?.caseNameVisible ||
      ""
  ).toLowerCase();
  return t.includes("internal") && t.includes("know") && t.includes("how");
});

console.log("[SC suggest] has Internal Know How in cases[]:", hasIKH);

// also print all case titles quickly (temporary)
console.log(
  "[SC suggest] case titles sample:",
  (cases || []).slice(0, 15).map((c: any) => ({
    id: String(c?.id ?? ""),
    title: String(c?.title || c?.name || c?.label || c?.caseTitle || ""),
  }))
);

// also print all case titles quickly (temporary)
console.log(
  "[SC suggest] case titles sample:",
  (cases || []).slice(0, 15).map((c: any) => ({
    id: String(c?.id ?? ""),
    title: String(c?.title || c?.name || c?.label || c?.caseTitle || ""),
  }))
);

    // Only update state if it really changed
    setSuggestions((prev) => {
      const prevKey = (prev || []).map((x: any) => `${String(x?.caseId || "")}:${String(x?.score ?? "")}`).join("|");
      const nextKey = (res.suggestions || []).map((x: any) => `${String(x?.caseId || "")}:${String(x?.score ?? "")}`).join("|");
      return prevKey === nextKey ? prev : res.suggestions;
    });

    setAutoSelectCaseId((prev) => (prev === res.autoSelectCaseId ? prev : res.autoSelectCaseId));

    // Never override manual selection
    if (selectedSource === "manual") return;

    const canOverride =
      !selectedCaseId ||
      selectedSource === "remembered" ||
      selectedSource === "suggested" ||
      selectedSource === "" ||
      selectedSource === undefined;

    if (!canOverride) return;

    const nextId = String(res.autoSelectCaseId || "");
    if (!nextId) return;
    if (String(selectedCaseId || "") === nextId) return;

    const already =
      lastAutoPickRef.current.emailItemId === String(emailItemId) &&
      lastAutoPickRef.current.caseId === nextId;

    if (already) return;

    lastAutoPickRef.current = { emailItemId: String(emailItemId), caseId: nextId };
    onAutoSelectRef.current?.(nextId);
  }, [
    enabled,
    emailItemId,
    conversationKey,
    subject,
    fromEmail,
    bodySnippet,
    attachmentsKey,
    casesKey,
    selectedCaseId,
    selectedSource,
    topK,
  ]);

  return { suggestions, autoSelectCaseId };
}