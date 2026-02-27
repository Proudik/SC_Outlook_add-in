import { CaseOption } from "../services/singlecase";
import { getSuggestStats, getThreadMappedCaseId } from "./suggestionStorage";

export type CaseSuggestion = {
  caseId: string;
  score: number;
  confidencePct: number;
  reasons: string[];
};

function safeLower(s: string): string {
  return String(s || "").trim().toLowerCase();
}

function emailDomain(email: string): string {
  const e = safeLower(email);
  const at = e.lastIndexOf("@");
  return at >= 0 ? e.slice(at + 1) : "";
}

function getCaseVisibleId(c: any): string {
  return String(c?.case_id_visible || c?.caseIdVisible || c?.caseIdVisibleText || c?.visibleId || "").trim();
}

// IMPORTANT: prefer caseName (raw name without visible-ID prefix) over the
// combined title field so fuzzy matching compares "Human Ressource" instead of
// "2026-0001 · Human Ressource".
function getCaseTitle(c: any): string {
  return String(
    c?.caseName ||
      c?.name ||
      c?.case_name ||
      c?.caseNameVisible ||
      c?.case_name_visible ||
      c?.title ||
      c?.label ||
      c?.caseTitle ||
      ""
  ).trim();
}

function stripDiacritics(s: string): string {
  try {
    return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  } catch {
    return s;
  }
}

// Keeps hyphens (good for IDs), but we will also create a "loose" version for matching titles.
function normText(s: string): string {
  return safeLower(stripDiacritics(s))
    .replace(/[^a-z0-9\s.:-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Loose matching: treat "-" as whitespace so "know-how" matches "know how".
function normLoose(s: string): string {
  return normText(s).replace(/-/g, " ").replace(/\s+/g, " ").trim();
}

function tokenizeLoose(s: string): string[] {
  return normLoose(s)
    .split(" ")
    .map((t) => t.trim())
    .filter((t) => t.length >= 3);
}

function tokenOverlapScore(aTokens: string[], bText: string): { hits: number; total: number } {
  if (!aTokens.length) return { hits: 0, total: 0 };
  const b = ` ${normLoose(bText)} `;
  let hits = 0;
  for (const t of aTokens) {
    if (b.includes(` ${t} `)) hits += 1;
  }
  return { hits, total: aTokens.length };
}

/**
 * Levenshtein edit distance (character level).
 * O(min(|a|, |b|)) space via single-row DP.
 */
function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  // Keep `a` as the shorter string so the row fits in less memory
  if (a.length > b.length) { const t = a; a = b; b = t; }

  const m = a.length;
  const n = b.length;
  const row: number[] = Array.from({ length: m + 1 }, (_, i) => i);

  for (let j = 1; j <= n; j++) {
    let prev = row[0];
    row[0] = j;
    for (let i = 1; i <= m; i++) {
      const tmp = row[i];
      row[i] = a[i - 1] === b[j - 1]
        ? prev
        : 1 + Math.min(prev, row[i], row[i - 1]);
      prev = tmp;
    }
  }
  return row[m];
}

/**
 * Normalized Levenshtein similarity in [0, 1].
 *   sim = 1 − levenshtein(a, b) / max(|a|, |b|)
 */
function strSimilarity(a: string, b: string): number {
  if (a === b) return 1;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - levenshtein(a, b) / maxLen;
}

/**
 * Typo-tolerant similarity between a normalized subject and a normalized case title.
 * Takes the best of:
 *   - full-string similarity (good when lengths are similar, e.g. "human recources" vs "human ressource")
 *   - prefix-truncated similarity (good when one string is much longer, e.g. subject is shorter
 *     than a longer case title: "Human Resources" vs "Human Resources Management")
 *
 * Returns 0 when either string is shorter than 5 chars (too short for reliable fuzzy matching).
 */
function subjectTitleSimilarity(subjectLoose: string, titleLoose: string): number {
  if (!subjectLoose || !titleLoose) return 0;
  if (subjectLoose.length < 5 || titleLoose.length < 5) return 0;

  const full = strSimilarity(subjectLoose, titleLoose);

  // Also compare both strings truncated to the length of the shorter one so that
  // "human recources" (15) vs "human ressource management" (26) is not penalised
  // for length difference alone.
  const minLen = Math.min(subjectLoose.length, titleLoose.length);
  const prefix = strSimilarity(
    subjectLoose.slice(0, minLen),
    titleLoose.slice(0, minLen),
  );

  return Math.max(full, prefix);
}

function clamp(n: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, n));
}

function log1p(x: number): number {
  return Math.log(1 + Math.max(0, x));
}

function confidencePctFor(sortedScores: number[], idx: number): number {
  const score = sortedScores[idx] ?? 0;
  const top = sortedScores[0] ?? 0;
  const second = sortedScores[1] ?? 0;

  const base = clamp(score / 120, 0, 1);

  const gapRef = idx === 0 ? second : top;
  const gap = Math.max(0, score - gapRef);
  const sep = clamp(gap / 60, 0, 1);

  const pct = Math.round(100 * (0.65 * base + 0.35 * sep));
  return clamp(pct, 0, 100);
}

export function suggestCasesLocal(params: {
  conversationKey?: string;
  subject?: string;
  bodySnippet?: string;
  attachmentNames?: string[];
  fromEmail?: string;
  cases: CaseOption[];
  topK?: number;
}): { suggestions: CaseSuggestion[]; autoSelectCaseId: string } {
  const topK = params.topK ?? 3;

  const conversationKey = String(params.conversationKey || "").trim();

  const subjectRaw = String(params.subject || "");
  const subjectLoose = normLoose(subjectRaw);

  const bodyRaw = String(params.bodySnippet || "");
  const bodyLoose = normLoose(bodyRaw);

  const attachmentNames = (params.attachmentNames || []).map((x) => normText(x));
  const fromEmail = safeLower(params.fromEmail || "");
  const domain = emailDomain(fromEmail);

  const { senderToCase, domainToCase, recentCases } = getSuggestStats();

  const reasonsByCase: Record<string, string[]> = {};
  const scoreByCase: Record<string, number> = {};

  const add = (caseId: string, delta: number, reason?: string) => {
    if (!caseId) return;
    scoreByCase[caseId] = (scoreByCase[caseId] || 0) + delta;
    if (reason) {
      if (!reasonsByCase[caseId]) reasonsByCase[caseId] = [];
      if (!reasonsByCase[caseId].includes(reason)) reasonsByCase[caseId].push(reason);
    }
  };

  // 1) Thread mapping (strongest)
  if (conversationKey) {
    const mapped = getThreadMappedCaseId(conversationKey);
    if (mapped) add(mapped, 100, "Same email thread previously attached to this case.");
  }

  // 2) Visible case reference match (very strong)
  // Use the stricter normText here to avoid breaking IDs like 2023-0006.
  const subjectStrict = normText(subjectRaw);
  const bodyStrict = normText(bodyRaw);

  for (const c of params.cases) {
    const anyC = c as any;
    const caseId = String(anyC?.id || "");
    const vis = normText(getCaseVisibleId(anyC));
    if (!caseId || !vis) continue;

    const inSubject = subjectStrict.includes(vis);
    const inBody = bodyStrict.includes(vis);
    const inAtt = attachmentNames.some((n) => n.includes(vis));

    if (inSubject || inBody || inAtt) {
      add(caseId, 95, "Case reference found in the email.");
    }
  }

  // 3) Case title matching
  // Subject should dominate. Body can be suggestion 2 but must be stricter to avoid random matches.
  for (const c of params.cases) {
    const anyC = c as any;
    const caseId = String(anyC?.id || "");
    if (!caseId) continue;

    const titleRaw = getCaseTitle(anyC);
    const titleLoose = normLoose(titleRaw);
    if (!titleLoose) continue;

    const titleTokens = tokenizeLoose(titleRaw);

    // 3a) Strong rule: subject matches case title even with hyphens/prefixes (eg "Internal Know-How (2023-0006)")
    if (
      subjectLoose &&
      (titleLoose === subjectLoose || titleLoose.includes(subjectLoose) || subjectLoose.includes(titleLoose))
    ) {
      add(caseId, 98, "Email subject matches the case name.");
      continue;
    }

    // 3b) Token overlap in subject (fallback)
    if (titleTokens.length > 0 && subjectRaw) {
      const subjOverlap = tokenOverlapScore(titleTokens, subjectRaw);
      if (subjOverlap.hits >= 2) {
        const ratio = subjOverlap.total > 0 ? subjOverlap.hits / subjOverlap.total : 0;
        const boost = 60 + 30 * clamp(ratio, 0, 1); // 60..90
        add(caseId, boost, "Case name matches the email subject.");
      } else {
        // 3b-alt) Reverse direction: check what proportion of the *subject's* significant
        // tokens (>= 5 chars) appear in the case title. Handles short subjects like
        // "Amazon services" for case "Amazon vs. Gebrüder Weiss" — the title has 4 tokens
        // so only 1/4 hit the subject (misses the >=2 threshold above), but "amazon"
        // covers 50% of the meaningful subject words, which is a strong content signal.
        let altFired = false;
        const subjectSigTokens = tokenizeLoose(subjectRaw).filter((t) => t.length >= 5);
        if (subjectSigTokens.length > 0) {
          const { hits: sh, total: st } = tokenOverlapScore(subjectSigTokens, titleRaw);
          if (sh >= 1 && sh / st >= 0.5) {
            const boost = Math.round(35 + 25 * clamp(sh / st, 0, 1)); // 35..60
            add(caseId, boost, "Case name matches the email subject.");
            altFired = true;
          }
        }
        // 3b-alt2) Title-anchored check: any significant title token (≥6 chars) found
        // in the subject is a weak entity-name signal. Catches long subjects like
        // "Fwd: Amazon term sheet and due diligence materials" for case "Amazon vs. Gebrüder Weiss"
        // where subject-side ratio (1/5=20%) misses the 50% threshold above, but the
        // distinctive party name is clearly present in the subject.
        if (!altFired) {
          const titleSigTokens = titleTokens.filter((t) => t.length >= 6);
          if (titleSigTokens.length > 0) {
            const { hits: th } = tokenOverlapScore(titleSigTokens, subjectRaw);
            if (th >= 1) {
              const boost = 25 + 10 * Math.min(th - 1, 2); // 25 for 1 hit, 35 for 2, 45 for 3
              add(caseId, boost, "Case name matches the email subject.");
            }
          }
        }
      }
    }

    // 3b-bis) Fuzzy full-string similarity (typo-tolerant matching).
    // Catches cases like "Human recources" ↔ "Human Ressource" where exact token
    // overlap fails but the strings are clearly similar after edit-distance analysis.
    // Runs even when 3b fires so both signals can accumulate independently.
    if (subjectLoose) {
      const sim = subjectTitleSimilarity(subjectLoose, titleLoose);
      if (sim >= 0.88) {
        add(caseId, 90, `Subject similar to case name `);
      } else if (sim >= 0.78) {
        add(caseId, 75, `Subject similar to case name `);
      } else if (sim >= 0.70) {
        add(caseId, 55, `Subject similar to case name `);
      }
    }

    // 3c) Body mention: stricter to avoid false positives
    // Rule:
    // - if title is short (<= 2 tokens), require exact substring match in body
    // - otherwise require at least 2 token hits
    if (bodyRaw && titleTokens.length > 0) {
      const bodyHasExact = titleLoose.length >= 4 && bodyLoose.includes(titleLoose);

      if (titleTokens.length <= 2) {
        if (bodyHasExact) {
          add(caseId, 45, "Case name mentioned in the email body.");
        }
      } else {
        const bodyOverlap = tokenOverlapScore(titleTokens, bodyRaw);
        if (bodyOverlap.hits >= 2) {
          const ratio = bodyOverlap.total > 0 ? bodyOverlap.hits / bodyOverlap.total : 0;
          const boost = 35 + 25 * clamp(ratio, 0, 1); // 35..60
          add(caseId, boost, "Case name mentioned in the email body.");
        } else if (bodyHasExact) {
          add(caseId, 40, "Case name mentioned in the email body.");
        }
      }
    }
  }

  // Generic consumer domains that are too broad for useful domain-level suggestions
  const GENERIC_DOMAINS = new Set([
    "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com",
    "yahoo.com", "yahoo.co.uk", "icloud.com", "me.com", "mac.com",
    "protonmail.com", "proton.me", "aol.com", "msn.com",
  ]);

  // 4) Sender email history (count + recency weighted)
  const now = Date.now();
  if (fromEmail && senderToCase[fromEmail]) {
    for (const [caseId, meta] of Object.entries(senderToCase[fromEmail])) {
      const ageDays = (now - (meta.lastSeenAt || 0)) / (24 * 60 * 60 * 1000);
      const recency = Math.exp(-ageDays / 30); // half-weight at ~30 days
      const countWeight = log1p(meta.count || 0) / log1p(5); // normalise: count=5 → 1.0
      const baseBoost = Math.round(65 * (0.65 * clamp(countWeight, 0, 1) + 0.35 * recency));
      // Extra bonus for very recent filings (within last 2 days) so that even a single
      // filing is enough to auto-select when there is no competing signal.
      const freshBonus = ageDays < 2 ? 30 : 0;
      const boost = baseBoost + freshBonus;
      if (boost > 0) add(caseId, boost, "You often attach emails from this sender to this case.");
    }
  }

  // 5) Domain history — skip generic consumer domains (gmail, outlook, etc.)
  if (domain && !GENERIC_DOMAINS.has(domain) && domainToCase[domain]) {
    for (const [caseId, meta] of Object.entries(domainToCase[domain])) {
      const ageDays = (now - (meta.lastSeenAt || 0)) / (24 * 60 * 60 * 1000);
      const recency = Math.exp(-ageDays / 30);
      const countWeight = log1p(meta.count || 0) / log1p(5);
      const boost = Math.round(40 * (0.65 * clamp(countWeight, 0, 1) + 0.35 * recency));
      if (boost > 0) add(caseId, boost, "This domain often maps to this case.");
    }
  }

  // 6) Recent cases (low weight)
  for (const rc of recentCases || []) {
    const ageDays = (now - (rc.lastUsedAt || 0)) / (24 * 60 * 60 * 1000);
    const decay = clamp(1 - ageDays / 14, 0, 1);
    const boost = 12 * decay;
    if (boost > 0) add(rc.caseId, boost, "Recently used.");
  }

  const sorted = Object.entries(scoreByCase)
    .map(([caseId, score]) => ({
      caseId,
      score,
      reasons: reasonsByCase[caseId] || [],
    }))
    .sort((a, b) => b.score - a.score);

  const sortedScores = sorted.map((s) => s.score);

  // Debug: Show top 5 scores even if below threshold
  if (sorted.length > 0) {
    console.log("[suggestCasesLocal] 🔍 Top scores (before filtering):",
      sorted.slice(0, 5).map(s => ({
        caseId: s.caseId,
        score: s.score.toFixed(1),
        reasons: s.reasons,
      }))
    );
  } else {
    console.log("[suggestCasesLocal] ⚠️ No cases scored any points!");
    console.log("[suggestCasesLocal] Input:", {
      subjectLoose,
      bodyLoose: bodyLoose.substring(0, 100),
      casesCount: params.cases.length,
    });
  }

  const MIN_CONFIDENCE_PCT = 10;

  const suggestionsAll: CaseSuggestion[] = sorted.map((s, idx) => ({
    caseId: s.caseId,
    score: s.score,
    confidencePct: confidencePctFor(sortedScores, idx),
    reasons: s.reasons,
  }));

  const suggestions: CaseSuggestion[] = suggestionsAll
    .filter((s) => s.confidencePct >= MIN_CONFIDENCE_PCT)
    .slice(0, topK);

  const top = suggestions[0];
  const autoSelectCaseId = top && top.confidencePct >= 70 ? top.caseId : "";

  return { suggestions, autoSelectCaseId };
}

/**
 * Content-based case suggestions (triggered when user clicks "Vybrat jiný spis")
 *
 * This function ONLY analyzes subject and body content, ignoring:
 * - Thread history
 * - Sender/domain history
 * - Recent cases
 *
 * Use this when the user explicitly wants to re-suggest based on email content,
 * not on recipient history.
 */
export function suggestCasesByContent(params: {
  subject?: string;
  bodySnippet?: string;
  cases: CaseOption[];
  topK?: number;
}): { suggestions: CaseSuggestion[] } {
  const topK = params.topK ?? 5;

  const subjectRaw = String(params.subject || "");
  const subjectLoose = normLoose(subjectRaw);
  const subjectStrict = normText(subjectRaw);

  const bodyRaw = String(params.bodySnippet || "").slice(0, 1500); // Limit to first 1500 chars
  const bodyLoose = normLoose(bodyRaw);
  const bodyStrict = normText(bodyRaw);

  const reasonsByCase: Record<string, string[]> = {};
  const scoreByCase: Record<string, number> = {};

  const add = (caseId: string, delta: number, reason?: string) => {
    if (!caseId) return;
    scoreByCase[caseId] = (scoreByCase[caseId] || 0) + delta;
    if (reason) {
      if (!reasonsByCase[caseId]) reasonsByCase[caseId] = [];
      if (!reasonsByCase[caseId].includes(reason)) reasonsByCase[caseId].push(reason);
    }
  };

  // 1) Visible case reference match (very strong)
  for (const c of params.cases) {
    const anyC = c as any;
    const caseId = String(anyC?.id || "");
    const vis = normText(getCaseVisibleId(anyC));
    if (!caseId || !vis) continue;

    const inSubject = subjectStrict.includes(vis);
    const inBody = bodyStrict.includes(vis);

    if (inSubject || inBody) {
      add(caseId, 95, "Case reference found in email content");
    }
  }

  // 2) Case title matching
  for (const c of params.cases) {
    const anyC = c as any;
    const caseId = String(anyC?.id || "");
    if (!caseId) continue;

    const titleRaw = getCaseTitle(anyC);
    const titleLoose = normLoose(titleRaw);
    if (!titleLoose) continue;

    const titleTokens = tokenizeLoose(titleRaw);

    // 2a) Strong rule: subject matches case title
    if (
      subjectLoose &&
      (titleLoose === subjectLoose || titleLoose.includes(subjectLoose) || subjectLoose.includes(titleLoose))
    ) {
      add(caseId, 98, "Subject matches case name");
      continue;
    }

    // 2b) Token overlap in subject
    if (titleTokens.length > 0 && subjectRaw) {
      const subjOverlap = tokenOverlapScore(titleTokens, subjectRaw);
      if (subjOverlap.hits >= 2) {
        const ratio = subjOverlap.total > 0 ? subjOverlap.hits / subjOverlap.total : 0;
        const boost = 60 + 30 * clamp(ratio, 0, 1);
        add(caseId, boost, "Case name matches subject keywords");
      } else {
        // 2b-alt) Reverse direction: same logic as 3b-alt in suggestCasesLocal.
        let altFired = false;
        const subjectSigTokens = tokenizeLoose(subjectRaw).filter((t) => t.length >= 5);
        if (subjectSigTokens.length > 0) {
          const { hits: sh, total: st } = tokenOverlapScore(subjectSigTokens, titleRaw);
          if (sh >= 1 && sh / st >= 0.5) {
            const boost = Math.round(35 + 25 * clamp(sh / st, 0, 1)); // 35..60
            add(caseId, boost, "Case name matches subject keywords");
            altFired = true;
          }
        }
        // 2b-alt2) Title-anchored check (same logic as 3b-alt2 in suggestCasesLocal).
        if (!altFired) {
          const titleSigTokens = titleTokens.filter((t) => t.length >= 6);
          if (titleSigTokens.length > 0) {
            const { hits: th } = tokenOverlapScore(titleSigTokens, subjectRaw);
            if (th >= 1) {
              const boost = 25 + 10 * Math.min(th - 1, 2); // 25..45
              add(caseId, boost, "Case name matches subject keywords");
            }
          }
        }
      }
    }

    // 2b-bis) Fuzzy full-string similarity (same logic as suggestCasesLocal 3b-bis)
    if (subjectLoose) {
      const sim = subjectTitleSimilarity(subjectLoose, titleLoose);
      if (sim >= 0.88) {
        add(caseId, 90, `Subject similar to case name `);
      } else if (sim >= 0.78) {
        add(caseId, 75, `Subject similar to case name `);
      } else if (sim >= 0.70) {
        add(caseId, 55, `Subject similar to case name `);
      }
    }

    // 2c) Body mention
    if (bodyRaw && titleTokens.length > 0) {
      const bodyHasExact = titleLoose.length >= 4 && bodyLoose.includes(titleLoose);

      if (titleTokens.length <= 2) {
        if (bodyHasExact) {
          add(caseId, 45, "Case name found in email body");
        }
      } else {
        const bodyOverlap = tokenOverlapScore(titleTokens, bodyRaw);
        if (bodyOverlap.hits >= 2) {
          const ratio = bodyOverlap.total > 0 ? bodyOverlap.hits / bodyOverlap.total : 0;
          const boost = 35 + 25 * clamp(ratio, 0, 1);
          add(caseId, boost, "Case name found in email body");
        } else if (bodyHasExact) {
          add(caseId, 40, "Case name found in email body");
        }
      }
    }
  }

  const sorted = Object.entries(scoreByCase)
    .map(([caseId, score]) => ({
      caseId,
      score,
      reasons: reasonsByCase[caseId] || [],
    }))
    .sort((a, b) => b.score - a.score);

  const sortedScores = sorted.map((s) => s.score);

  const MIN_CONFIDENCE_PCT = 10;

  const suggestionsAll: CaseSuggestion[] = sorted.map((s, idx) => ({
    caseId: s.caseId,
    score: s.score,
    confidencePct: confidencePctFor(sortedScores, idx),
    reasons: s.reasons,
  }));

  const suggestions: CaseSuggestion[] = suggestionsAll
    .filter((s) => s.confidencePct >= MIN_CONFIDENCE_PCT)
    .slice(0, topK);

  return { suggestions };
}