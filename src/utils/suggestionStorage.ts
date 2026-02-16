// src/utils/caseSuggestStorage.ts

export type CaseStats = { count: number; lastSeenAt: number };
export type PerEntityCaseStats = Record<string, Record<string, CaseStats>>;
// senderEmail -> caseId -> stats OR domain -> caseId -> stats

export type RecentCase = { caseId: string; lastUsedAt: number; useCount: number };

export type CaseSuggestState = {
  version: 1;
  threadToCase: Record<string, { caseId: string; lastSeenAt: number }>;
  senderToCase: PerEntityCaseStats;
  domainToCase: PerEntityCaseStats;
  recentCases: RecentCase[];
};

const VERSION = 1;

const LIMITS = {
  maxThreads: 400,
  maxSenders: 300,
  maxDomains: 200,
  maxCasesPerSender: 30,
  maxCasesPerDomain: 30,
  maxRecentCases: 30,
};

function nowMs(): number {
  return Date.now();
}

function safeLower(s: string): string {
  return String(s || "").trim().toLowerCase();
}

function domainFromEmail(email: string): string {
  const e = safeLower(email);
  const at = e.lastIndexOf("@");
  if (at < 0) return "";
  return e.slice(at + 1).trim();
}

function getMailboxKeySafe(): string {
  try {
    const p = (Office as any)?.context?.mailbox?.userProfile;
    const email = String(p?.emailAddress || "").toLowerCase().trim();
    return email || "default";
  } catch {
    return "default";
  }
}

function suggestStorageKey(): string {
  return `sc_case_suggest:${getMailboxKeySafe()}`;
}

const DEFAULT_STATE: CaseSuggestState = {
  version: 1,
  threadToCase: {},
  senderToCase: {},
  domainToCase: {},
  recentCases: [],
};

export function loadCaseSuggestState(): CaseSuggestState {
  try {
    const raw = localStorage.getItem(suggestStorageKey());
    if (!raw) return DEFAULT_STATE;

    const parsed = JSON.parse(raw) as CaseSuggestState;
    if (!parsed || parsed.version !== VERSION) return DEFAULT_STATE;

    return {
      ...DEFAULT_STATE,
      ...parsed,
      threadToCase: parsed.threadToCase || {},
      senderToCase: parsed.senderToCase || {},
      domainToCase: parsed.domainToCase || {},
      recentCases: Array.isArray(parsed.recentCases) ? parsed.recentCases : [],
    };
  } catch {
    return DEFAULT_STATE;
  }
}

export function saveCaseSuggestState(state: CaseSuggestState) {
  try {
    localStorage.setItem(suggestStorageKey(), JSON.stringify(state));
  } catch {
    // ignore
  }
}

function pruneThreadMap(
  threadToCase: Record<string, { caseId: string; lastSeenAt: number }>
) {
  const entries = Object.entries(threadToCase);
  if (entries.length <= LIMITS.maxThreads) return threadToCase;

  entries.sort((a, b) => (b[1].lastSeenAt || 0) - (a[1].lastSeenAt || 0));
  return Object.fromEntries(entries.slice(0, LIMITS.maxThreads));
}

function pruneEntityStats(stats: PerEntityCaseStats, maxEntities: number, maxCasesPerEntity: number) {
  const entities = Object.keys(stats);
  if (entities.length > maxEntities) {
    const entityLastSeen: Array<[string, number]> = entities.map((entity) => {
      const perCase = stats[entity] || {};
      const last = Math.max(0, ...Object.values(perCase).map((x) => x.lastSeenAt || 0));
      return [entity, last];
    });

    entityLastSeen.sort((a, b) => b[1] - a[1]);
    const keep = new Set(entityLastSeen.slice(0, maxEntities).map((x) => x[0]));

    for (const e of entities) {
      if (!keep.has(e)) delete stats[e];
    }
  }

  for (const entity of Object.keys(stats)) {
    const perCase = stats[entity] || {};
    const entries = Object.entries(perCase);
    if (entries.length <= maxCasesPerEntity) continue;

    entries.sort((a, b) => (b[1].lastSeenAt || 0) - (a[1].lastSeenAt || 0));
    stats[entity] = Object.fromEntries(entries.slice(0, maxCasesPerEntity));
  }
}

function pruneRecentCases(recent: RecentCase[]) {
  const sorted = [...recent].sort((a, b) => (b.lastUsedAt || 0) - (a.lastUsedAt || 0));
  return sorted.slice(0, LIMITS.maxRecentCases);
}

export function getThreadMappedCaseId(conversationKey: string): string {
  const key = safeLower(conversationKey);
  if (!key) return "";

  const state = loadCaseSuggestState();
  return state.threadToCase[key]?.caseId || "";
}

export function recordSuccessfulAttach(params: {
  caseId: string;
  conversationKey?: string;
  senderEmail?: string;
}) {
  const caseId = String(params.caseId || "").trim();
  if (!caseId) return;

  const t = nowMs();
  const state = loadCaseSuggestState();

  const conv = safeLower(params.conversationKey || "");
  if (conv) {
    state.threadToCase[conv] = { caseId, lastSeenAt: t };
  }

  const sender = safeLower(params.senderEmail || "");
  if (sender) {
    if (!state.senderToCase[sender]) state.senderToCase[sender] = {};
    const sSlot = state.senderToCase[sender][caseId] || { count: 0, lastSeenAt: 0 };
    sSlot.count += 1;
    sSlot.lastSeenAt = t;
    state.senderToCase[sender][caseId] = sSlot;

    const dom = domainFromEmail(sender);
    if (dom) {
      if (!state.domainToCase[dom]) state.domainToCase[dom] = {};
      const dSlot = state.domainToCase[dom][caseId] || { count: 0, lastSeenAt: 0 };
      dSlot.count += 1;
      dSlot.lastSeenAt = t;
      state.domainToCase[dom][caseId] = dSlot;
    }
  }

  const existing = state.recentCases.find((x) => x.caseId === caseId);
  if (existing) {
    existing.useCount += 1;
    existing.lastUsedAt = t;
  } else {
    state.recentCases.push({ caseId, useCount: 1, lastUsedAt: t });
  }

  state.threadToCase = pruneThreadMap(state.threadToCase);
  pruneEntityStats(state.senderToCase, LIMITS.maxSenders, LIMITS.maxCasesPerSender);
  pruneEntityStats(state.domainToCase, LIMITS.maxDomains, LIMITS.maxCasesPerDomain);
  state.recentCases = pruneRecentCases(state.recentCases);

  saveCaseSuggestState({ ...state, version: 1 });
}

export function getSuggestStats() {
  const state = loadCaseSuggestState();
  return {
    threadToCase: state.threadToCase,
    senderToCase: state.senderToCase,
    domainToCase: state.domainToCase,
    recentCases: state.recentCases,
  };
}
