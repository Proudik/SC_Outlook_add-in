import * as React from "react";
import PanelCard from "./PanelCard";
import CaseSelector from "./CaseSelector";
import {
  listCases,
  listClients,
  listTimers,
  createTimer,
  updateTimer,
  listUsers,
  CaseOption,
  TimerItem,
  UpsertTimerPayload,
  UserOption,
  ClientOption,
} from "../../services/singlecase";
import type { CaseListScope } from "./SettingsModal";
import { useOutlookSuggestions } from "../../hooks/useOutlookSuggestions";
import { useCaseSuggestions } from "../../hooks/useCaseSuggestions";
import { loadSentPill } from "../../utils/sentPillStore";
import "./TimesheetsPanel.css";
import { getThreadMappedCaseId } from "../../utils/suggestionStorage";

type Props = {
  token: string;
  onBack: () => void;
};

type TsSelectedSource = "" | "attached" | "thread" | "suggested" | "manual";

// what useCaseSuggestions expects in your codebase is very likely only these:
type SuggestHookSelectedSource = "" | "suggested" | "manual";

function pad2(n: number) {
  return String(n).padStart(2, "0");
}

function toYmd(d: Date) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function secondsToHm(totalSeconds: number) {
  const s = Math.max(0, Number(totalSeconds || 0));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  return { h, m };
}

function hmToSeconds(h: number, m: number) {
  const hh = Number.isFinite(h) ? Math.max(0, Math.floor(h)) : 0;
  const mm = Number.isFinite(m) ? Math.max(0, Math.floor(m)) : 0;
  return hh * 3600 + mm * 60;
}

function formatUserLabel(u: UserOption) {
  const name = `${(u.firstName || "").trim()} ${(u.lastName || "").trim()}`.trim();
  if (name && u.username) return `${name} (${u.username})`;
  return name || u.username || `User ${u.id}`;
}

function isClosedStatus(status?: string | null): boolean {
  const s = (status || "").toLowerCase();
  if (!s) return false;
  return s.includes("closed") || s.includes("uzav") || s.includes("archiv") || s.includes("done");
}

function getDefaultRange() {
  const now = new Date();
  const from = toYmd(new Date(now.getFullYear(), now.getMonth(), 1));
  const to = toYmd(now);
  return { from, to };
}

export default function TimesheetsPanel({ token, onBack }: Props) {
  const { email, error: emailError } = useOutlookSuggestions();

  const emailItemId = ((email as any)?.itemId || (email as any)?.id || "").toString();
  const emailSubject = ((email as any)?.subject || "").toString();
  const conversationKey = ((email as any)?.conversationKey || "").toString();

  const fromEmail = (
    (email as any)?.fromEmail ||
    (email as any)?.from?.email ||
    (email as any)?.from?.emailAddress ||
    (email as any)?.from?.address ||
    ""
  ).toString();

  const [cases, setCases] = React.useState<CaseOption[]>([]);
  const [clientNamesById, setClientNamesById] = React.useState<Record<string, string>>({});
  const [users, setUsers] = React.useState<UserOption[]>([]);

  const [caseScope, setCaseScope] = React.useState<CaseListScope>("all");
  const [isLoadingCases, setIsLoadingCases] = React.useState(false);

  const [selectedCaseId, setSelectedCaseId] = React.useState("");
  const [selectedSource, setSelectedSource] = React.useState<TsSelectedSource>("");

  const [isLoading, setIsLoading] = React.useState(false);
  const [items, setItems] = React.useState<TimerItem[]>([]);
  const [error, setError] = React.useState<string | null>(null);

  const [isModalOpen, setIsModalOpen] = React.useState(false);
  const [editingId, setEditingId] = React.useState<string | null>(null);

  React.useEffect(() => {
    let mounted = true;
    void (async () => {
      setIsLoadingCases(true);
      try {
        const [casesRes, clientsRes, usersRes] = await Promise.all([
          listCases(token, caseScope),
          listClients(token),
          listUsers(token),
        ]);

        if (!mounted) return;

        setCases(casesRes);
        setUsers(usersRes);

        const map: Record<string, string> = {};
        for (const c of clientsRes as ClientOption[]) map[c.id] = c.name;
        setClientNamesById(map);
      } catch (e) {
        if (!mounted) return;
        setError(e instanceof Error ? e.message : String(e));
      } finally {
        if (mounted) setIsLoadingCases(false);
      }
    })();

    return () => {
      mounted = false;
    };
  }, [token, caseScope]);

  const visibleCases = React.useMemo(() => {
    if (caseScope === "all") return cases;
    return cases.filter((c) => !isClosedStatus((c as any)?.status));
  }, [cases, caseScope]);

  React.useEffect(() => {
    let mounted = true;

    void (async () => {
      if (!emailItemId) return;

      try {
        const pill = await loadSentPill(emailItemId);
        const pillCaseId = pill?.caseId ? String(pill.caseId) : "";

        const threadCaseId =
          !pillCaseId && conversationKey ? getThreadMappedCaseId(conversationKey) : "";

        const picked = pillCaseId || threadCaseId;

        if (!picked) return;
        if (!mounted) return;

        setSelectedCaseId(picked);
        setSelectedSource(pillCaseId ? "attached" : "thread");
      } catch {
        // ignore
      }
    })();

    return () => {
      mounted = false;
    };
  }, [emailItemId, conversationKey]);

  const handleAutoSelectCaseId = React.useCallback(
    (id: string) => {
      if (selectedSource === "attached" || selectedSource === "thread") return;
      setSelectedCaseId(id);
      setSelectedSource("suggested");
    },
    [selectedSource]
  );

  const suggestSelectedSource: SuggestHookSelectedSource =
    selectedSource === "suggested" ? "suggested" : selectedSource === "manual" ? "manual" : "";

  const { suggestions: caseSuggestions } = useCaseSuggestions({
    enabled: true,
    emailItemId,
    conversationKey,
    subject: emailSubject,
    bodySnippet: "",
    fromEmail,
    attachments: [],
    cases: visibleCases,
    selectedCaseId,
    selectedSource: suggestSelectedSource,
    onAutoSelectCaseId: handleAutoSelectCaseId,
    topK: 3,
  });

  React.useEffect(() => {
    if (!selectedCaseId) return;

    const existsInVisible = visibleCases.some((c) => String(c.id) === String(selectedCaseId));
    if (existsInVisible) return;

    if (caseScope !== "all") setCaseScope("all");
  }, [selectedCaseId, visibleCases, caseScope]);

  React.useEffect(() => {
    let mounted = true;

    if (!selectedCaseId) {
      setItems([]);
      return () => {};
    }

    void (async () => {
      setIsLoading(true);
      setError(null);
      try {
        const { from, to } = getDefaultRange();

        const res = await listTimers(token, {
          projectId: selectedCaseId,
          from,
          to,
        });

        if (!mounted) return;
        setItems(res);
      } catch (e) {
        if (!mounted) return;
        setError(e instanceof Error ? e.message : String(e));
        setItems([]);
      } finally {
        if (mounted) setIsLoading(false);
      }
    })();

    return () => {
      mounted = false;
    };
  }, [token, selectedCaseId]);

  const totals = React.useMemo(() => {
    const total = items.reduce((sum, x) => sum + Number(x.total_time || 0), 0);
    const billed = items.reduce((sum, x) => sum + Number(x.total_billed_time || 0), 0);
    return { total, billed };
  }, [items]);

  const openCreate = () => {
    setEditingId(null);
    setIsModalOpen(true);
  };

  const openEdit = (it: TimerItem) => {
    setEditingId((it as any).id ? String((it as any).id) : null);
    setIsModalOpen(true);
  };

  const onSave = async (payload: UpsertTimerPayload) => {
    if (!selectedCaseId) return;

    if (editingId) {
      await updateTimer(token, editingId, payload);
    } else {
      await createTimer(token, payload);
    }

    const { from, to } = getDefaultRange();
    const refreshed = await listTimers(token, { projectId: selectedCaseId, from, to });
    setItems(refreshed);

    setIsModalOpen(false);
    setEditingId(null);
  };

  const totalHm = secondsToHm(totals.total);
  const billedHm = secondsToHm(totals.billed);

  return (
    <PanelCard>
      <div className="ts-root">
        <div className="ts-body">
          <div className="ts-topbar">
            <button type="button" className="ts-back" onClick={onBack}>
              ‹ Back
            </button>
            <div className="ts-title">Timesheets</div>
            <div className="ts-spacer" />
          </div>

          {emailError ? <div className="ts-error">{String(emailError)}</div> : null}

          <div className="ts-section">
            <CaseSelector
              title="Case"
              scope={caseScope}
              onScopeChange={(scope) => setCaseScope(scope)}
              selectedCaseId={selectedCaseId}
              onSelectCaseId={(id) => {
                setSelectedCaseId(id);
                setSelectedSource("manual");
              }}
              suggestions={caseSuggestions}
              cases={visibleCases}
              isLoadingCases={isLoadingCases}
              clientNamesById={clientNamesById}
            />

            {selectedSource === "attached" ? (
              <div className="ts-muted" style={{ marginTop: 6 }}>
                Case preselected from attached email
              </div>
            ) : null}

            {selectedSource === "thread" ? (
              <div className="ts-muted" style={{ marginTop: 6 }}>
                Case preselected from this email thread
              </div>
            ) : null}
          </div>

          <div className="ts-summary">
            <div className="ts-summary-left">
              <div className="ts-summary-title">Vykazy</div>
              <div className="ts-summary-sub">
                Total {totalHm.h}h {pad2(totalHm.m)}m Billed {billedHm.h}h {pad2(billedHm.m)}m
              </div>
            </div>
          </div>

          {error ? <div className="ts-error">{error}</div> : null}
          {isLoading ? <div className="ts-muted">Loading…</div> : null}

          <div className="ts-list">
            {items.map((it, idx) => {
              const dt = (it.date || "").slice(0, 10);
              const t = secondsToHm(Number(it.total_time || 0));
              const b = secondsToHm(Number(it.total_billed_time || 0));
              const note = (it.note || "").trim();

              return (
                <div
                  key={(it as any).id || `${dt}-${idx}`}
                  className="ts-row"
                  onClick={() => openEdit(it)}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" || e.key === " ") openEdit(it);
                  }}
                >
                  <div className="ts-row-left">
                    <div className="ts-row-date">{dt}</div>
                    <div className="ts-row-note">{note || "No note"}</div>
                  </div>
                  <div className="ts-row-right">
                    <div className="ts-row-time">
                      {t.h}h {pad2(t.m)}m
                    </div>
                    {Number(it.total_billed_time || 0) !== Number(it.total_time || 0) ? (
                      <div className="ts-row-billed">
                        Billed {b.h}h {pad2(b.m)}m
                      </div>
                    ) : null}
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        <div className="ts-footer">
          <button
            type="button"
            className="ts-primary"
            onClick={openCreate}
            disabled={!selectedCaseId}
          >
            Add timesheet
          </button>
        </div>

        {isModalOpen ? (
          <TimesheetModal
            token={token}
            caseId={selectedCaseId}
            users={users}
            editingId={editingId}
            onClose={() => {
              setIsModalOpen(false);
              setEditingId(null);
            }}
            onSave={onSave}
          />
        ) : null}
      </div>
    </PanelCard>
  );
}

type ModalProps = {
  token: string;
  caseId: string;
  users: UserOption[];
  editingId: string | null;
  onClose: () => void;
  onSave: (payload: UpsertTimerPayload) => Promise<void>;
};

function TimesheetModal({ token, caseId, users, editingId, onClose, onSave }: ModalProps) {
  void token;

  const [date, setDate] = React.useState(() => toYmd(new Date()));
  const [h, setH] = React.useState(0);
  const [m, setM] = React.useState(0);

  const [billableDifferent, setBillableDifferent] = React.useState(false);
  const [bh, setBh] = React.useState(0);
  const [bm, setBm] = React.useState(0);

  const [note, setNote] = React.useState("");
  const [userId, setUserId] = React.useState<string>("");

  React.useEffect(() => {
    if (users.length === 1) setUserId(users[0].id);
  }, [users]);

  const [saving, setSaving] = React.useState(false);
  const [err, setErr] = React.useState<string | null>(null);

  const submit = async () => {
    setErr(null);

    if (!caseId) {
      setErr("Select a case first.");
      return;
    }
    if (!userId) {
      setErr("User is required.");
      return;
    }
    if (!date) {
      setErr("Date is required.");
      return;
    }
    if (m < 0 || m > 59 || bm < 0 || bm > 59) {
      setErr("Minutes must be 0 to 59.");
      return;
    }

    const total = hmToSeconds(h, m);
    if (total <= 0) {
      setErr("Time must be greater than 0.");
      return;
    }

    const billed = billableDifferent ? hmToSeconds(bh, bm) : total;
    if (billed < 0) {
      setErr("Billed time must be valid.");
      return;
    }

    const payload: UpsertTimerPayload = {
      user_id: Number(userId),
      project_id: Number(caseId),
      date: `${date} 00:00:00`,
      total_time: total,
      total_billed_time: billed,
      sheet_activity_id: 1,
      note: note || "",
    };

    setSaving(true);
    try {
      await onSave(payload);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="ts-modal-overlay" role="dialog" aria-modal="true">
      <div className="ts-modal">
        <div className="ts-modal-header">
          <div className="ts-modal-title">{editingId ? "Edit timesheet" : "Create timesheet"}</div>
          <button type="button" className="ts-x" onClick={onClose}>
            ×
          </button>
        </div>

        <div className="ts-modal-grid">
          <div>
            <div className="ts-label">User</div>
            <select className="ts-select" value={userId} onChange={(e) => setUserId(e.target.value)}>
              <option value="">Select user</option>
              {users.map((u) => (
                <option key={u.id} value={u.id}>
                  {formatUserLabel(u)}
                </option>
              ))}
            </select>
          </div>

          <div>
            <div className="ts-label">Date</div>
            <input className="ts-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
          </div>

          <div className="ts-time-row">
            <div>
              <div className="ts-label">Hours</div>
              <input className="ts-input" type="number" value={h} onChange={(e) => setH(Number(e.target.value))} />
            </div>
            <div>
              <div className="ts-label">Minutes</div>
              <input className="ts-input" type="number" value={m} onChange={(e) => setM(Number(e.target.value))} />
            </div>
          </div>

          <label className="ts-check">
            <input type="checkbox" checked={billableDifferent} onChange={(e) => setBillableDifferent(e.target.checked)} />
            Billable time is different
          </label>

          {billableDifferent ? (
            <div className="ts-time-row">
              <div>
                <div className="ts-label">Billed hours</div>
                <input className="ts-input" type="number" value={bh} onChange={(e) => setBh(Number(e.target.value))} />
              </div>
              <div>
                <div className="ts-label">Billed minutes</div>
                <input className="ts-input" type="number" value={bm} onChange={(e) => setBm(Number(e.target.value))} />
              </div>
            </div>
          ) : null}

          <div className="ts-note">
            <div className="ts-label">Details</div>
            <textarea className="ts-textarea" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Note" />
          </div>

          {err ? <div className="ts-error">{err}</div> : null}
        </div>

        <div className="ts-modal-actions">
          <button type="button" className="ts-secondary" onClick={onClose} disabled={saving}>
            Cancel
          </button>
          <button type="button" className="ts-primary" onClick={submit} disabled={saving}>
            {saving ? "Saving…" : "Save"}
          </button>
        </div>
      </div>
    </div>
  );
}