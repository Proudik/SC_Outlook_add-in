// Folder: src/taskpane/components/CaseIntakePanel.tsx

import * as React from "react";
import SideKickCard from "./SideKickCard";
import { useOutlookSuggestions } from "../../hooks/useOutlookSuggestions";
import { useCaseSuggestions } from "../../hooks/useCaseSuggestions";
import { listCases, listClients, submitEmailToCase, CaseOption } from "../../services/singlecase";
import EmailContextPanel from "./EmailContextPanel";
import CaseSelector from "./CaseSelector";
import AttachmentsPicker, { OutlookAttachment } from "./AttachmentsPicker";
import "./CaseIntakePanel.css";
import type { CaseIntakeSettings } from "./CaseIntakeSettingsModal";
import {
  loadLastCaseId,
  saveLastCaseId,
  loadDuplicateCache,
  saveDuplicateCache,
  hasAttached,
  markAttached,
  markEverFiled,
  isDiscardedEmail,
  setDiscardedEmail,
} from "../../utils/caseIntakeStorage";
import { uploadDocumentToCase, uploadDocumentVersion, getDocumentMeta } from "../../services/singlecaseDocuments";
import StatusPill from "./StatusPill";
import { loadSentPill, saveSentPill, SentPillData } from "../../utils/sentPillStore";
import { recordSuccessfulAttach } from "../../utils/caseSuggestStorage";
import { getStored } from "../../utils/storage";
import { STORAGE_KEYS } from "../../utils/constants";
import { loadUploadedLinks, saveUploadedLinks } from "../../utils/uploadedLinksStore";

type Props = {
  token: string;
  onBack: () => void;
  onSignOut: () => Promise<void> | void;

  settings: CaseIntakeSettings;
  onChangeSettings: React.Dispatch<React.SetStateAction<CaseIntakeSettings>>;
};

type AttachmentPayload = {
  id: string;
  name: string;
  contentType?: string;
  size: number;
  contentBase64: string;
};

type UploadedItem = {
  id: string;
  name: string;
  url: string;
  kind: "email" | "attachment";
  atIso: string;
};

type ChatMsg =
  | { id: string; kind: "system"; title?: string; text?: string }
  | { id: string; kind: "embed"; node: "email" | "caseSelector" | "attachments" }
  | { id: string; kind: "timeline"; items: Array<{ label: string; meta?: string }> }
  | { id: string; kind: "actions"; row: "choose" | "attachments" | "ready" | "discardedUndo" | "sent" };

type Mode = "idle" | "new" | "attachments" | "ready" | "discarded" | "sent";

function allSettled<T>(promises: Array<Promise<T>>) {
  return Promise.all(
    promises.map((p) =>
      p.then(
        (value) => ({ status: "fulfilled" as const, value }),
        (reason) => ({ status: "rejected" as const, reason })
      )
    )
  );
}

function extractDocMeta(raw: any): any {
  if (!raw) return null;
  if (raw.document) return raw.document;
  if (raw.data) return raw.data;
  if (raw.result) return raw.result;
  return raw;
}

function extractLatestRevisionNumber(doc: any): number | null {
  const n = doc?.latest_version?.revision_number;
  const num = Number(n);
  return Number.isFinite(num) && num > 0 ? num : null;
}

function isClosedStatus(status?: string | null): boolean {
  const s = (status || "").toLowerCase();
  if (!s) return false;
  return s.includes("closed") || s.includes("uzav") || s.includes("archiv") || s.includes("done");
}

function getCaseDisplay(
  cases: CaseOption[],
  clientNamesById: Record<string, string>,
  caseId: string
): { caseLabel: string; clientLabel: string } {
  const c: any = cases.find((x: any) => String(x.id) === String(caseId));

  const caseLabel =
    (c?.number && c?.name ? `${c.number} · ${c.name}` : "") ||
    (c?.caseNumber && c?.title ? `${c.caseNumber} · ${c.title}` : "") ||
    (c?.label ? String(c.label) : "") ||
    (c?.title ? String(c.title) : "") ||
    caseId;

  const clientLabelFromCase =
    c?.clientName ||
    c?.client_name ||
    c?.clientTitle ||
    c?.client?.name ||
    c?.client_label ||
    c?.clientDisplay ||
    "";

  const clientId = c?.clientId || c?.client_id || c?.client?.id;

  const clientLabel =
    (clientLabelFromCase ? String(clientLabelFromCase) : "") ||
    (clientId && clientNamesById[String(clientId)]) ||
    "Client";

  return { caseLabel, clientLabel };
}

async function getEmailBodySnippet(maxLen: number): Promise<string> {
  try {
    const item = Office?.context?.mailbox?.item as any;
    if (!item?.body?.getAsync) return "";

    const text: string = await new Promise((resolve) => {
      item.body.getAsync(Office.CoercionType.Text, (res: any) => {
        if (res?.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || ""));
        else resolve("");
      });
    });

    const trimmed = text.trim();
    if (!trimmed) return "";
    return trimmed.length > maxLen ? trimmed.slice(0, maxLen) : trimmed;
  } catch {
    return "";
  }
}

function toBase64Utf8(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

function safeFileName(value: string): string {
  const v = (value || "").trim();
  const cleaned = v.replace(/[<>:"/\\|?*\x00-\x1F]/g, " ").replace(/\s+/g, " ").trim();
  return cleaned.slice(0, 80) || "email";
}

function buildDocumentUrl(host: string, documentId: string): string {
  const h = (host || "").trim().replace(/^https?:\/\//i, "").split("/")[0];
  if (!h || !documentId) return "";
  return `https://${h}/?/documents/view/${encodeURIComponent(documentId)}`;
}

function getCurrentOutlookAttachments(): OutlookAttachment[] {
  try {
    const item = Office?.context?.mailbox?.item as any;
    const atts = (item?.attachments || []) as any[];
    return atts.map((a) => ({
      id: String(a.id),
      name: String(a.name || ""),
      size: Number(a.size || 0),
      isInline: Boolean(a.isInline),
      contentType: a.contentType ? String(a.contentType) : undefined,
    }));
  } catch {
    return [];
  }
}

function getAttachmentContentBase64(attachmentId: string): Promise<string | null> {
  return new Promise((resolve) => {
    const item = Office?.context?.mailbox?.item as any;
    if (!item?.getAttachmentContentAsync) {
      resolve(null);
      return;
    }

    item.getAttachmentContentAsync(attachmentId, (res: any) => {
      if (!res || res.status !== Office.AsyncResultStatus.Succeeded) {
        resolve(null);
        return;
      }

      const value = res.value || {};
      const content = value.content;
      const format = String(value.format || "").toLowerCase();

      if (!content || !format) {
        resolve(null);
        return;
      }

      if (format === "base64") {
        resolve(String(content));
        return;
      }

      resolve(null);
    });
  });
}

function extractRevisionFromVersionUploadResponse(raw: any): number | null {
  if (!raw) return null;
  const v = raw.data || raw.result || raw.document || raw.version || raw;

  const candidate =
    v?.latest_version?.revision_number ??
    v?.latestVersion?.revisionNumber ??
    v?.latestVersion?.revision_number ??
    v?.revision_number ??
    v?.revisionNumber ??
    v?.revision;

  const n = Number(candidate);
  return Number.isFinite(n) && n > 0 ? n : null;
}

function fmtCs(iso?: string): string {
  if (!iso) return "";
  try {
    return new Date(iso).toLocaleString("cs-CZ");
  } catch {
    return "";
  }
}

export default function CaseIntakePanel({
  token,
 
 
  settings,
  onChangeSettings,
}: Props) {  const { email, error: emailError } = useOutlookSuggestions();

  const [clientNamesById, setClientNamesById] = React.useState<Record<string, string>>({});
  const [cases, setCases] = React.useState<CaseOption[]>([]);
  const [isLoadingCases, setIsLoadingCases] = React.useState(false);

  const [selectedCaseId, setSelectedCaseId] = React.useState<string>("");
  const [selectedSource, setSelectedSource] = React.useState<"" | "remembered" | "suggested" | "manual">("");

  const [workspaceHost, setWorkspaceHost] = React.useState<string>("");

  const [attachments, setAttachments] = React.useState<OutlookAttachment[]>([]);
  const [selectedAttachmentIds, setSelectedAttachmentIds] = React.useState<Set<string>>(new Set());
  const [isLoadingAttachments, setIsLoadingAttachments] = React.useState(false);

  const [suggestBodySnippet, setSuggestBodySnippet] = React.useState("");
  const [uploadedItems, setUploadedItems] = React.useState<UploadedItem[]>([]);
  const [sentPill, setSentPill] = React.useState<SentPillData | null>(null);

  const [dupPromptOpen, setDupPromptOpen] = React.useState(false);
  const dupDecisionRef = React.useRef<null | "confirm">(null);

  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [submitResult, setSubmitResult] = React.useState<{
    ok: boolean;
    id?: string;
    emailDocumentId?: string;
    uploadedCount?: number;
    failedUploads?: Array<{ name: string; error: string }>;
  } | null>(null);

  const [error, setError] = React.useState<string | null>(null);

  const emailItemId = ((email as any)?.itemId || (email as any)?.id || "").toString();
  const emailSubject = ((email as any)?.subject || "").toString();

  const fromName = (
    (email as any)?.fromName ||
    (email as any)?.from?.name ||
    (email as any)?.from?.displayName ||
    ""
  ).toString();

  const fromEmail = (
    (email as any)?.fromEmail ||
    (email as any)?.from?.email ||
    (email as any)?.from?.emailAddress ||
    (email as any)?.from?.address ||
    ""
  ).toString();

  const visibleCases = React.useMemo(() => {
    if (settings.caseListScope === "all") return cases;
    return cases.filter((c) => !isClosedStatus((c as any)?.status));
  }, [cases, settings.caseListScope]);

  const selectedCaseDisplay = React.useMemo(() => {
    if (!selectedCaseId) return null;
    return getCaseDisplay(visibleCases, clientNamesById, selectedCaseId);
  }, [visibleCases, clientNamesById, selectedCaseId]);

  const existingEmailDocId = React.useMemo(() => {
    const fromPill = sentPill?.documentId ? String(sentPill.documentId) : "";
    if (fromPill) return fromPill;
    const fromUploaded = uploadedItems.find((x) => x.kind === "email")?.id || "";
    return fromUploaded ? String(fromUploaded) : "";
  }, [sentPill, uploadedItems]);

  const hasNonInlineAttachments = React.useMemo(() => {
    return settings.includeAttachments && attachments.filter((a) => !a.isInline).length > 0;
  }, [settings.includeAttachments, attachments]);

  const handleAutoSelectCaseId = React.useCallback((id: string) => {
    setSelectedCaseId(id);
    setSelectedSource("suggested");
  }, []);

  const { suggestions: caseSuggestions } = useCaseSuggestions({
    enabled: settings.autoSuggestCase,
    emailItemId,
    conversationKey: (email as any)?.conversationKey || "",
    subject: emailSubject,
    bodySnippet: suggestBodySnippet,
    fromEmail,
    attachments,
    cases: visibleCases,
    selectedCaseId,
    selectedSource,
    onAutoSelectCaseId: handleAutoSelectCaseId,
    topK: 3,
  });

  React.useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const hostRaw = (await getStored(STORAGE_KEYS.workspaceHost)) || "";
        const host = hostRaw.replace(/^https?:\/\//i, "").split("/")[0].trim();
        if (!cancelled) setWorkspaceHost(host);
      } catch {
        // ignore
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  React.useEffect(() => {
    let mounted = true;
    void (async () => {
      setIsLoadingCases(true);
      try {
        const [casesRes, clientsRes] = await Promise.all([
          listCases(token, settings.caseListScope),
          listClients(token),
        ]);
        if (!mounted) return;

        setCases(casesRes);

        const map: Record<string, string> = {};
        for (const c of clientsRes) map[c.id] = c.name;
        setClientNamesById(map);
      } catch (e) {
        if (!mounted) return;
        const msg = e instanceof Error ? e.message : typeof e === "string" ? e : JSON.stringify(e);
        setError(`Cases load error: ${msg}`);
      } finally {
        if (mounted) setIsLoadingCases(false);
      }
    })();
    return () => {
      mounted = false;
    };
  }, [token, settings.caseListScope]);

  React.useEffect(() => {
    let mounted = true;
    void (async () => {
      if (!emailItemId) {
        if (mounted) setSuggestBodySnippet("");
        return;
      }
      const snip = await getEmailBodySnippet(600);
      if (mounted) setSuggestBodySnippet(snip || "");
    })();
    return () => {
      mounted = false;
    };
  }, [emailItemId]);

  React.useEffect(() => {
    if (settings.rememberLastCase && selectedCaseId) saveLastCaseId(selectedCaseId);
  }, [selectedCaseId, settings.rememberLastCase]);

  const toggleAttachment = (id: string) => {
    setSelectedAttachmentIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const selectAllAttachments = () => setSelectedAttachmentIds(new Set(attachments.map((a) => a.id)));
  const clearAllAttachments = () => setSelectedAttachmentIds(new Set());

  const upsertUploaded = async (item: UploadedItem) => {
    const existing = await loadUploadedLinks(emailItemId);
    const merged = [item, ...existing.filter((x) => x.id !== item.id)].slice(0, 5);
    setUploadedItems(merged);
    await saveUploadedLinks(emailItemId, merged);
  };

  const resetLocalEvidenceForThisEmail = async () => {
    await saveUploadedLinks(emailItemId, []);
    setUploadedItems([]);
    setSentPill(null);
    setSubmitResult(null);
    saveDuplicateCache({});
  };

  const [mode, setMode] = React.useState<Mode>("idle");
  const [chat, setChat] = React.useState<ChatMsg[]>([]);
  const timersRef = React.useRef<number[]>([]);

  const clearTimers = () => {
    timersRef.current.forEach((t) => window.clearTimeout(t));
    timersRef.current = [];
  };

  const pushChat = (m: ChatMsg) => setChat((prev) => [...prev, m]);

  const later = (ms: number, fn: () => void) => {
    const t = window.setTimeout(fn, ms);
    timersRef.current.push(t);
  };

  function buildSentTimelineFrom(
    pill: SentPillData | null,
    items: UploadedItem[]
  ): Array<{ label: string; meta?: string }> {
    const out: Array<{ label: string; meta?: string }> = [];

    if (pill?.sent) {
      out.push({ label: "Zařazeno do SingleCase", meta: fmtCs(pill.atIso) });
      if (pill.caseId) out.push({ label: `Případ: ${pill.caseId}` });
      if (pill.documentId) out.push({ label: "Dokument uložen", meta: `ID ${pill.documentId}` });
      if (typeof pill.revisionNumber === "number") out.push({ label: "Verze dokumentu", meta: `v${pill.revisionNumber}` });
      if (pill.singlecaseRecordId) out.push({ label: "SingleCase záznam", meta: `${pill.singlecaseRecordId}` });
    }

    const emailDoc = items.find((x) => x.kind === "email");
    if (emailDoc) out.push({ label: "Odkaz na email v SingleCase", meta: emailDoc.name });

    const attDocs = items.filter((x) => x.kind === "attachment");
    if (attDocs.length) out.push({ label: "Nahrané přílohy", meta: `${attDocs.length}` });

    return out;
  }

  React.useEffect(() => {
    let cancelled = false;

    const run = async () => {
      try {
        setError(null);
        setSubmitResult(null);
        setDupPromptOpen(false);
        dupDecisionRef.current = null;

        if (!emailItemId) {
          if (!cancelled) {
            setUploadedItems([]);
            setAttachments([]);
            setSelectedAttachmentIds(new Set());
            setSelectedCaseId("");
            setSelectedSource("");
            setSentPill(null);
            setMode("idle");
            clearTimers();
            setChat([{ id: "idle0", kind: "system", text: "Otevři nejdřív email." }]);
          }
          return;
        }

        const disc = isDiscardedEmail(emailItemId);

        const storedLinks = await loadUploadedLinks(emailItemId);

        const results = await allSettled(
          storedLinks.slice(0, 10).map(async (it) => {
            if (!it?.id) return null;

            const metaRaw = await getDocumentMeta(it.id);
            const meta = extractDocMeta(metaRaw);
            if (!meta) return null;

            const name = (it.name || meta.name || "").trim();
            return { ...it, name } as UploadedItem;
          })
        );

        const existing = results
          .filter((r) => r.status === "fulfilled")
          .map((r) => (r as any).value)
          .filter(Boolean) as UploadedItem[];

        const seen = new Set<string>();
        const deduped: UploadedItem[] = [];
        for (const it of existing) {
          if (!it.id) continue;
          if (seen.has(it.id)) continue;
          seen.add(it.id);
          deduped.push(it);
        }

        if (cancelled) return;

        await saveUploadedLinks(emailItemId, deduped);
        if (cancelled) return;

        setUploadedItems(deduped);

        const pill = await loadSentPill(emailItemId);
        if (cancelled) return;

        setSentPill(pill);

        if (settings.rememberLastCase) {
          const last = loadLastCaseId();
          setSelectedCaseId(last || "");
          setSelectedSource(last ? "remembered" : "");
        } else {
          setSelectedCaseId("");
          setSelectedSource("");
        }

        setSelectedAttachmentIds(new Set());

        if (settings.includeAttachments) {
          setIsLoadingAttachments(true);
          try {
            const list = getCurrentOutlookAttachments()
              .filter((a) => !a.isInline)
              .sort((a, b) => a.name.localeCompare(b.name));
            if (!cancelled) setAttachments(list);
          } finally {
            if (!cancelled) setIsLoadingAttachments(false);
          }
        } else {
          setAttachments([]);
        }

        if (cancelled) return;

        clearTimers();
        setChat([]);
        setMode("idle");

        // Start chat flow
        pushChat({ id: "e0", kind: "embed", node: "email" });

        if (disc) {
          setMode("discarded");
          later(120, () => {
            pushChat({
              id: "d1",
              kind: "system",
              title: "Tento email se nebude ukládat",
              text: "Označil jsi ho jako Nechci ukládat. Nebudeme ho znovu nabízet.",
            });
            pushChat({ id: "d2", kind: "actions", row: "discardedUndo" });
          });
          return;
        }

        if (pill?.sent) {
          setMode("sent");
          const timeline = buildSentTimelineFrom(pill, deduped);

          later(120, () => {
            pushChat({ id: "s1", kind: "system", title: "Zařazeno", text: "Přehled akcí:" });
            pushChat({ id: "s2", kind: "timeline", items: timeline });
            pushChat({ id: "s3", kind: "actions", row: "sent" });
          });
          return;
        }

        setMode("new");
        later(140, () => {
          pushChat({
            id: "n1",
            kind: "system",
            title: "Navržený případ",
            text: "Vyber případ. Pak potvrď zařazení nebo zvol Nechci ukládat.",
          });
        });
        later(260, () => pushChat({ id: "n2", kind: "embed", node: "caseSelector" }));
        later(380, () => pushChat({ id: "n3", kind: "actions", row: "choose" }));
      } catch (e) {
        if (cancelled) return;
        const msg = e instanceof Error ? e.message : String(e);
        setError(`Email switch failed: ${msg}`);
        clearTimers();
        setMode("idle");
        setChat([{ id: "err0", kind: "system", title: "Chyba", text: msg }]);
      }
    };

    void run();

    return () => {
      cancelled = true;
      clearTimers();
    };
  }, [
    emailItemId,
    token,
    settings.includeAttachments,
    settings.rememberLastCase,
    settings.caseListScope,
  ]);

  const onSubmit = async () => {
    if (isSubmitting) return;

    setError(null);
    setSubmitResult(null);

    const allowResend = dupDecisionRef.current === "confirm";
    if (allowResend) {
      dupDecisionRef.current = null;
      setDupPromptOpen(false);
    }

    if (!emailItemId) {
      setError("Open an email first.");
      return;
    }

    if (!selectedCaseId) {
      setError("Select a case.");
      return;
    }

    if (settings.preventDuplicates && !allowResend) {
      try {
        const cache = loadDuplicateCache();
        const alreadyMarked = hasAttached(cache, selectedCaseId, emailItemId);
        const hasExistingDoc = Boolean(existingEmailDocId);

        if (alreadyMarked || hasExistingDoc) {
          const remoteExists = existingEmailDocId ? Boolean(await getDocumentMeta(existingEmailDocId)) : false;
          if (!remoteExists) await resetLocalEvidenceForThisEmail();
          else {
            setDupPromptOpen(true);
            return;
          }
        }
      } catch {
        // ignore
      }
    }

    setIsSubmitting(true);

    try {
      const bodySnippetFull = await getEmailBodySnippet(8000);
      const bodyForEml = (bodySnippetFull || suggestBodySnippet || "").trim() || "[No body content available from Outlook]";

      let attachmentPayloads: AttachmentPayload[] = [];

      if (settings.includeAttachments && selectedAttachmentIds.size > 0) {
        const selected = attachments.filter((a) => selectedAttachmentIds.has(a.id));
        const totalBytes = selected.reduce((sum, a) => sum + (a.size || 0), 0);
        const maxTotalBytes = 12 * 1024 * 1024;
        if (totalBytes > maxTotalBytes) {
          setError("Selected attachments are too large. Please select fewer or smaller files.");
          return;
        }

        for (const att of selected) {
          const base64 = await getAttachmentContentBase64(att.id);
          if (!base64) {
            setError(`Cannot read attachment "${att.name}". Unselect it and try again.`);
            return;
          }

          attachmentPayloads.push({
            id: att.id,
            name: att.name,
            contentType: att.contentType,
            size: att.size,
            contentBase64: base64,
          });
        }
      }

      const payload: any = {
        caseId: selectedCaseId,
        outlookItemId: emailItemId,
        subject: emailSubject,
        fromEmail,
        fromName,
      };

      if (settings.includeBodySnippet && bodySnippetFull) payload.bodySnippet = bodySnippetFull;
      if (settings.includeAttachments && attachmentPayloads.length) payload.attachments = attachmentPayloads;

      const res = await submitEmailToCase(token, payload);

      const emailText =
        `From: ${fromName} <${fromEmail}>\r\n` +
        `To: SingleCase <noreply@singlecase>\r\n` +
        `Subject: ${emailSubject}\r\n` +
        `Date: ${new Date().toUTCString()}\r\n` +
        `Message-ID: <${emailItemId}@outlook>\r\n` +
        `MIME-Version: 1.0\r\n` +
        `Content-Type: text/plain; charset=UTF-8\r\n` +
        `Content-Transfer-Encoding: 8bit\r\n` +
        `\r\n` +
        `${bodyForEml}\r\n`;

      const emailBase64 = toBase64Utf8(emailText);
      let docId = existingEmailDocId;

      if (allowResend && !docId) throw new Error("Cannot upload a new version because no existing document was found.");

      let revisionNumber = sentPill?.revisionNumber ?? 1;
      let emailDocIdForUi = "";
      const baseName = safeFileName(emailSubject);

      if (allowResend && docId) {
        const v = await uploadDocumentVersion({
          documentId: docId,
          fileName: `${baseName}.eml`,
          mimeType: "text/plain",
          dataBase64: emailBase64,
        });

        let rev = extractRevisionFromVersionUploadResponse(v);

        if (!rev) {
          for (let i = 0; i < 3; i += 1) {
            const metaRaw = await getDocumentMeta(docId);
            const meta = extractDocMeta(metaRaw);
            rev = extractLatestRevisionNumber(meta);
            if (rev) break;
            await new Promise((r) => setTimeout(r, 300));
          }
        }

        if (!rev) throw new Error(`Version upload succeeded but revision could not be determined. docId=${docId}`);

        revisionNumber = rev;
        emailDocIdForUi = docId;

        const emailUrl = buildDocumentUrl(workspaceHost, docId);
        if (emailUrl) {
          await upsertUploaded({
            id: docId,
            name: `${baseName}-v${revisionNumber}.eml`,
            url: emailUrl,
            kind: "email",
            atIso: new Date().toISOString(),
          });
        }
      } else {
        const created = await uploadDocumentToCase({
          caseId: selectedCaseId,
          fileName: `${baseName}.eml`,
          mimeType: "text/plain",
          dataBase64: emailBase64,
        });

        const newDoc: any = created.documents?.[0];
        docId = newDoc?.id ? String(newDoc.id) : docId;
        revisionNumber = newDoc?.latest_version?.revision_number ?? 1;
        emailDocIdForUi = newDoc?.id ? String(newDoc.id) : "";

        const emailUrl = buildDocumentUrl(workspaceHost, docId);
        if (docId && emailUrl) {
          await upsertUploaded({
            id: docId,
            name: revisionNumber > 1 ? `${baseName}-v${revisionNumber}.eml` : `${baseName}.eml`,
            url: emailUrl,
            kind: "email",
            atIso: new Date().toISOString(),
          });
        }
      }

      const failedUploads: Array<{ name: string; error: string }> = [];
      let uploadedCount = 0;

      if (settings.includeAttachments && attachmentPayloads.length > 0) {
        for (const att of attachmentPayloads) {
          try {
            const createdAtt = await uploadDocumentToCase({
              caseId: selectedCaseId,
              fileName: att.name,
              mimeType: att.contentType || "application/octet-stream",
              dataBase64: att.contentBase64,
            });

            const newDoc: any = createdAtt.documents?.[0];
            const newId = newDoc?.id ? String(newDoc.id) : "";
            if (!newId) throw new Error("No document id returned");

            const attUrl = buildDocumentUrl(workspaceHost, newId);
            if (attUrl) {
              await upsertUploaded({
                id: newId,
                name: att.name,
                url: attUrl,
                kind: "attachment",
                atIso: new Date().toISOString(),
              });
            }

            uploadedCount += 1;
          } catch (e) {
            failedUploads.push({ name: att.name, error: e instanceof Error ? e.message : String(e) });
          }
        }
      }

      if (settings.preventDuplicates) {
        const cache = loadDuplicateCache();
        markAttached(cache, selectedCaseId, emailItemId);
        saveDuplicateCache(cache);
      }

      recordSuccessfulAttach({
        caseId: selectedCaseId,
        conversationKey: (email as any)?.conversationKey || "",
        senderEmail: fromEmail || "",
      });

      setSubmitResult({ ok: true, id: res.singlecaseRecordId, emailDocumentId: emailDocIdForUi, uploadedCount, failedUploads });

      const pill: SentPillData = {
        sent: true,
        atIso: new Date().toISOString(),
        caseId: selectedCaseId,
        singlecaseRecordId: res.singlecaseRecordId,
        documentId: docId || undefined,
        revisionNumber,
      };

      await saveSentPill(emailItemId, pill);
      setSentPill(pill);
      markEverFiled(emailItemId);

      const timeline = buildSentTimelineFrom(pill, [
        // use the newest local upload state if you want, but uploadedItems is fine too
        // (this runs after uploads are done anyway)
        ...uploadedItems,
      ]);

      clearTimers();
      setChat((prev) => prev.filter((x) => x.kind !== "actions"));
      setMode("sent");
      pushChat({ id: `done_${Date.now()}`, kind: "system", title: "Zařazeno", text: "Přehled akcí:" });
      pushChat({ id: `tl_${Date.now()}`, kind: "timeline", items: timeline });
      pushChat({ id: `sentAct_${Date.now()}`, kind: "actions", row: "sent" });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setIsSubmitting(false);
    }
  };

  const startAttachmentsOrReady = () => {
    setChat((prev) => prev.filter((x) => x.kind !== "actions"));

    if (hasNonInlineAttachments) {
      setMode("attachments");
      pushChat({ id: `a1_${Date.now()}`, kind: "system", title: "Přílohy", text: "Vyber přílohy nebo přeskoč." });
      pushChat({ id: `a2_${Date.now()}`, kind: "embed", node: "attachments" });
      pushChat({ id: `a3_${Date.now()}`, kind: "actions", row: "attachments" });
      return;
    }

    setMode("ready");
    pushChat({ id: `r1_${Date.now()}`, kind: "system", title: "Připraveno", text: "Odešlu email do SingleCase." });
    pushChat({ id: `r2_${Date.now()}`, kind: "actions", row: "ready" });
  };

  const setDiscardedNow = () => {
    if (!emailItemId) return;
    setDiscardedEmail(emailItemId, true);
    clearTimers();
    setChat([{ id: "e0", kind: "embed", node: "email" }]);
    setMode("discarded");
    later(120, () => {
      pushChat({
        id: `d_${Date.now()}`,
        kind: "system",
        title: "Tento email se nebude ukládat",
        text: "Označeno jako Nechci ukládat. Můžeš to vrátit zpět.",
      });
      pushChat({ id: `du_${Date.now()}`, kind: "actions", row: "discardedUndo" });
    });
  };

  const undoDiscard = () => {
    if (!emailItemId) return;
    setDiscardedEmail(emailItemId, false);
    clearTimers();
    setChat([{ id: "e0", kind: "embed", node: "email" }]);
    setMode("new");
    later(140, () => pushChat({ id: "n1", kind: "system", title: "Navržený případ", text: "Vyber případ a potvrď." }));
    later(260, () => pushChat({ id: "n2", kind: "embed", node: "caseSelector" }));
    later(380, () => pushChat({ id: "n3", kind: "actions", row: "choose" }));
  };

  return (
    <SideKickCard>
      <div className="case-intake-panel-root">
        <div className="case-intake-panel-body">
          <div className="chat">
            <div className="chat-header">
              <div className="chat-header-title">Zařazení emailu</div>
              {sentPill?.sent ? (
                <StatusPill
                  label="Zařazeno"
                  tone="success"
                  version={sentPill.revisionNumber}
                  title={sentPill.caseId ? `Zařazeno do případu ${sentPill.caseId}` : "Email byl zařazen do SingleCase"}
                />
              ) : null}
            </div>

            {chat.map((m) => {
              if (m.kind === "system") {
                return (
                  <div key={m.id} className="chat-bubble chat-in">
                    {m.title ? <div className="chat-title">{m.title}</div> : null}
                    {m.text ? <div className="chat-text">{m.text}</div> : null}
                  </div>
                );
              }

              if (m.kind === "embed" && m.node === "email") {
                return (
                  <div key={m.id} className="chat-bubble chat-bubble-embed chat-in">
                    <EmailContextPanel
                      emailError={emailError}
                      emailItemId={emailItemId}
                      fromName={fromName}
                      fromEmail={fromEmail}
                      subject={emailSubject}
                      isSwitchingEmail={false}
                    />
                  </div>
                );
              }

              if (m.kind === "embed" && m.node === "caseSelector") {
                return (
                  <div key={m.id} className="chat-bubble chat-bubble-embed chat-in">
                    <CaseSelector
                      title="Case"
                      scope={settings.caseListScope}
                      onScopeChange={(scope) => {
                        onChangeSettings((prev) => {
                          if (prev.caseListScope === scope) return prev;
                          return { ...prev, caseListScope: scope };
                        });
                      }}                      selectedCaseId={selectedCaseId}
                      onSelectCaseId={(id) => {
                        setSelectedCaseId(id);
                        setSelectedSource("manual");
                      }}
                      suggestions={caseSuggestions}
                      cases={visibleCases}
                      isLoadingCases={isLoadingCases}
                      clientNamesById={clientNamesById}
                    />
                  </div>
                );
              }

              if (m.kind === "embed" && m.node === "attachments") {
                return (
                  <div key={m.id} className="chat-bubble chat-bubble-embed chat-in">
                    <AttachmentsPicker
                      enabled={settings.includeAttachments}
                      attachments={attachments}
                      selectedAttachmentIds={selectedAttachmentIds}
                      isLoadingAttachments={isLoadingAttachments}
                      onToggleAttachment={toggleAttachment}
                      onSelectAll={selectAllAttachments}
                      onClearAll={clearAllAttachments}
                    />
                  </div>
                );
              }

              if (m.kind === "timeline") {
                return (
                  <div key={m.id} className="chat-bubble chat-in">
                    <div className="timeline">
                      {m.items.map((it, idx) => (
                        <div key={`${m.id}_${idx}`} className="timeline-row">
                          <div className="timeline-dot" />
                          <div className="timeline-body">
                            <div className="timeline-label">{it.label}</div>
                            {it.meta ? <div className="timeline-meta">{it.meta}</div> : null}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              }

              if (m.kind === "actions" && m.row === "choose") {
                return (
                  <div key={m.id} className="chat-actions chat-in">
                    <button
                      type="button"
                      className="chat-btn chat-btn-primary"
                      disabled={!selectedCaseId}
                      onClick={startAttachmentsOrReady}
                    >
                      Ano, zatřídit
                    </button>

                    <button type="button" className="chat-btn chat-btn-danger" onClick={setDiscardedNow}>
                      Nechci ukládat
                    </button>
                  </div>
                );
              }

              if (m.kind === "actions" && m.row === "attachments") {
                return (
                  <div key={m.id} className="chat-actions chat-actions-right chat-in">
                    <button
                      type="button"
                      className="chat-btn"
                      onClick={() => {
                        setSelectedAttachmentIds(new Set());
                        setChat((prev) => prev.filter((x) => x.kind !== "actions"));
                        setMode("ready");
                        pushChat({ id: `r1_${Date.now()}`, kind: "system", title: "Připraveno", text: "Odešlu email do SingleCase." });
                        pushChat({ id: `r2_${Date.now()}`, kind: "actions", row: "ready" });
                      }}
                    >
                      Přeskočit
                    </button>

                    <button
                      type="button"
                      className="chat-btn chat-btn-primary"
                      onClick={() => {
                        setChat((prev) => prev.filter((x) => x.kind !== "actions"));
                        setMode("ready");
                        pushChat({ id: `r1_${Date.now()}`, kind: "system", title: "Připraveno", text: "Odešlu email a vybrané přílohy do SingleCase." });
                        pushChat({ id: `r2_${Date.now()}`, kind: "actions", row: "ready" });
                      }}
                    >
                      Continue
                    </button>
                  </div>
                );
              }

              if (m.kind === "actions" && m.row === "ready") {
                return (
                  <div key={m.id} className="chat-actions chat-actions-right chat-in">
                    <button
                      type="button"
                      className="chat-btn chat-btn-primary"
                      disabled={isSubmitting || !emailItemId || !selectedCaseId}
                      onClick={onSubmit}
                    >
                      {isSubmitting ? "Sending..." : "Send to SingleCase"}
                    </button>
                  </div>
                );
              }

              if (m.kind === "actions" && m.row === "discardedUndo") {
                return (
                  <div key={m.id} className="chat-actions chat-in">
                    <button type="button" className="chat-btn" onClick={undoDiscard}>
                      Vrátit zpět
                    </button>
                  </div>
                );
              }

              if (m.kind === "actions" && m.row === "sent") {
                const caseLine = selectedCaseDisplay
                  ? `${selectedCaseDisplay.caseLabel} (${selectedCaseDisplay.clientLabel})`
                  : sentPill?.caseId || "";

                return (
                  <div key={m.id} className="chat-actions chat-in">
                    {caseLine ? <div className="chat-sent-note">Zařazeno do: {caseLine}</div> : null}
                  </div>
                );
              }

              return null;
            })}

            {error ? <div className="chat-error chat-in">{error}</div> : null}

            {dupPromptOpen ? (
              <div className="chat-bubble chat-in">
                <div className="dup-prompt">
                  <div className="dup-text">
                    Tento email už v případu existuje. Chceš ho nahrát znovu jako novou verzi?
                  </div>
                  <div className="dup-actions">
                    <button
                      type="button"
                      className="dup-cancel"
                      onClick={() => setDupPromptOpen(false)}
                      disabled={isSubmitting}
                    >
                      Zrušit
                    </button>
                    <button
                      type="button"
                      className="dup-confirm"
                      disabled={isSubmitting}
                      onClick={async () => {
                        dupDecisionRef.current = "confirm";
                        await onSubmit();
                      }}
                    >
                      Nahrát novou verzi
                    </button>
                  </div>
                </div>
              </div>
            ) : null}

            {submitResult?.ok ? (
              <div className="chat-bubble chat-in">
                <div className="chat-title">Hotovo</div>
                <div className="chat-text">
                  Zařazeno. {submitResult.uploadedCount ? `Přílohy nahrány: ${submitResult.uploadedCount}.` : ""}
                  {submitResult.failedUploads?.length ? ` Některé přílohy selhaly: ${submitResult.failedUploads.map((x) => x.name).join(", ")}.` : ""}
                </div>
              </div>
            ) : null}
          </div>
        </div>

        <div className="case-intake-panel-footer">
          {uploadedItems.length > 0 ? (
            <div className="case-intake-panel-uploaded">
              <div className="case-intake-panel-uploaded-title">Nahráno</div>
              <div className="case-intake-panel-uploaded-list">
                {uploadedItems.slice(0, 2).map((it) => (
                  <div key={it.id} className="case-intake-panel-uploaded-item">
                    <div className="case-intake-panel-uploaded-row">
                      <a className="case-intake-panel-uploaded-link" href={it.url} target="_blank" rel="noreferrer" title={it.name}>
                        {it.name}
                      </a>
                      <div className="case-intake-panel-uploaded-actions">
                        <a href={it.url} target="_blank" rel="noreferrer" title="View">
                          ⏿
                        </a>
                      </div>
                    </div>
                    <span className="case-intake-panel-uploaded-timestamp">
                      {it.atIso ? new Date(it.atIso).toLocaleString("cs-CZ") : ""}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          ) : null}
        </div>
      </div>
    </SideKickCard>
  );
}