import { useCallback, useEffect, useRef, useState } from "react";

export type EmailMode = "read" | "compose";

export type EmailContext = {
  itemId: string;
  itemKey: string;
  mode: EmailMode;
  subject: string;
  fromName?: string;
  fromEmail?: string;
  conversationKey?: string;
};

function safeSucceeded(res: any): boolean {
  const status = res?.status;
  const officeSucceeded = (globalThis as any)?.Office?.AsyncResultStatus?.Succeeded;

  if (officeSucceeded !== undefined) return status === officeSucceeded;

  // Fallback for some hosts that return "succeeded"
  return String(status || "").toLowerCase() === "succeeded";
}

function officeReady(): Promise<void> {
  return new Promise((resolve) => {
    try {
      const OfficeAny = (globalThis as any)?.Office;
      if (OfficeAny && typeof OfficeAny.onReady === "function") {
        OfficeAny.onReady(() => resolve());
      } else {
        resolve();
      }
    } catch {
      resolve();
    }
  });
}

function getMode(item: any): EmailMode {
  const subj = item?.subject;
  if (subj && typeof subj === "object" && typeof subj.getAsync === "function") return "compose";
  return "read";
}

function getUserProfile(): { name?: string; email?: string } {
  try {
    const OfficeAny = (globalThis as any)?.Office;
    const p = OfficeAny?.context?.mailbox?.userProfile;
    return {
      name: p?.displayName ? String(p.displayName) : undefined,
      email: p?.emailAddress ? String(p.emailAddress) : undefined,
    };
  } catch {
    return {};
  }
}

function getSubjectAsync(item: any, mode: EmailMode): Promise<string> {
  if (mode === "read") return Promise.resolve(String(item?.subject || ""));

  return new Promise((resolve) => {
    try {
      const fn = item?.subject?.getAsync;
      if (typeof fn !== "function") {
        resolve("");
        return;
      }

      fn.call(item.subject, (res: any) => {
        try {
          if (safeSucceeded(res)) resolve(String(res?.value || ""));
          else resolve("");
        } catch {
          resolve("");
        }
      });
    } catch {
      resolve("");
    }
  });
}

function getConversationKeyAsync(item: any, itemId: string, mode: EmailMode): Promise<string> {
  const cid = String(item?.conversationId || "").trim();
  if (cid) return Promise.resolve(`cid:${cid}`);

  const getIdx = item?.getConversationIndexAsync;
  if (mode === "read" && typeof getIdx === "function") {
    return new Promise((resolve) => {
      try {
        getIdx((res: any) => {
          try {
            if (safeSucceeded(res)) {
              const v = String(res?.value || "").trim();
              if (v) resolve(`cidx:${v}`);
              else resolve(itemId ? `item:${itemId}` : "");
            } else {
              resolve(itemId ? `item:${itemId}` : "");
            }
          } catch {
            resolve(itemId ? `item:${itemId}` : "");
          }
        });
      } catch {
        resolve(itemId ? `item:${itemId}` : "");
      }
    });
  }

  return Promise.resolve(itemId ? `item:${itemId}` : "");
}

export function useOutlookSuggestions() {
  const [email, setEmail] = useState<EmailContext | null>(null);
  const [suggestions, setSuggestions] = useState<any[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const lastKeyRef = useRef<string>("");
  const handlerRef = useRef<((arg?: any) => void) | null>(null);
  const runningRef = useRef(false);

  const refresh = useCallback(async () => {
    // Avoid overlapping refresh calls during rapid item switches
    if (runningRef.current) return;
    runningRef.current = true;

    try {
      setError(null);

      const OfficeAny = (globalThis as any)?.Office;
      const item = OfficeAny?.context?.mailbox?.item as any;
      if (!item) return;

      const mode = getMode(item);

      const itemId = String(item?.itemId || "");

      const fallbackKey =
        `compose:${String(item?.conversationId || "")}:${String(item?.dateTimeCreated || "")}`.trim();

      const itemKey = itemId || fallbackKey || `compose:${Date.now()}`;

      // If Outlook re fires ItemChanged for the same item, skip heavy work
      if (itemKey && itemKey === lastKeyRef.current) return;
      lastKeyRef.current = itemKey;

      const subject = await getSubjectAsync(item, mode);

      let fromName = "";
      let fromEmail = "";

      if (mode === "read") {
        fromName = String(item?.from?.displayName || "");
        fromEmail = String(item?.from?.emailAddress || item?.from?.address || "");
      } else {
        const up = getUserProfile();
        fromName = String(up.name || "");
        fromEmail = String(up.email || "");
      }

      const conversationKey = await getConversationKeyAsync(item, itemId, mode);

      setEmail({ itemId, itemKey, mode, subject, fromName, fromEmail, conversationKey });

      setIsLoading(true);
      setSuggestions([]);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      setSuggestions([]);
    } finally {
      setIsLoading(false);
      runningRef.current = false;
    }
  }, []);

  useEffect(() => {
    let cancelled = false;

    void (async () => {
      await officeReady();
      if (cancelled) return;

      await refresh();
      if (cancelled) return;

      const OfficeAny = (globalThis as any)?.Office;
      const mailbox = OfficeAny?.context?.mailbox;

      if (!mailbox?.addHandlerAsync) return;

      const handler = () => {
        try {
          void refresh();
        } catch {
          // never throw
        }
      };

      handlerRef.current = handler;

      try {
        // Use callback signature if host expects it
        mailbox.addHandlerAsync(OfficeAny.EventType.ItemChanged, handler, () => {
          // ignore result
        });
      } catch {
        // ignore
      }
    })();

    return () => {
      cancelled = true;

      try {
        const OfficeAny = (globalThis as any)?.Office;
        const mailbox = OfficeAny?.context?.mailbox;
        const handler = handlerRef.current;

        if (mailbox?.removeHandlerAsync && handler) {
          mailbox.removeHandlerAsync(OfficeAny.EventType.ItemChanged, { handler }, () => {
            // ignore
          });
        }
      } catch {
        // ignore
      }

      handlerRef.current = null;
      lastKeyRef.current = "";
      runningRef.current = false;
    };
  }, [refresh]);

  return { email, suggestions, isLoading, error, refresh };
}
