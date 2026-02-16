// src/services/graphAuth.ts
declare const Office: any;

const DIALOG_URL = `${window.location.origin}/dialog.html?mode=graph`;
const DIALOG_OPTIONS = { height: 60, width: 30, displayInIframe: true };

type GraphAuthOk = {
  type: "graph_auth_complete";
  accessToken: string;
  account?: { username?: string; homeAccountId?: string };
};

type GraphAuthCancel = { type: "graph_auth_cancel" };

function safeJsonParse(s: string): any {
  try {
    return JSON.parse(s);
  } catch {
    return null;
  }
}

export async function getGraphToken(): Promise<string> {
  return new Promise((resolve, reject) => {
    try {
      if (!Office?.context?.ui?.displayDialogAsync) {
        reject(new Error("Office dialog API not available."));
        return;
      }

      Office.context.ui.displayDialogAsync(DIALOG_URL, DIALOG_OPTIONS, (asyncResult: any) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(asyncResult.error?.message || "Failed to open auth dialog."));
          return;
        }

        const dialog = asyncResult.value;

        const cleanup = () => {
          try {
            dialog.removeEventHandler(Office.EventType.DialogMessageReceived, onMsg);
          } catch {
            // ignore
          }
          try {
            dialog.removeEventHandler(Office.EventType.DialogEventReceived, onEvt);
          } catch {
            // ignore
          }
          try {
            dialog.close();
          } catch {
            // ignore
          }
        };

        const onEvt = (evt: any) => {
          cleanup();
          reject(new Error(evt?.error?.message || "Graph auth dialog closed."));
        };

        const onMsg = (arg: any) => {
          const data = safeJsonParse(String(arg?.message || ""));
          if (!data?.type) return;

          if (data.type === "graph_auth_cancel") {
            cleanup();
            reject(new Error("User cancelled Graph login."));
            return;
          }

          if (data.type === "graph_auth_complete") {
            const payload = data as GraphAuthOk;
            const token = String(payload.accessToken || "");
            if (!token) {
              cleanup();
              reject(new Error("Graph token missing in dialog response."));
              return;
            }
            cleanup();
            resolve(token);
          }
        };

        try {
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, onMsg);
          dialog.addEventHandler(Office.EventType.DialogEventReceived, onEvt);
        } catch (e) {
          cleanup();
          reject(e instanceof Error ? e : new Error(String(e)));
        }
      });
    } catch (e) {
      reject(e instanceof Error ? e : new Error(String(e)));
    }
  });
}