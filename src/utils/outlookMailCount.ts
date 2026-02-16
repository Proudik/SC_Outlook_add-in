/* Folder: src/utils/outlookMailCount.ts */

type CountArgs = {
    filedCategoryName: string;
    maxToScanPerFolder?: number;
  };
  
  const GRAPH_TOKEN_KEY = "sc_graph_access_token_v1";
  
  function saveGraphToken(token: string) {
    try {
      window.localStorage.setItem(GRAPH_TOKEN_KEY, token);
    } catch {
      // ignore
    }
  }
  
  function loadGraphToken(): string {
    try {
      return String(window.localStorage.getItem(GRAPH_TOKEN_KEY) || "").trim();
    } catch {
      return "";
    }
  }
  
  export function setGraphAccessToken(token: string) {
    const t = String(token || "").trim();
    if (t) saveGraphToken(t);
  }
  
  export function clearGraphAccessToken() {
    try {
      window.localStorage.removeItem(GRAPH_TOKEN_KEY);
    } catch {
      // ignore
    }
  }
  
  function encodeODataString(value: string): string {
    return String(value || "").replace(/'/g, "''");
  }
  
  async function graphGetJson<T>(url: string, accessToken: string): Promise<T> {
    const resp = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });
  
    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      throw new Error(`Graph ${resp.status} ${resp.statusText}. ${text}`.trim());
    }
  
    return (await resp.json()) as T;
  }
  
  async function countUnfiledInFolder(args: {
    folderWellKnownName: "inbox" | "sentitems";
    filedCategoryName: string;
    maxToScan: number;
    accessToken: string;
  }): Promise<number> {
    const { folderWellKnownName, filedCategoryName, maxToScan, accessToken } = args;
  
    const cat = encodeODataString(filedCategoryName);
  
    let count = 0;
    let scanned = 0;
  
    // We request categories and then double check in JS (extra safety).
    // Filter is optional but speeds it up when supported.
    let nextUrl =
      `https://graph.microsoft.com/v1.0/me/mailFolders/${folderWellKnownName}/messages` +
      `?$select=id,categories,receivedDateTime` +
      `&$top=50` +
      `&$orderby=receivedDateTime desc` +
      `&$filter=not(categories/any(c:c eq '${cat}'))`;
  
    while (nextUrl && scanned < maxToScan) {
      const data: any = await graphGetJson<any>(nextUrl, accessToken);
  
      const items: any[] = Array.isArray(data?.value) ? data.value : [];
      scanned += items.length;
  
      for (const m of items) {
        const cats: string[] = Array.isArray(m?.categories) ? m.categories : [];
        const isFiled = cats.some(
          (x) => String(x).trim().toLowerCase() === filedCategoryName.trim().toLowerCase()
        );
        if (!isFiled) count += 1;
      }
  
      const nextLink: string | undefined = data?.["@odata.nextLink"];
      if (!nextLink) break;
      nextUrl = nextLink;
    }
  
    return count;
  }
  
  export async function countEmailsToFile(args: CountArgs): Promise<number> {
    const filedCategoryName = String(args.filedCategoryName || "").trim();
    if (!filedCategoryName) throw new Error("filedCategoryName is required.");
  
    const accessToken = loadGraphToken();
    if (!accessToken) {
      // Caller should open MSAL dialog and then call setGraphAccessToken(token)
      throw new Error("GRAPH_AUTH_REQUIRED");
    }
  
    const maxToScan = Math.max(50, Math.min(1000, Number(args.maxToScanPerFolder || 200)));
  
    const [inbox, sent] = await Promise.all([
      countUnfiledInFolder({ folderWellKnownName: "inbox", filedCategoryName, maxToScan, accessToken }),
      countUnfiledInFolder({ folderWellKnownName: "sentitems", filedCategoryName, maxToScan, accessToken }),
    ]);
  
    return inbox + sent;
  }