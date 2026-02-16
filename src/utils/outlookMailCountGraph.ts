type Params = {
    filedCategoryName: string;
    maxPagesPerFolder: number; // defensive cap
  };
  
  async function getGraphAccessToken(): Promise<string> {
    const OfficeAny: any = Office as any;
    if (!OfficeAny?.auth?.getAccessToken) {
      throw new Error("Office.auth.getAccessToken is not available in this host.");
    }
  
    return new Promise((resolve, reject) => {
      OfficeAny.auth.getAccessToken(
        {
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        },
        (result: any) => {
          if (result?.status === "succeeded" && result?.value) {
            resolve(String(result.value));
            return;
          }
  
          const code = result?.error?.code ?? "unknown";
          const msg = result?.error?.message ?? "Unknown error";
          reject(new Error(`getAccessToken failed. code=${code} message=${msg}`));
        }
      );
    });
  }
  
  async function graphCountNotCategorised(folderId: string, categoryName: string, maxPages: number): Promise<number> {
    const token = await getGraphAccessToken();
  
    // We use $count=true + ConsistencyLevel header.
    // Filter: NOT categories/any(c:c eq '...')
    const filter = `not(categories/any(c:c eq '${categoryName.replace(/'/g, "''")}'))`;
  
    let url =
      `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages` +
      `?$select=id&$top=1&$count=true&$filter=${encodeURIComponent(filter)}`;
  
    let total = 0;
    let pages = 0;
  
    while (url && pages < maxPages) {
      pages += 1;
  
      const res = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          ConsistencyLevel: "eventual",
        },
      });
  
      if (!res.ok) {
        const body = await res.text().catch(() => "");
        throw new Error(`Graph request failed (${res.status}). ${body}`);
      }
  
      const data: any = await res.json();
  
      // @odata.count contains the total count for the query (not just the page size)
      const count = Number(data["@odata.count"]);
      if (Number.isFinite(count)) {
        total = count;
        break;
      }
  
      // Fallback if count missing for some reason
      url = data["@odata.nextLink"] ? String(data["@odata.nextLink"]) : "";
    }
  
    return total;
  }
  
  export async function countEmailsToFileGraph(params: Params): Promise<number> {
    // Inbox + SentItems
    const inbox = await graphCountNotCategorised("Inbox", params.filedCategoryName, params.maxPagesPerFolder);
    const sent = await graphCountNotCategorised("SentItems", params.filedCategoryName, params.maxPagesPerFolder);
    return inbox + sent;
  }