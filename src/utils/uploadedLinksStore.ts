// uploadedLinksStore uses localStorage directly (not setStored/getStored) because:
// - These are display-only UI links for FiledSummaryCard — no cross-device sync needed
// - setStored falls back to roamingSettings in OWA (OfficeRuntime.storage is Desktop-only),
//   and accumulating one key per filed email quickly blows the 32KB roamingSettings limit.

type UploadedItem = {
  id: string;
  name: string;
  url: string;
  kind: "email" | "attachment";
  atIso: string;
};

function key(emailItemId: string) {
  return `sc:uploadedLinks:${emailItemId}`;
}

export async function loadUploadedLinks(emailItemId: string): Promise<UploadedItem[]> {
  if (!emailItemId) return [];
  const raw = (typeof localStorage !== "undefined" ? localStorage.getItem(key(emailItemId)) : null);
  if (!raw) return [];

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];

    return parsed
      .filter((x) => x && typeof x.id === "string" && typeof x.url === "string")
      .map((x) => ({
        id: String(x.id),
        name: typeof x.name === "string" ? x.name : "",
        url: String(x.url),
        kind: x.kind === "attachment" ? "attachment" : "email",
        atIso: typeof x.atIso === "string" ? x.atIso : "",
      }));
  } catch {
    return [];
  }
}

export async function saveUploadedLinks(emailItemId: string, items: UploadedItem[]): Promise<void> {
  if (!emailItemId) return;
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(key(emailItemId), JSON.stringify(items || []));
    }
  } catch {
    // localStorage full or unavailable — silently ignore, links are display-only
  }
}
