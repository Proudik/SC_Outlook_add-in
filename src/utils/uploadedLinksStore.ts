import { getStored, setStored } from "./storage";

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
  const raw = await getStored(key(emailItemId));
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
  await setStored(key(emailItemId), JSON.stringify(items || []));
}
