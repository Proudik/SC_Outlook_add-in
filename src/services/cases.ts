import { scGet } from "./client";

export type CaseOption = {
  id: string;
  label: string;
};

type RawCase = any;

function mapCase(c: RawCase): CaseOption {
  return {
    id: String(c.id),
    label: c.name || c.case_id_visible || `Case ${c.id}`,
  };
}

export async function fetchCasesAll(filters?: {
  name?: string;
  client_id?: number | string;
}) {
  const data = await scGet<RawCase[]>("/cases", filters);
  return data.map(mapCase);
}
