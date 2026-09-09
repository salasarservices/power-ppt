import axios from "axios";

// Types mirror power-ppt-api app/schemas/models.py exactly.
export interface Table {
  header: string[];
  rows: string[][]; // includes the header as rows[0]
}

export interface Page {
  title: string; // "" -> the engine renders no heading
  body: string; // paragraphs separated by "\n\n"
  tables: Table[];
}

export interface SlidePlan {
  pages: Page[];
}

export interface AnalyzeResponse {
  slides: number;
  warnings: string[];
  plan: SlidePlan;
}

export interface Health {
  status: string;
  template_version: string;
}

// Single-service topology: the SPA is served by the same FastAPI process that
// exposes the API, so calls are same-origin (baseURL ""). In dev, Vite proxies
// the API routes to the local backend on :8077 (see vite.config.ts).
const client = axios.create({ baseURL: "" });

export async function getHealth(): Promise<Health> {
  const { data } = await client.get<Health>("/health");
  return data;
}

export async function analyzePptx(
  file: File,
  opts: { useOcr?: boolean; alwaysOcr?: boolean } = {},
): Promise<AnalyzeResponse> {
  const form = new FormData();
  form.append("file", file);
  const params: Record<string, string> = {};
  if (opts.useOcr !== undefined) params.use_ocr = String(opts.useOcr);
  if (opts.alwaysOcr) params.always_ocr = "true";

  const { data } = await client.post<AnalyzeResponse>("/analyze", form, {
    params,
  });
  return data;
}

/** POST the reviewed plan; returns the generated .pptx as a Blob to download. */
export async function generateDeck(plan: SlidePlan): Promise<Blob> {
  const { data } = await client.post<Blob>("/generate", plan, {
    responseType: "blob",
  });
  return data;
}
