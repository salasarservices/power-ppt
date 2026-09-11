import axios from "axios";

// Types mirror power-ppt-api app/schemas/models.py exactly.
export interface Table {
  header: string[];
  rows: string[][]; // includes the header as rows[0]
}

export interface PptImage {
  data: string; // base64-encoded image bytes
  content_type: string;
}

export interface Page {
  title: string; // "" -> the engine renders no heading
  body: string; // paragraphs separated by "\n\n"
  tables: Table[];
  images?: PptImage[];
}

export interface SlidePlan {
  pages: Page[];
}

export interface AnalyzeResponse {
  slides: number;
  warnings: string[];
  plan: SlidePlan;
}

// Layout (Placement JSON) — positioned output of the flow engine, the editor's
// round-trip contract. Positions are in inches on a 13.33 x 7.5 slide.
export type PlacementKind = "body" | "table" | "image";

export interface Placement {
  kind: PlacementKind;
  left: number;
  top: number;
  width: number;
  height?: number | null;
  text?: string | null;
  table?: Table | null;
  image?: PptImage | null;
}

export interface DeckSlide {
  title: string;
  placements: Placement[];
}

export interface Deck {
  slides: DeckSlide[];
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

/** Flow a reviewed SlidePlan into a positioned Deck (Placement JSON). */
export async function layoutPlan(plan: SlidePlan): Promise<Deck> {
  const { data } = await client.post<Deck>("/layout", plan);
  return data;
}

/** Render the edited Deck (Placement JSON) into the .pptx (WYSIWYG). */
export async function renderDeck(deck: Deck): Promise<Blob> {
  const { data } = await client.post<Blob>("/render", deck, {
    responseType: "blob",
  });
  return data;
}

// Brand constants for the canvas (mirror app/brand_engine/geometry.py).
export const SLIDE = { wIn: 13.33, hIn: 7.5 };
export const HEADING = { left: 0.75, top: 0.55, width: 9.5, height: 0.9 };
export const BRAND = {
  blue: "#1A3A8F",
  green: "#7AC143",
  slate: "#4D4D4D",
  headingPt: 20,
  bodyPt: 12,
};
