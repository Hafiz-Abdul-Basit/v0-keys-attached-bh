/* Types + numbering helper shared by the DOCX engine (server) and the
 * Esign Templates page (browser). Kept free of pizzip so the client
 * bundle does not pull the zip library in. */

import type { EsignAnalysis } from "./html-template";

export type CandidateKind =
  | "contentControl"
  | "legacyField"
  | "symbol"
  | "unicodeBox"
  | "bracketBox"
  | "blank"
  | "image"
  | "shape";

export type CandidateType = "checkbox" | "checkboxChecked" | "textbox" | "ignore";

export interface DocxCandidate {
  /** stable id, e.g. "document-3" — same on analyze and build */
  id: string;
  part: string;
  kind: CandidateKind;
  suggested: CandidateType | null;
  /** short human label, e.g. "Word checkbox (checked)" */
  label: string;
  /** paragraph text with ▮ where the candidate sits */
  context: string;
  /** data: URI for raster pictures */
  thumbnail?: string;
  widthIn?: number;
  heightIn?: number;
  /** the literal text that was found ("[ ]", "_____", shape text …) */
  text?: string;
}

export interface DocxAnalysis extends EsignAnalysis {
  candidates: DocxCandidate[];
}

export interface DocxAssignment {
  type: CandidateType;
  /** control id, e.g. "c1" / "t2" (checked flag is added from `type`) */
  id: string;
}

/**
 * Give every non-ignored candidate a control id in document order,
 * continuing after the highest <cN>/<tN> already typed in the file.
 */
export function assignControlIds(
  candidates: DocxCandidate[],
  types: Record<string, CandidateType | undefined>,
  analysis: Pick<EsignAnalysis, "controls">,
): Record<string, DocxAssignment> {
  let c = Math.max(
    0,
    ...analysis.controls.filter((x) => x.type === "checkbox").map((x) => x.n),
  );
  let t = Math.max(
    0,
    ...analysis.controls.filter((x) => x.type === "textbox").map((x) => x.n),
  );
  const out: Record<string, DocxAssignment> = {};
  for (const cand of candidates) {
    const type = types[cand.id] ?? cand.suggested ?? "ignore";
    if (type === "ignore") {
      out[cand.id] = { type, id: "" };
    } else if (type === "textbox") {
      out[cand.id] = { type, id: `t${++t}` };
    } else {
      out[cand.id] = { type, id: `c${++c}` };
    }
  }
  return out;
}
