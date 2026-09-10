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
  | "shape"
  /** a tag the team typed in Word: <c>, <cc>, <t> (unnumbered) or <c3>, <t1> */
  | "marker"
  /** a list paragraph whose bullet glyph is a box (☐ from numbering.xml) */
  | "bullet";

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
  /** kind "marker" with a number typed in Word ("c7"): kept unless renumbering */
  fixedId?: string;
}

export interface DocxAnalysis extends EsignAnalysis {
  candidates: DocxCandidate[];
}

export interface DocxAssignment {
  type: CandidateType;
  /** control id, e.g. "c1" / "t2" (checked flag is added from `type`) */
  id: string;
}

export interface AssignOptions {
  /** ignore the numbers typed in Word and number every tag afresh in
   *  document order (c1, c2 … / t1, t2 …) */
  renumber?: boolean;
}

/**
 * Give every non-ignored candidate a control id in document order.
 *
 * Default: tags the team already typed with a number (<c7>) keep it and
 * everything new continues after the highest number of that type.
 * With `renumber`: all tags — typed or detected — are numbered afresh in
 * the order they appear, so a tag inserted in the middle shifts the rest.
 */
export function assignControlIds(
  candidates: DocxCandidate[],
  types: Record<string, CandidateType | undefined>,
  analysis: Pick<EsignAnalysis, "controls">,
  options: AssignOptions = {},
): Record<string, DocxAssignment> {
  const renumber = options.renumber === true;
  const fixedNumbers = (letter: "c" | "t") =>
    candidates
      .filter((x) => x.fixedId?.startsWith(letter))
      .map((x) => parseInt((x.fixedId as string).slice(1), 10) || 0);
  let c = renumber
    ? 0
    : Math.max(
        0,
        ...analysis.controls.filter((x) => x.type === "checkbox").map((x) => x.n),
        ...fixedNumbers("c"),
      );
  let t = renumber
    ? 0
    : Math.max(
        0,
        ...analysis.controls.filter((x) => x.type === "textbox").map((x) => x.n),
        ...fixedNumbers("t"),
      );
  const out: Record<string, DocxAssignment> = {};
  for (const cand of candidates) {
    const type = types[cand.id] ?? cand.suggested ?? "ignore";
    if (type === "ignore") {
      out[cand.id] = { type, id: "" };
      continue;
    }
    const wantsText = type === "textbox";
    const keepTyped =
      !renumber &&
      cand.fixedId &&
      cand.fixedId.startsWith(wantsText ? "t" : "c");
    if (keepTyped) {
      out[cand.id] = { type, id: cand.fixedId as string };
    } else if (wantsText) {
      out[cand.id] = { type, id: `t${++t}` };
    } else {
      out[cand.id] = { type, id: `c${++c}` };
    }
  }
  return out;
}
