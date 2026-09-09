/* ──────────────────────────────────────────────────────────────────────
 * Esign-Templates / Customized — DOCX engine
 *
 * The real esign workflow starts with the client's Word file. Inside it
 * the checkboxes / textboxes are never defined the same way twice: a
 * Word checkbox content control, a legacy form field, a Wingdings box
 * symbol, a "[ ]" typed by hand, a row of underscores, a tiny picture
 * of a box, a drawn text-box shape… This engine
 *
 *   1. analyses the .docx and lists every such *candidate* with a
 *      suggested type (checkbox / checked checkbox / textbox), a text
 *      snippet showing where it sits and, for pictures, a thumbnail;
 *   2. also discovers placeholder keys (<<KEY>>, @KEY, bare KEY) and
 *      any markers the team already typed (<c1>, <c1c>, <t1>);
 *   3. builds the *mapped* .docx: each assigned candidate becomes the
 *      literal marker text (<c1>, <c2c>, <t1> …) and every key is
 *      rewritten to its canonical <<KEY>> — while keeping run formatting
 *      intact (text is spliced inside the existing <w:t> runs, never
 *      merged into one run).
 *
 * Converting the mapped .docx to HTML happens in the browser (see
 * docx-to-html.client.ts) and the HTML engine then turns the markers into
 * real form controls. Nothing here is shared with /api/replace.
 * ────────────────────────────────────────────────────────────────────── */

import PizZip from "pizzip";
import {
  analyzeEsignText,
  decodeEntities,
  hasMappingFor,
  keyTargetSpecs,
  normalizeKey,
  resolveToken,
  type AnalyzeOptions,
  type EsignAnalysis,
  type EsignTokenForm,
} from "./html-template";

import {
  assignControlIds,
  type CandidateKind,
  type CandidateType,
  type DocxAnalysis,
  type DocxAssignment,
  type DocxCandidate,
} from "./docx-shared";

export { assignControlIds };
export type {
  CandidateKind,
  CandidateType,
  DocxAnalysis,
  DocxAssignment,
  DocxCandidate,
};

/* ───────────────────────────── types ───────────────────────────── */

export interface DocxBuildOptions {
  assignments?: Record<string, DocxAssignment>;
  mappings?: Record<string, string>;
  /** raw tokens the user chose NOT to replace */
  exclude?: string[];
  replaceKeys?: boolean;
  /** also rewrite key names found as plain text (default true) */
  plainText?: boolean;
}

export interface DocxBuildStats {
  keys: Record<string, number>;
  /** candidate id → marker written, e.g. "document-3": "c1c" */
  markers: Record<string, string>;
  skipped: string[];
}

export interface DocxBuildResult {
  bytes: Uint8Array;
  stats: DocxBuildStats;
}

/* ─────────────────────────── xml helpers ─────────────────────────── */

const PART_RE = /^word\/(document|header\d*|footer\d*)\.xml$/;
const EMU_PER_INCH = 914400;

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function partRank(name: string): number {
  if (/document\.xml$/.test(name)) return 0;
  if (/header/.test(name)) return 1;
  return 2;
}

function partKey(name: string): string {
  return name.replace(/^word\//, "").replace(/\.xml$/, "");
}

export function listDocxParts(zip: PizZip): string[] {
  return Object.keys(zip.files)
    .filter((f) => PART_RE.test(f))
    .sort((a, b) => partRank(a) - partRank(b) || a.localeCompare(b));
}

/** innermost <w:p> elements (no nested <w:p> inside) */
const PARAGRAPH_RE = /<w:p(?:\s[^>]*)?>(?:(?!<w:p[\s>])[\s\S])*?<\/w:p>/g;
/** innermost <w:r> elements */
const RUN_RE = /<w:r(?:\s[^>]*)?>(?:(?!<w:r[\s>])[\s\S])*?<\/w:r>/g;
/** <w:t> (not <w:tab/>, <w:tbl> …) */
const WT_RE = /<w:t(?![a-zA-Z])([^>]*?)(?:\/>|>([\s\S]*?)<\/w:t>)/g;

function paragraphText(xml: string): string {
  let out = "";
  for (const m of xml.matchAll(WT_RE)) out += decodeEntities(m[2] ?? "");
  return out;
}

/** Plain-text projection of a part: innermost paragraphs joined by spaces. */
function partText(xml: string): string {
  const parts: string[] = [];
  for (const m of xml.matchAll(PARAGRAPH_RE)) parts.push(paragraphText(m[0]));
  // paragraphs (and table cells) are separated by line breaks so a key
  // name can never be assembled across two of them
  return parts.join("\n");
}

function relsFor(zip: PizZip, part: string): Map<string, string> {
  const relsPath = part.replace(/word\/(.*)$/, "word/_rels/$1.rels");
  const rels = new Map<string, string>();
  const xml = zip.file(relsPath)?.asText();
  if (!xml) return rels;
  for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = m[0].match(/\bId="([^"]+)"/)?.[1];
    const target = m[0].match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) rels.set(id, target);
  }
  return rels;
}

const RASTER_MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
};

function thumbnailFor(
  zip: PizZip,
  rels: Map<string, string>,
  rId: string | undefined,
): { thumbnail?: string; ext?: string } {
  if (!rId) return {};
  const target = rels.get(rId);
  if (!target) return {};
  const path = target.startsWith("/")
    ? target.slice(1)
    : `word/${target.replace(/^\.\//, "")}`;
  const ext = (path.split(".").pop() ?? "").toLowerCase();
  const file = zip.file(path);
  if (!file) return { ext };
  const mime = RASTER_MIME[ext];
  if (!mime) return { ext };
  const data = file.asUint8Array();
  if (data.length > 400_000) return { ext };
  let bin = "";
  for (let i = 0; i < data.length; i++) bin += String.fromCharCode(data[i]);
  const b64 =
    typeof btoa === "function"
      ? btoa(bin)
      : Buffer.from(bin, "binary").toString("base64");
  return { thumbnail: `data:${mime};base64,${b64}`, ext };
}

/* ───────────────────── symbol / character tables ───────────────────── */

/** Wingdings & friends: char code (low byte) → checkbox state */
const SYMBOL_FONTS: Record<string, Record<number, boolean>> = {
  wingdings: {
    0x6f: false, // ☐
    0x70: false,
    0x71: false,
    0xa8: false, // ▫
    0xfd: true, // ☒
    0xfe: true, // ☑
  },
  "wingdings 2": {
    0x52: false,
    0x53: false,
    0xa2: false,
    0xa3: false,
    0x50: true,
    0x51: true,
    0x54: true,
    0x55: true,
    0x56: true,
  },
  "segoe ui symbol": { 0x2610: false, 0x2611: true, 0x2612: true },
  "ms gothic": { 0x2610: false, 0x2611: true, 0x2612: true },
};

/** Unicode box characters typed directly into the text */
const UNICODE_BOXES: Record<string, boolean> = {
  "☐": false, // ☐
  "☑": true, // ☑
  "☒": true, // ☒
  "□": false, // □
  "▢": false, // ▢
  "◻": false, // ◻
  "❏": false, // ❏
  "❐": false, // ❐
  "❑": false, // ❑
  "❒": false, // ❒
  "⬜": false, // ⬜
};

function symbolState(font: string, code: number): boolean | null {
  const table = SYMBOL_FONTS[font.toLowerCase().trim()];
  if (!table) return null;
  const low = code & 0xff;
  if (code in table) return table[code];
  if (low in table && code >= 0xf000) return table[low];
  return null;
}

function runFont(run: string): string {
  const m = run.match(
    /<w:rFonts\b[^>]*\bw:(?:ascii|hAnsi|cs|eastAsia)="([^"]+)"/,
  );
  return m ? m[1] : "";
}

/* ─────────────────────── candidate detection ─────────────────────── */

interface Span {
  start: number;
  end: number;
  kind: CandidateKind;
  suggested: CandidateType | null;
  label: string;
  text?: string;
  thumbnail?: string;
  widthIn?: number;
  heightIn?: number;
  /** true when the outer element must win over anything nested in it */
  strong: boolean;
  /** span sits inside a <w:t> text node (vs. being a whole element) */
  inText?: boolean;
  /** typed tag with a number ("c7"), kept unless the user renumbers */
  fixedId?: string;
  /** build the replacement XML for this span given the escaped marker */
  replace: (markerXml: string) => string;
}

function firstRPr(xml: string): string {
  return xml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
}

function markerRun(rPr: string, markerXml: string): string {
  return `<w:r>${rPr}<w:t xml:space="preserve">${markerXml}</w:t></w:r>`;
}

/** The <w:r> … </w:r> element that contains `index` (skips <w:rPr>). */
function enclosingRun(xml: string, index: number): { start: number; end: number } | null {
  let start = xml.lastIndexOf("<w:r", index);
  while (start >= 0 && !/^<w:r[\s>]/.test(xml.slice(start, start + 5))) {
    start = xml.lastIndexOf("<w:r", start - 1);
  }
  if (start < 0) return null;
  const end = xml.indexOf("</w:r>", index);
  return end < 0 ? null : { start, end: end + 6 };
}

function detectContentControls(xml: string): Span[] {
  const spans: Span[] = [];
  const stack: number[] = [];
  for (const m of xml.matchAll(/<w:sdt(?=[\s>])[^>]*>|<\/w:sdt>/g)) {
    const idx = m.index ?? 0;
    if (m[0].startsWith("</")) {
      const start = stack.pop();
      if (start === undefined) continue;
      const end = idx + m[0].length;
      const el = xml.slice(start, end);
      const pr = el.match(/<w:sdtPr>[\s\S]*?<\/w:sdtPr>/)?.[0] ?? "";
      const content =
        el.match(/<w:sdtContent>[\s\S]*<\/w:sdtContent>/)?.[0] ?? "";
      const block = /<w:p[\s>]/.test(content);
      const rPr = firstRPr(content);
      const replace = (markerXml: string) =>
        block
          ? `<w:p>${markerRun(rPr, markerXml)}</w:p>`
          : markerRun(rPr, markerXml);
      const text = paragraphText(content).trim();

      if (/<w:group\b/.test(pr)) continue; // group control: look inside
      if (/<w14:checkbox\b/.test(pr)) {
        const checked = /<w14:checked\b[^>]*w14:val="(?:1|true)"/.test(pr);
        spans.push({
          start,
          end,
          kind: "contentControl",
          suggested: checked ? "checkboxChecked" : "checkbox",
          label: checked ? "Word checkbox (checked)" : "Word checkbox",
          text,
          strong: true,
          replace,
        });
      } else if (/<w:text\b/.test(pr)) {
        spans.push({
          start, end, kind: "contentControl", suggested: "textbox",
          label: "Plain-text content control", text, strong: true, replace,
        });
      } else if (/<w:(?:date|comboBox|dropDownList)\b/.test(pr)) {
        spans.push({
          start, end, kind: "contentControl", suggested: "textbox",
          label: "Date / list content control", text, strong: true, replace,
        });
      } else if (/<w:picture\b/.test(pr)) {
        spans.push({
          start, end, kind: "contentControl", suggested: null,
          label: "Picture content control", text, strong: true, replace,
        });
      } else if (!block && text.length <= 80) {
        spans.push({
          start, end, kind: "contentControl", suggested: null,
          label: "Rich-text content control", text, strong: false, replace,
        });
      }
    } else {
      stack.push(idx);
    }
  }
  return spans;
}

function detectLegacyFields(xml: string): Span[] {
  const spans: Span[] = [];
  const re = /<w:fldChar\b[^>]*w:fldCharType="begin"[^>]*>(?:[\s\S]*?<\/w:fldChar>)?/g;
  for (const m of xml.matchAll(re)) {
    const idx = m.index ?? 0;
    if (!/<w:ffData>/.test(m[0])) continue;
    const ff = m[0];
    let suggested: CandidateType | null = null;
    let label = "";
    if (/<w:checkBox>/.test(ff)) {
      const checked = /<w:default\b[^>]*w:val="(?:1|true)"/.test(ff);
      suggested = checked ? "checkboxChecked" : "checkbox";
      label = checked ? "Legacy form checkbox (checked)" : "Legacy form checkbox";
    } else if (/<w:textInput\b/.test(ff)) {
      suggested = "textbox";
      label = "Legacy text form field";
    } else if (/<w:ddList\b/.test(ff)) {
      suggested = "textbox";
      label = "Legacy drop-down form field";
    } else continue;

    const beginRun = enclosingRun(xml, idx);
    if (!beginRun) continue;
    const endTag = xml.indexOf('w:fldCharType="end"', idx);
    if (endTag < 0) continue;
    const endRun = enclosingRun(xml, endTag);
    if (!endRun) continue;
    const rPr = firstRPr(xml.slice(beginRun.start, beginRun.end));
    const text = paragraphText(xml.slice(beginRun.end, endRun.start)).trim();
    spans.push({
      start: beginRun.start,
      end: endRun.end,
      kind: "legacyField",
      suggested,
      label,
      text,
      strong: true,
      replace: (markerXml) => markerRun(rPr, markerXml),
    });
  }
  return spans;
}

function detectRunCharacters(xml: string): Span[] {
  const spans: Span[] = [];
  for (const run of xml.matchAll(RUN_RE)) {
    const runStart = run.index ?? 0;
    const runXml = run[0];
    const font = runFont(runXml);

    // <w:sym w:font="Wingdings" w:char="F06F"/>
    for (const sym of runXml.matchAll(/<w:sym\b[^>]*\/>/g)) {
      const symFont = sym[0].match(/w:font="([^"]+)"/)?.[1] ?? "";
      const code = parseInt(sym[0].match(/w:char="([^"]+)"/)?.[1] ?? "", 16);
      if (Number.isNaN(code)) continue;
      const state = symbolState(symFont, code);
      const start = runStart + (sym.index ?? 0);
      spans.push({
        start,
        end: start + sym[0].length,
        kind: "symbol",
        suggested: state === null ? null : state ? "checkboxChecked" : "checkbox",
        label:
          state === null
            ? `${symFont} symbol ${code.toString(16).toUpperCase()}`
            : `${symFont} ${state ? "checked box" : "box"} symbol`,
        text: symFont,
        strong: true,
        replace: (markerXml) => `<w:t xml:space="preserve">${markerXml}</w:t>`,
      });
    }

    // characters inside <w:t>
    for (const wt of runXml.matchAll(WT_RE)) {
      const raw = wt[2];
      if (!raw) continue;
      const textStart = runStart + (wt.index ?? 0) + wt[0].indexOf(">") + 1;

      // tags the team typed in Word: <c> <cc> <t> (unnumbered) or <c7> <t2>
      // — in the XML they are entity-escaped: &lt;c7&gt;
      for (const m of raw.matchAll(/&lt;\s*([ct])\s*(\d*)\s*(c?)\s*&gt;/gi)) {
        const letter = m[1].toLowerCase() as "c" | "t";
        const digits = m[2];
        const checked = letter === "c" && m[3] !== "";
        const typed = `<${letter}${digits}${checked ? "c" : ""}>`;
        const start = textStart + (m.index ?? 0);
        spans.push({
          start,
          end: start + m[0].length,
          kind: "marker",
          suggested: letter === "t" ? "textbox" : checked ? "checkboxChecked" : "checkbox",
          label: digits ? `Typed ${typed} (numbered)` : `Typed ${typed} tag`,
          text: typed,
          fixedId: digits ? `${letter}${digits}` : undefined,
          strong: true,
          inText: true,
          replace: (markerXml) => markerXml,
        });
      }
      const charRe =
        /&#x([0-9a-f]+);|&#(\d+);|[-☐-☒□▢◻❏-❒⬜]|\[\s?[xX✓✔√]\s?\]|\[\s?\]|_{3,}/gi;
      for (const c of raw.matchAll(charRe)) {
        const start = textStart + (c.index ?? 0);
        const end = start + c[0].length;
        const replace = (markerXml: string) => markerXml;
        let ch = c[0];
        if (c[1]) ch = String.fromCodePoint(parseInt(c[1], 16));
        else if (c[2]) ch = String.fromCodePoint(parseInt(c[2], 10));

        if (/^_{3,}$/.test(ch)) {
          spans.push({
            start, end, kind: "blank", suggested: "textbox",
            label: `Underscore blank (${ch.length})`, text: ch, strong: true, inText: true, replace,
          });
        } else if (/^\[/.test(ch)) {
          const checked = /[xX✓✔√]/.test(ch);
          spans.push({
            start, end, kind: "bracketBox",
            suggested: checked ? "checkboxChecked" : "checkbox",
            label: checked ? "Typed [x] box" : "Typed [ ] box", text: ch, strong: true, inText: true, replace,
          });
        } else if (ch in UNICODE_BOXES) {
          const checked = UNICODE_BOXES[ch];
          spans.push({
            start, end, kind: "unicodeBox",
            suggested: checked ? "checkboxChecked" : "checkbox",
            label: `${ch} ${checked ? "checked box" : "box"} character`, text: ch, strong: true, inText: true, replace,
          });
        } else {
          const code = ch.codePointAt(0) ?? 0;
          if (code < 0xf000 || code > 0xf0ff) continue;
          const state = symbolState(font, code);
          spans.push({
            start, end, kind: "symbol",
            suggested: state === null ? null : state ? "checkboxChecked" : "checkbox",
            label:
              state === null
                ? `${font || "Symbol"} character ${code.toString(16).toUpperCase()}`
                : `${font} ${state ? "checked box" : "box"} symbol`,
            text: font, strong: true, inText: true, replace,
          });
        }
      }
    }
  }
  return spans;
}

function detectGraphics(
  xml: string,
  zip: PizZip,
  rels: Map<string, string>,
): Span[] {
  const spans: Span[] = [];
  const re =
    /<w:drawing>[\s\S]*?<\/w:drawing>|<w:pict(?:\s[^>]*)?>[\s\S]*?<\/w:pict>|<w:object(?:\s[^>]*)?>[\s\S]*?<\/w:object>/g;
  for (const m of xml.matchAll(re)) {
    const start = m.index ?? 0;
    const el = m[0];
    const end = start + el.length;

    let widthIn: number | undefined;
    let heightIn: number | undefined;
    const extent = el.match(/<wp:extent\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
    if (extent) {
      widthIn = +(parseInt(extent[1], 10) / EMU_PER_INCH).toFixed(2);
      heightIn = +(parseInt(extent[2], 10) / EMU_PER_INCH).toFixed(2);
    } else {
      const w = el.match(/width:\s*([\d.]+)(pt|in|px)/);
      const h = el.match(/height:\s*([\d.]+)(pt|in|px)/);
      const toIn = (v: string, u: string) =>
        +(parseFloat(v) / (u === "pt" ? 72 : u === "px" ? 96 : 1)).toFixed(2);
      if (w) widthIn = toIn(w[1], w[2]);
      if (h) heightIn = toIn(h[1], h[2]);
    }
    const size =
      widthIn !== undefined && heightIn !== undefined
        ? ` ${widthIn}×${heightIn} in`
        : "";
    const descr = el.match(/\bdescr="([^"]*)"/)?.[1]?.trim();
    const replace = (markerXml: string) =>
      `<w:t xml:space="preserve">${markerXml}</w:t>`;

    const hasText = /<wps:txbx>|<v:textbox\b/.test(el);
    const hasPic = /<pic:pic\b|<v:imagedata\b|<a:blip\b/.test(el);

    if (hasText) {
      const text = paragraphText(el).trim();
      spans.push({
        start, end, kind: "shape", suggested: "textbox",
        label: `Text box shape${size}`, text, widthIn, heightIn,
        strong: true, replace,
      });
      continue;
    }

    if (hasPic) {
      const rId =
        el.match(/<a:blip\b[^>]*\br:embed="([^"]+)"/)?.[1] ??
        el.match(/<v:imagedata\b[^>]*\br:id="([^"]+)"/)?.[1];
      const { thumbnail, ext } = thumbnailFor(zip, rels, rId);
      let suggested: CandidateType | null = null;
      let hint = "";
      if (widthIn !== undefined && heightIn !== undefined && heightIn > 0) {
        const ratio = widthIn / heightIn;
        const smallSquare =
          widthIn <= 0.45 && heightIn <= 0.45 && ratio > 0.7 && ratio < 1.4;
        if (smallSquare) {
          suggested = "checkbox";
        } else if (heightIn <= 0.35 && ratio >= 4) {
          suggested = "textbox"; // a thin line-like picture
        } else if (widthIn >= 1 && heightIn >= 0.5) {
          suggested = "ignore"; // logo / photo
          hint = " — looks like a logo or photo";
        }
      }
      // the alt text the client typed is the best hint we have
      if (descr) {
        if (/\b(checked|ticked|tick)\b/i.test(descr) && suggested !== "textbox") {
          suggested = "checkboxChecked";
        } else if (/check\s?box|\bbox\b/i.test(descr) && suggested !== "textbox") {
          suggested = "checkbox";
        } else if (/text\s?box|input|field|line|blank/i.test(descr)) {
          suggested = "textbox";
        }
      }
      spans.push({
        start, end, kind: "image", suggested,
        label: `Picture${ext ? ` (${ext})` : ""}${size}${descr ? ` — ${descr}` : hint}`,
        thumbnail, widthIn, heightIn, text: descr, strong: true, replace,
      });
      continue;
    }

    // a drawn line (Insert → Shapes → Line, or VML <v:line>) is a blank to
    // write on → textbox
    const isLine =
      /prst="(?:line|straightConnector1|bentConnector\d)"/.test(el) ||
      /<v:line\b/.test(el) ||
      (heightIn !== undefined && heightIn <= 0.05 && (widthIn ?? 0) >= 0.3);
    if (isLine) {
      spans.push({
        start, end, kind: "shape", suggested: "textbox",
        label: `Line shape${size}${descr ? ` — ${descr}` : ""}`,
        widthIn, heightIn, text: descr, strong: true, replace,
      });
      continue;
    }

    // drawn shape without text (rectangle etc.)
    let suggested: CandidateType | null = null;
    if (widthIn !== undefined && heightIn !== undefined && heightIn > 0) {
      const ratio = widthIn / heightIn;
      if (widthIn <= 0.45 && heightIn <= 0.45 && ratio > 0.7 && ratio < 1.4) {
        suggested = "checkbox";
      } else if (ratio >= 2.5) suggested = "textbox";
    }
    spans.push({
      start, end, kind: "shape", suggested,
      label: `Drawn shape${size}${descr ? ` — ${descr}` : ""}`,
      widthIn, heightIn, text: descr, strong: true, replace,
    });
  }
  return spans;
}

/** Resolve overlaps: strong outer spans swallow inner ones, weak ones give way. */
function resolveOverlaps(spans: Span[]): Span[] {
  const sorted = spans.slice().sort((a, b) => a.start - b.start || b.end - a.end);
  const drop = new Set<Span>();
  for (let i = 0; i < sorted.length; i++) {
    const outer = sorted[i];
    if (drop.has(outer)) continue;
    for (let j = i + 1; j < sorted.length && sorted[j].start < outer.end; j++) {
      const inner = sorted[j];
      if (inner.end > outer.end) continue; // partial overlap: keep both, first wins later
      if (outer.strong) drop.add(inner);
      else drop.add(outer);
    }
  }
  const kept = sorted.filter((s) => !drop.has(s));
  // remove any remaining partial overlaps (keep the earlier one)
  const out: Span[] = [];
  let lastEnd = -1;
  for (const s of kept) {
    if (s.start < lastEnd) continue;
    out.push(s);
    lastEnd = s.end;
  }
  return out;
}

const CONTEXT_MARK = "█"; // temporary sentinel, swapped for ▮ at the end

/** Text of `xml` with the span replaced by a sentinel that survives paragraphText(). */
function textWithMark(xml: string, span: Span, from: number, to: number): string {
  const marker = span.inText ? CONTEXT_MARK : `<w:t>${CONTEXT_MARK}</w:t>`;
  const marked = xml.slice(from, span.start) + marker + xml.slice(span.end, to);
  // paragraphText only reads complete <w:t> elements; PARAGRAPH_RE is not
  // needed here because we already sliced to one paragraph / row
  let out = "";
  for (const m of marked.matchAll(WT_RE)) out += decodeEntities(m[2] ?? "");
  return out.replace(/\s+/g, " ");
}

function trimAround(text: string): string {
  const i = text.indexOf(CONTEXT_MARK);
  if (i < 0) return "▮";
  let before = text.slice(0, i);
  let after = text.slice(i + 1);
  if (before.length > 60) before = "…" + before.slice(-60);
  if (after.length > 60) after = after.slice(0, 60) + "…";
  return `${before}▮${after}`.trim();
}

/**
 * Where the candidate sits: its paragraph text with ▮ in place of the
 * candidate. Inside a table cell that only holds the control, the whole
 * row is used instead so the label in the neighbouring cell shows up.
 */
function contextFor(xml: string, span: Span): string {
  const pStart = Math.max(
    xml.lastIndexOf("<w:p ", span.start),
    xml.lastIndexOf("<w:p>", span.start),
  );
  const pEnd = xml.indexOf("</w:p>", span.end);
  if (pStart < 0 || pEnd < 0) return "▮";

  const own = trimAround(textWithMark(xml, span, pStart, pEnd));
  if (own !== "▮") return own;

  // empty paragraph → try the table row
  const trStart = Math.max(
    xml.lastIndexOf("<w:tr ", span.start),
    xml.lastIndexOf("<w:tr>", span.start),
  );
  const trEnd = xml.indexOf("</w:tr>", span.end);
  const rowEndsFirst = xml.indexOf("</w:tbl>", span.end);
  const insideRow =
    trStart >= 0 &&
    trEnd >= 0 &&
    xml.lastIndexOf("</w:tr>", span.start) < trStart && // that row is still open
    xml.lastIndexOf("</w:tbl>", span.start) < trStart && // its table too
    (rowEndsFirst < 0 || trEnd < rowEndsFirst);
  if (insideRow) {
    const row = trimAround(textWithMark(xml, span, trStart, trEnd));
    if (row !== "▮") return row;
  }

  // still nothing → previous paragraph's text as a hint
  const prevEnd = xml.lastIndexOf("</w:p>", pStart);
  const prevStart = Math.max(
    xml.lastIndexOf("<w:p ", prevEnd),
    xml.lastIndexOf("<w:p>", prevEnd),
  );
  if (prevStart >= 0 && prevEnd > prevStart) {
    const prev = paragraphText(xml.slice(prevStart, prevEnd)).replace(/\s+/g, " ").trim();
    if (prev) return `${prev.length > 60 ? "…" + prev.slice(-60) : prev} ⏎ ▮`;
  }
  return "▮";
}

/**
 * "Lines" people draw to write on — all of them are textboxes:
 *   • underlined spaces / tabs  ("Name: ______" made with Ctrl+U)
 *   • tab stops with an underscore / dot leader
 *   • an empty paragraph with a bottom border (signature line)
 * (drawn line shapes are handled in detectGraphics, "____" in
 * detectRunCharacters)
 */
function detectLines(xml: string): Span[] {
  const spans: Span[] = [];

  // 1. underlined blank runs — merge neighbours, Word splits them freely
  const underlined: Span[] = [];
  for (const run of xml.matchAll(RUN_RE)) {
    const runXml = run[0];
    const runStart = run.index ?? 0;
    const rPr = runXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
    if (!/<w:u\b(?![^>]*w:val="none")/.test(rPr)) continue;
    const body = runXml.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, "");
    const onlyBlankParts = body
      .replace(/<w:t\b[^>]*>[^<]*<\/w:t>|<w:t\b[^>]*\/>|<w:tab\/>/g, "")
      .replace(/^<w:r\b[^>]*>|<\/w:r>$/g, "")
      .trim() === "";
    if (!onlyBlankParts) continue;
    const text = paragraphText(body);
    const hasTab = /<w:tab\/>/.test(body);
    if (!/^[\s _]*$/.test(text)) continue;
    if (!hasTab && text.replace(/_/g, "").length < 2) continue; // a single underlined space is nothing
    underlined.push({
      start: runStart,
      end: runStart + runXml.length,
      kind: "blank",
      suggested: "textbox",
      label: hasTab ? "Underlined tab (line)" : `Underlined blank (${text.length} spaces)`,
      text: hasTab ? "underlined tab" : "underlined spaces",
      strong: true,
      replace: (markerXml) => markerRun(rPr, markerXml),
    });
  }
  underlined.sort((a, b) => a.start - b.start);
  for (const s of underlined) {
    const last = spans[spans.length - 1];
    if (last && last.kind === "blank" && last.end === s.start && /Underlined/.test(last.label)) {
      last.end = s.end; // one visual line made of several runs
      last.label = "Underlined blank (line)";
    } else {
      spans.push(s);
    }
  }

  // 2. + 3. per paragraph: tab leaders and bordered empty paragraphs
  for (const pm of xml.matchAll(PARAGRAPH_RE)) {
    const pXml = pm[0];
    const pStart = pm.index ?? 0;
    const pPr = pXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? "";
    const body = pXml.slice(pPr ? pXml.indexOf(pPr) + pPr.length : 0);
    const bodyStart = pStart + (pPr ? pXml.indexOf(pPr) + pPr.length : 0);

    if (/<w:tab\b[^>]*w:leader="(?:underscore|dot|hyphen|heavy|middleDot)"/.test(pPr)) {
      for (const t of body.matchAll(/<w:tab\/>/g)) {
        const start = bodyStart + (t.index ?? 0);
        spans.push({
          start,
          end: start + t[0].length,
          kind: "blank",
          suggested: "textbox",
          label: "Tab leader line",
          text: "tab leader",
          strong: false, // an underlined run around it wins
          replace: (markerXml) => `<w:t xml:space="preserve">${markerXml}</w:t><w:tab/>`,
        });
      }
    }

    const bordered = /<w:pBdr>[\s\S]*?<w:bottom\b(?![^>]*w:val="(?:none|nil)")[\s\S]*?<\/w:pBdr>/.test(pPr);
    if (bordered && paragraphText(body).trim() === "" && !/<w:(?:drawing|pict|sdt|fldChar)\b/.test(body)) {
      const at = pStart + pXml.lastIndexOf("</w:p>");
      spans.push({
        start: at,
        end: at,
        kind: "blank",
        suggested: "textbox",
        label: "Line under empty paragraph (border)",
        text: "bottom border",
        strong: true,
        replace: (markerXml) => markerRun("", markerXml),
      });
    }
  }

  return spans;
}

function detectSpans(xml: string, zip: PizZip, part: string): Span[] {
  const rels = relsFor(zip, part);
  return resolveOverlaps([
    ...detectContentControls(xml),
    ...detectLegacyFields(xml),
    ...detectRunCharacters(xml),
    ...detectGraphics(xml, zip, rels),
    ...detectLines(xml),
  ]);
}

/* ───────────────────────────── analyse ───────────────────────────── */

export function analyzeEsignDocx(
  input: ArrayBuffer | Uint8Array,
  options: AnalyzeOptions = {},
): DocxAnalysis {
  const zip = new PizZip(input instanceof Uint8Array ? input : new Uint8Array(input));
  const parts = listDocxParts(zip);

  let text = "";
  const candidates: DocxCandidate[] = [];

  for (const part of parts) {
    const xml = zip.file(part)?.asText();
    if (!xml) continue;
    text += "\n" + partText(xml);

    const key = partKey(part);
    detectSpans(xml, zip, part).forEach((span, i) => {
      candidates.push({
        id: `${key}-${i}`,
        part,
        kind: span.kind,
        suggested: span.suggested,
        label: span.label,
        context: contextFor(xml, span),
        thumbnail: span.thumbnail,
        widthIn: span.widthIn,
        heightIn: span.heightIn,
        text: span.text,
        fixedId: span.fixedId,
      });
    });
  }

  const analysis = analyzeEsignText(
    text.replace(/[ \t\r]+/g, " ").replace(/\s*\n\s*/g, "\n").trim(),
    options,
  );
  return { ...analysis, candidates };
}

/* ───────────────────── key replacement (run-safe) ───────────────────── */

interface Segment {
  index: number;
  length: number;
  attrs: string;
  text: string;
  changed: boolean;
}

/** Replace [start,end) of the merged text with `replacement`, keeping runs. */
function splice(segments: Segment[], start: number, end: number, replacement: string) {
  let offset = 0;
  let inserted = false;
  for (const seg of segments) {
    const segStart = offset;
    const segEnd = offset + seg.text.length;
    offset = segEnd;
    if (segEnd <= start || segStart >= end) continue;
    const localStart = Math.max(start, segStart) - segStart;
    const localEnd = Math.min(end, segEnd) - segStart;
    const before = seg.text.slice(0, localStart);
    const after = seg.text.slice(localEnd);
    seg.text = inserted ? before + after : before + replacement + after;
    seg.changed = true;
    inserted = true;
  }
}

interface TextRule {
  regex: RegExp;
  replacement: string;
  onHit: () => void;
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function plainRegex(form: EsignTokenForm | "text", raw: string, norm: string): RegExp {
  const loose = norm.split("").map(escapeRegExp).join("[\\s_-]*");
  if (form === "bracket") return new RegExp(`(?:<<|«)\\s*${loose}\\s*(?:>>|»)`, "gi");
  if (form === "at") return new RegExp(`@${escapeRegExp(norm)}(?![A-Za-z0-9_])`, "gi");
  if (form === "bare") {
    return new RegExp(`(?<![A-Za-z0-9_@])${escapeRegExp(norm)}(?![A-Za-z0-9_])`, "g");
  }
  if (form === "plain") {
    // exact spelling of this variant; only separators are flexible
    const letters = raw.replace(/[\s_-]+/g, "");
    const relaxed = letters.split("").map(escapeRegExp).join("[\\s_-]*");
    return new RegExp(`(?<![A-Za-z0-9])${relaxed}(?![A-Za-z0-9])`, "g");
  }
  const words = raw.trim().split(/\s+/).map(escapeRegExp).join("\\s+");
  return new RegExp(`(?<![A-Za-z0-9])${words}(?![A-Za-z0-9])`, "gi");
}

function replaceKeysInParagraph(paragraph: string, rules: TextRule[]): string {
  const segments: Segment[] = [];
  for (const m of paragraph.matchAll(WT_RE)) {
    segments.push({
      index: m.index ?? 0,
      length: m[0].length,
      attrs: m[1] ?? "",
      text: decodeEntities(m[2] ?? ""),
      changed: false,
    });
  }
  if (!segments.length) return paragraph;

  for (const rule of rules) {
    let from = 0;
    for (;;) {
      const merged = segments.map((s) => s.text).join("");
      rule.regex.lastIndex = from;
      const m = rule.regex.exec(merged);
      if (!m) break;
      splice(segments, m.index, m.index + m[0].length, rule.replacement);
      rule.onHit();
      from = m.index + rule.replacement.length;
    }
  }

  if (!segments.some((s) => s.changed)) return paragraph;

  let out = "";
  let last = 0;
  for (const seg of segments) {
    out += paragraph.slice(last, seg.index);
    if (seg.changed) {
      let attrs = seg.attrs;
      if (!/xml:space=/.test(attrs)) attrs += ' xml:space="preserve"';
      out += `<w:t${attrs}>${escapeXml(seg.text)}</w:t>`;
    } else {
      out += paragraph.slice(seg.index, seg.index + seg.length);
    }
    last = seg.index + seg.length;
  }
  return out + paragraph.slice(last);
}

/* ────────────────────────────── build ────────────────────────────── */

const SYMBOL_FONT_RE =
  /<w:rFonts\b[^>]*(?:Wingdings|Webdings|Symbol|MS Gothic|Segoe UI Symbol)[^>]*\/>/gi;

/** A marker written into a Wingdings run would render as glyph soup — drop the font. */
function stripSymbolFontsOnMarkers(xml: string): string {
  return xml.replace(RUN_RE, (run) =>
    /<w:t[^>]*>&lt;[ct]\d+c?&gt;<\/w:t>/i.test(run) ? run.replace(SYMBOL_FONT_RE, "") : run,
  );
}

export function buildEsignDocx(
  input: ArrayBuffer | Uint8Array,
  options: DocxBuildOptions = {},
): DocxBuildResult {
  const zip = new PizZip(input instanceof Uint8Array ? input : new Uint8Array(input));
  const parts = listDocxParts(zip);
  const assignments = options.assignments ?? {};
  const mappings = options.mappings ?? {};
  const excluded = new Set(options.exclude ?? []);
  const replaceKeys = options.replaceKeys ?? true;
  const stats: DocxBuildStats = { keys: {}, markers: {}, skipped: [] };

  // key rules come from the same analysis the client saw
  const analysis = analyzeEsignDocx(input, { plainText: options.plainText });
  const rules: TextRule[] = [];
  if (replaceKeys) {
    for (const spec of keyTargetSpecs(analysis, mappings)) {
      // an explicit mapping always wins over "not ticked"
      if (excluded.has(spec.raw) && !hasMappingFor(spec.raw, mappings)) continue;
      const resolved = resolveToken(spec.raw, mappings);
      if (!resolved) {
        if (!spec.fromMapping) stats.skipped.push(spec.raw);
        continue;
      }
      const replacement =
        resolved.kind === "placeholder" ? `<<${resolved.key}>>` : resolved.text;
      const statKey =
        resolved.kind === "placeholder" ? resolved.key : `"${resolved.text}"`;
      rules.push({
        regex: plainRegex(spec.form, spec.raw, spec.norm),
        replacement,
        onHit: () => {
          stats.keys[statKey] = (stats.keys[statKey] ?? 0) + 1;
        },
      });
    }
  }

  for (const part of parts) {
    const xml = zip.file(part)?.asText();
    if (!xml) continue;
    const key = partKey(part);
    const spans = detectSpans(xml, zip, part);

    // 1. candidates → marker text, applied back-to-front so offsets stay valid
    let out = xml;
    for (let i = spans.length - 1; i >= 0; i--) {
      const span = spans[i];
      const a = assignments[`${key}-${i}`];
      if (!a || a.type === "ignore" || !/^[ct]\d+$/i.test(a.id)) continue;
      const marker = `${a.id.toLowerCase()}${a.type === "checkboxChecked" ? "c" : ""}`;
      const markerXml = `&lt;${marker}&gt;`;
      out = out.slice(0, span.start) + span.replace(markerXml) + out.slice(span.end);
      stats.markers[`${key}-${i}`] = marker;
    }
    out = stripSymbolFontsOnMarkers(out);

    // 2. keys → canonical placeholders, run formatting preserved
    if (rules.length) {
      out = out.replace(PARAGRAPH_RE, (p) => replaceKeysInParagraph(p, rules));
    }

    if (out !== xml) zip.file(part, out);
  }

  const bytes = zip.generate({ type: "uint8array", compression: "DEFLATE" }) as Uint8Array;
  return { bytes, stats };
}

/* ─────────────────── manual tags: insert after some text ─────────────────── */

export interface TagInsert {
  /** text to look for (case-insensitive, whitespace-tolerant) */
  afterText: string;
  type: "checkbox" | "checkboxChecked" | "textbox";
  /** insert after every occurrence instead of only the first */
  all?: boolean;
}

export interface TagInsertResult {
  bytes: Uint8Array;
  /** how many tags were written, per insert (same order as the input) */
  inserted: number[];
}

const TAG_TEXT: Record<TagInsert["type"], string> = {
  checkbox: "<c>",
  checkboxChecked: "<cc>",
  textbox: "<t>",
};

/** Insert `text` into the run that contains character position `pos`. */
function insertAt(segments: Segment[], pos: number, text: string): boolean {
  let offset = 0;
  for (const seg of segments) {
    const segEnd = offset + seg.text.length;
    if (pos >= offset && pos <= segEnd) {
      const local = pos - offset;
      seg.text = seg.text.slice(0, local) + text + seg.text.slice(local);
      seg.changed = true;
      return true;
    }
    offset = segEnd;
  }
  return false;
}

function rebuildParagraph(paragraph: string, segments: Segment[]): string {
  let out = "";
  let last = 0;
  for (const seg of segments) {
    out += paragraph.slice(last, seg.index);
    if (seg.changed) {
      let attrs = seg.attrs;
      if (!/xml:space=/.test(attrs)) attrs += ' xml:space="preserve"';
      out += `<w:t${attrs}>${escapeXml(seg.text)}</w:t>`;
    } else {
      out += paragraph.slice(seg.index, seg.index + seg.length);
    }
    last = seg.index + seg.length;
  }
  return out + paragraph.slice(last);
}

/**
 * Write unnumbered tags (<c>, <cc>, <t>) right after a piece of text the
 * user names, inside the run that holds that text — so the tag picks up
 * the surrounding formatting and nothing else in the document moves.
 * The result is a normal .docx that analyzeEsignDocx() then treats like
 * any file with typed tags (numbered on build, in document order).
 */
export function insertTypedTags(
  input: ArrayBuffer | Uint8Array,
  inserts: TagInsert[],
): TagInsertResult {
  const zip = new PizZip(input instanceof Uint8Array ? input : new Uint8Array(input));
  const inserted = inserts.map(() => 0);

  const rules = inserts.map((ins, i) => {
    const words = ins.afterText.trim().split(/\s+/).filter(Boolean).map(escapeRegExp);
    return {
      i,
      regex: words.length ? new RegExp(words.join("\\s+"), "gi") : null,
      text: " " + TAG_TEXT[ins.type],
      all: ins.all === true,
    };
  });

  for (const part of listDocxParts(zip)) {
    const xml = zip.file(part)?.asText();
    if (!xml) continue;
    let changed = false;

    const out = xml.replace(PARAGRAPH_RE, (paragraph) => {
      const segments: Segment[] = [];
      for (const m of paragraph.matchAll(WT_RE)) {
        segments.push({
          index: m.index ?? 0,
          length: m[0].length,
          attrs: m[1] ?? "",
          text: decodeEntities(m[2] ?? ""),
          changed: false,
        });
      }
      if (!segments.length) return paragraph;

      let touched = false;
      for (const rule of rules) {
        if (!rule.regex) continue;
        if (!rule.all && inserted[rule.i] > 0) continue;
        let from = 0;
        for (;;) {
          const merged = segments.map((s) => s.text).join("");
          rule.regex.lastIndex = from;
          const m = rule.regex.exec(merged);
          if (!m) break;
          const at = m.index + m[0].length;
          if (insertAt(segments, at, rule.text)) {
            inserted[rule.i]++;
            touched = true;
          }
          if (!rule.all) break;
          from = at + rule.text.length;
        }
      }
      if (!touched) return paragraph;
      changed = true;
      return rebuildParagraph(paragraph, segments);
    });

    if (changed) zip.file(part, out);
  }

  const bytes = zip.generate({ type: "uint8array", compression: "DEFLATE" }) as Uint8Array;
  return { bytes, inserted };
}

export { normalizeKey };
