/* ──────────────────────────────────────────────────────────────────────
 * Legacy adapter — DocX Key Replacer (/api/analyze, /api/replace)
 *
 * The original DocX Key Replacer UI (app/page.tsx) is unchanged; its two
 * endpoints now run on the esign engine through this adapter so that:
 *   • run formatting is preserved (text is spliced inside the existing
 *     <w:t> runs instead of merging a paragraph into its first run),
 *   • <<KEY>>, &lt;&lt;KEY&gt;&gt;, «KEY», @KEY and bare KEY are all found
 *     and rewritten, also when Word split them across runs,
 *   • HTML files keep their charset and get browser-safe encoded keys,
 *   • footnotes / comments are not touched (same parts as before:
 *     document, headers, footers).
 *
 * Behaviour kept from the old module on purpose:
 *   • only bracket / @ / CAPITALS forms are rewritten — key names written
 *     as ordinary text ("Student Name:") are left alone,
 *   • a mapping value that is not a placeholder is written as literal text
 *     (the KeysList date picker relies on this),
 *   • checkbox / textbox markers are never converted here.
 * ────────────────────────────────────────────────────────────────────── */

import {
  analyzeEsignHtml,
  decodeHtmlBytes,
  encodeHtmlForCharset,
  transformEsignHtml,
  type EsignAnalysis,
} from "./html-template";
import { analyzeEsignDocx, buildEsignDocx } from "./docx-template";

export const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export interface LegacyAnalysis {
  matchedKeys: string[];
  unmatchedKeys: string[];
}

export interface LegacyReplaceResult {
  bytes: Uint8Array;
  contentType: string;
  /** canonical key → occurrences rewritten */
  replaced: Record<string, number>;
  /** tokens with no known key and no mapping, left as they were */
  skipped: string[];
}

export function isHtmlName(name: string): boolean {
  const n = name.toLowerCase();
  return n.endsWith(".html") || n.endsWith(".htm");
}

function toLegacy(analysis: EsignAnalysis): LegacyAnalysis {
  // matched: keys.json names; unmatched: raw <<…>> / @… tokens that are
  // not known keys (same shape the UI always received)
  return {
    matchedKeys: analysis.matchedKeys,
    unmatchedKeys: analysis.tokens
      .filter((t) => !t.knownKey && (t.form === "bracket" || t.form === "at"))
      .map((t) => t.raw),
  };
}

export function legacyAnalyze(bytes: Uint8Array, fileName: string): LegacyAnalysis {
  if (isHtmlName(fileName)) {
    const { html } = decodeHtmlBytes(bytes);
    return toLegacy(analyzeEsignHtml(html, { plainText: false }));
  }
  return toLegacy(analyzeEsignDocx(bytes, { plainText: false }));
}

export function legacyReplace(
  bytes: Uint8Array,
  fileName: string,
  keyMappings: Record<string, string>,
): LegacyReplaceResult {
  // ignore empty mapping values — the UI sends "" for unmapped keys
  const mappings = Object.fromEntries(
    Object.entries(keyMappings).filter(([k, v]) => k.trim() && typeof v === "string" && v.trim()),
  );

  if (isHtmlName(fileName)) {
    const { html, charset } = decodeHtmlBytes(bytes);
    const analysis = analyzeEsignHtml(html, { plainText: false });
    const { html: out, stats } = transformEsignHtml(html, analysis, {
      mappings,
      keyOutput: "encoded",
      replaceKeys: true,
      replaceControls: false,
    });
    const encoded = encodeHtmlForCharset(out, charset);
    return {
      bytes: encoded.bytes,
      contentType: `text/html; charset=${encoded.charset}`,
      replaced: stats.keys,
      skipped: stats.skipped,
    };
  }

  const { bytes: out, stats } = buildEsignDocx(bytes, {
    mappings,
    replaceKeys: true,
    plainText: false,
    assignments: {}, // never touch checkbox / textbox candidates here
  });
  return {
    bytes: out,
    contentType: DOCX_MIME,
    replaced: stats.keys,
    skipped: stats.skipped,
  };
}
