/* ──────────────────────────────────────────────────────────────────────
 * Esign-Templates / Customized — HTML template engine
 *
 * Completely separate from the existing DOCX / HTML key-replacement
 * pipeline (/api/analyze, /api/replace, /api/replace-html). Nothing in
 * here is imported by that code and nothing from that code is imported
 * here, except the shared keys.json dictionary.
 *
 * Input:  an .html/.htm file that was produced by converting a Word
 *         document. Inside it we expect:
 *           • placeholder keys   <<KEY>>  — after Word → HTML conversion
 *                                they appear as  &lt;&lt;KEY&gt;&gt;
 *                                (Word AutoCorrect may also turn << >>
 *                                into « »; that form is handled too)
 *           • checkbox markers   <c1> <c2> …   → unchecked checkbox
 *                                <c1c> <c2c> … → already-checked checkbox
 *           • textbox markers    <t1> <t2> …   → text input
 *         Markers are typed as plain text in Word, so in the HTML they
 *         appear as &lt;c1&gt; etc. Word also likes to split runs across
 *         <span> tags, so every matcher tolerates tags / &nbsp; /
 *         whitespace between the characters of a token.
 *
 * Output: the same HTML with
 *           • every key rewritten to its canonical placeholder
 *             (<<KEY>> literal by default, or &lt;&lt;KEY&gt;&gt;)
 *           • every marker rewritten to a real form control built from
 *             a user-configurable template
 * ────────────────────────────────────────────────────────────────────── */

import keysJson from "../../keys.json";

/* ───────────────────────────── types ───────────────────────────── */

/**
 * bracket  <<KEY>> / &lt;&lt;KEY&gt;&gt; / «KEY»
 * at       @KEY
 * bare     KEY written in capitals exactly as in keys.json
 * plain    key name written any other way: studentname, Student Name …
 *          (found only when the same key is not already used in a
 *          stricter form, and only for names of 6+ characters)
 */
export type EsignTokenForm = "bracket" | "at" | "bare" | "plain";

export interface EsignToken {
  /** Token as it reads in the document, e.g. "<<Student Name>>", "@FIRSTNAME", "CAMPUSNAME" */
  raw: string;
  /** Normalised key name, e.g. "STUDENTNAME" */
  norm: string;
  form: EsignTokenForm;
  /** keys.json name when the token is a known key, otherwise null */
  knownKey: string | null;
  count: number;
}

export type EsignControlType = "checkbox" | "textbox";

export interface EsignControl {
  /** Marker text without brackets, lower-cased: "c1", "c1c", "t3" */
  marker: string;
  /** Control id: "c1", "t3" (the checked flag is stripped) */
  id: string;
  type: EsignControlType;
  n: number;
  checked: boolean;
  count: number;
}

export interface EsignAnalysis {
  tokens: EsignToken[];
  controls: EsignControl[];
  /** keys.json names found in the document */
  matchedKeys: string[];
  /** raw tokens that are not known keys and need a mapping */
  unmatchedKeys: string[];
  checkboxes: EsignControl[];
  textboxes: EsignControl[];
  /** ids used by more than one marker (e.g. <c1> twice, or <c1> and <c1c>) */
  duplicateIds: string[];
  missingNumbers: { checkbox: number[]; textbox: number[] };
}

export interface EsignTemplates {
  checkbox: string;
  checkboxChecked: string;
  textbox: string;
}

export type EsignKeyOutput = "literal" | "encoded";

export interface EsignReplaceOptions {
  /** raw token / free text → placeholder ("<<KEY>>") or plain text */
  mappings?: Record<string, string>;
  /** raw tokens the user chose NOT to replace (e.g. a "Student Name:" label) */
  exclude?: string[];
  keyOutput?: EsignKeyOutput;
  replaceKeys?: boolean;
  replaceControls?: boolean;
  templates?: Partial<EsignTemplates>;
}

export interface EsignReplaceStats {
  /** canonical key → number of occurrences rewritten */
  keys: Record<string, number>;
  /** marker → number of occurrences rewritten */
  controls: Record<string, number>;
  /** raw tokens that had no known key and no mapping, left untouched */
  skipped: string[];
}

export interface EsignReplaceResult {
  html: string;
  stats: EsignReplaceStats;
}

export const DEFAULT_TEMPLATES: EsignTemplates = {
  checkbox: '<input type="checkbox" id="{id}" name="{id}" />',
  checkboxChecked:
    '<input type="checkbox" id="{id}" name="{id}" checked="checked" />',
  textbox: '<input type="text" id="{id}" name="{id}" />',
};

/* ─────────────────────────── key dictionary ─────────────────────────── */

const KEYS = keysJson as Record<string, string>;
const KEY_BY_NORM = new Map<string, string>();
for (const name of Object.keys(KEYS)) KEY_BY_NORM.set(normalizeKey(name), name);

export const ESIGN_KNOWN_KEYS: string[] = Object.keys(KEYS);

/** "<<First Name>>" / "&lt;&lt;first name&gt;&gt;" / "@first_name" → "FIRSTNAME" */
export function normalizeKey(raw: string): string {
  return raw
    .trim()
    .replace(/^(?:<<|&lt;&lt;|«|&laquo;)\s*/i, "")
    .replace(/\s*(?:>>|&gt;&gt;|»|&raquo;)$/i, "")
    .replace(/^@/, "")
    .replace(/[\s_-]+/g, "") // "student_name" / "student-name" / "student name"
    .toUpperCase();
}

/** "<<GUARDIANGENDER-1>>" → "GUARDIANGENDER-1" (keys.json spelling kept verbatim) */
export function placeholderInner(v: string): string {
  return v
    .trim()
    .replace(/^(?:<<|&lt;&lt;|«|&laquo;)\s*/i, "")
    .replace(/\s*(?:>>|&gt;&gt;|»|&raquo;)$/i, "")
    .trim();
}

export function isPlaceholderValue(v: string): boolean {
  return /^\s*(?:<<|&lt;&lt;|«|&laquo;)[^<>]*(?:>>|&gt;&gt;|»|&raquo;)\s*$/i.test(
    v,
  );
}

/* ─────────────────────────── small helpers ─────────────────────────── */

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function decodeEntities(s: string): string {
  return s
    .replace(/&#x([0-9a-f]+);/gi, (m, h) => {
      try {
        return String.fromCodePoint(parseInt(h, 16));
      } catch {
        return m;
      }
    })
    .replace(/&#(\d+);/g, (m, d) => {
      try {
        return String.fromCodePoint(parseInt(d, 10));
      } catch {
        return m;
      }
    })
    .replace(/&nbsp;/gi, " ")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&laquo;/gi, "«")
    .replace(/&raquo;/gi, "»")
    .replace(/&amp;/gi, "&");
}

/* ───────────────────────── charset detection ───────────────────────── */

/**
 * Word saves HTML as windows-1252 (or utf-8 with BOM). Decode using the
 * charset declared in the file so special characters survive the trip.
 */
export function decodeHtmlBytes(input: ArrayBuffer | Uint8Array): {
  html: string;
  charset: string;
} {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);

  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return { html: new TextDecoder("utf-8").decode(bytes), charset: "utf-8" };
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return {
      html: new TextDecoder("utf-16le").decode(bytes),
      charset: "utf-16le",
    };
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return {
      html: new TextDecoder("utf-16be").decode(bytes),
      charset: "utf-16be",
    };
  }

  const head = decodeWindows1252(bytes.subarray(0, 4096));
  const declared = head.match(/charset\s*=\s*["']?\s*([\w-]+)/i)?.[1];
  const charset = (declared || "utf-8").toLowerCase();

  if (CP1252_LABELS.has(charset)) {
    return { html: decodeWindows1252(bytes), charset: "windows-1252" };
  }
  if (charset !== "utf-8" && charset !== "utf8") {
    try {
      return { html: new TextDecoder(charset).decode(bytes), charset };
    } catch {
      /* unknown label → fall through */
    }
  }
  try {
    return {
      html: new TextDecoder("utf-8", { fatal: true }).decode(bytes),
      charset: "utf-8",
    };
  } catch {
    return { html: decodeWindows1252(bytes), charset: "windows-1252" };
  }
}

const CP1252_LABELS = new Set([
  "windows-1252",
  "cp1252",
  "cp-1252",
  "x-cp1252",
  "latin1",
  "latin-1",
  "iso-8859-1",
  "iso8859-1",
  "us-ascii",
  "ascii",
  "ansi",
]);

/* 0x80–0x9F of Windows-1252; 0 = unmapped (kept as the raw code point) */
const CP1252_HIGH = [
  0x20ac, 0, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030,
  0x0160, 0x2039, 0x0152, 0, 0x017d, 0, 0, 0x2018, 0x2019, 0x201c, 0x201d,
  0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0, 0x017e,
  0x0178,
];

/**
 * Byte-exact Windows-1252 decoder. Node's TextDecoder maps this label to
 * plain ISO-8859-1, which turns Word's curly quotes / dashes (0x91–0x97)
 * into C1 control characters, so we do it by hand.
 */
export function decodeWindows1252(bytes: Uint8Array): string {
  let out = "";
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i];
    if (b >= 0x80 && b <= 0x9f) {
      const mapped = CP1252_HIGH[b - 0x80];
      out += String.fromCharCode(mapped || b);
    } else {
      out += String.fromCharCode(b);
    }
  }
  return out;
}

const CP1252_REVERSE = new Map<number, number>();
CP1252_HIGH.forEach((cp, i) => {
  if (cp) CP1252_REVERSE.set(cp, 0x80 + i);
});

/**
 * Encode back to Windows-1252 so a Word-exported file keeps its original
 * charset declaration byte-for-byte. Characters outside the code page are
 * written as numeric entities, which every browser renders identically.
 */
export function encodeWindows1252(text: string): Uint8Array {
  const out: number[] = [];
  for (const ch of text) {
    const cp = ch.codePointAt(0) as number;
    if (cp < 0x80 || (cp >= 0xa0 && cp <= 0xff)) {
      out.push(cp);
    } else if (CP1252_REVERSE.has(cp)) {
      out.push(CP1252_REVERSE.get(cp) as number);
    } else if (cp >= 0x80 && cp <= 0x9f) {
      out.push(cp);
    } else {
      const entity = `&#${cp};`;
      for (let i = 0; i < entity.length; i++) out.push(entity.charCodeAt(i));
    }
  }
  return Uint8Array.from(out);
}

/**
 * Turn the finished HTML back into bytes in the charset the file arrived
 * with, so nothing but the replaced tokens differs from the upload.
 * Charsets we cannot write (utf-16 etc.) fall back to UTF-8 and the
 * <meta charset> is updated to match.
 */
export function encodeHtmlForCharset(
  html: string,
  charset: string,
): { bytes: Uint8Array; charset: string } {
  const label = charset.toLowerCase();
  if (label === "windows-1252") {
    return { bytes: encodeWindows1252(html), charset: "windows-1252" };
  }
  if (label === "utf-8" || label === "utf8") {
    return { bytes: new TextEncoder().encode(html), charset: "utf-8" };
  }
  return {
    bytes: new TextEncoder().encode(fixMetaCharset(html)),
    charset: "utf-8",
  };
}

/* ────────────────────────── text view of html ────────────────────────── */

const BLOCK_TAG_RE =
  /<\/?(?:p|div|br|hr|td|th|tr|table|li|ul|ol|h[1-6]|section|article|header|footer|blockquote|pre|body|html)\b[^>]*>/gi;
const LITERAL_KEY_RE = /<<([^<>]*)>>/g;
const LITERAL_MARKER_RE = /<\s*([ct]\s*\d+\s*c?)\s*>/gi;
const PROTECTED_RE =
  /<!--[\s\S]*?-->|<(style|script|textarea|title|xml)\b[^>]*>[\s\S]*?<\/\1\s*>/gi;

/**
 * Plain-text projection of the HTML: comments/style/script removed,
 * inline tags dropped (so run-split tokens join back up), block tags
 * become a space, entities decoded. Used only for token discovery.
 */
export function htmlToText(html: string): string {
  let s = html.replace(PROTECTED_RE, " ");
  // keep literal <<KEY>> and <c1> alive through tag stripping
  s = s.replace(LITERAL_KEY_RE, "&lt;&lt;$1&gt;&gt;");
  s = s.replace(LITERAL_MARKER_RE, "&lt;$1&gt;");
  // block boundaries become line breaks so a key name can never be
  // assembled from the end of one paragraph/cell and the start of the next
  s = s.replace(BLOCK_TAG_RE, "\n");
  s = s.replace(/<[^>]*>/g, "");
  s = decodeEntities(s);
  return s
    .replace(/ /g, " ")
    .replace(/[ \t\r]+/g, " ")
    .replace(/\s*\n\s*/g, "\n")
    .trim();
}

/* ───────────────────────────── analysis ───────────────────────────── */

const TEXT_BRACKET_RE = /(?:<<|«)\s*([^<>«»]+?)\s*(?:>>|»)/g;
const TEXT_AT_RE = /@([A-Za-z][A-Za-z0-9_]*)/g;
const TEXT_BARE_RE = /(?<![A-Za-z0-9_@])[A-Z][A-Z0-9_]{3,}(?![A-Za-z0-9_])/g;
const PLAIN_MIN_LENGTH = 6; // "GRADE" would otherwise eat every "Grade:" label

/** "STUDENTNAME" → /(?<![A-Za-z0-9])s[\s_]*t[\s_]*u…e(?![A-Za-z0-9])/gi */
export function plainTextRegex(norm: string): RegExp {
  // spaces / _ / - may sit between the letters, but never a line break
  // (the text projection uses \n for paragraph and table-cell boundaries)
  const loose = norm.split("").map(escapeRegExp).join("[ \\t\\u00a0_-]*");
  return new RegExp(`(?<![A-Za-z0-9])${loose}(?![A-Za-z0-9])`, "gi");
}

const TEXT_CHECKBOX_RE = /<\s*c\s*(\d+)\s*(c?)\s*>/gi;
const TEXT_TEXTBOX_RE = /<\s*t\s*(\d+)\s*>/gi;

export interface AnalyzeOptions {
  /** also look for key names written as plain text (default true) */
  plainText?: boolean;
}

export function analyzeEsignHtml(
  html: string,
  options: AnalyzeOptions = {},
): EsignAnalysis {
  return analyzeEsignText(htmlToText(html), options);
}

/**
 * Token / marker discovery on a plain-text projection. Shared by the HTML
 * analyser and the DOCX analyser (which builds its text from <w:t> runs).
 */
export function analyzeEsignText(
  text: string,
  options: AnalyzeOptions = {},
): EsignAnalysis {
  const tokens = new Map<string, EsignToken>();
  const addToken = (raw: string, form: EsignTokenForm) => {
    const norm = normalizeKey(raw);
    if (!norm) return;
    // plain-text variants are kept apart by exact spelling, so that
    // "STUDENT ID" (a placeholder) and "Student ID" (a label) can be
    // included / excluded independently
    const id = form === "plain" ? `plain:${raw}` : `${form}:${norm}`;
    const existing = tokens.get(id);
    if (existing) {
      existing.count++;
      return;
    }
    tokens.set(id, {
      raw,
      norm,
      form,
      knownKey: KEY_BY_NORM.get(norm) ?? null,
      count: 1,
    });
  };

  for (const m of text.matchAll(TEXT_BRACKET_RE)) {
    addToken(`<<${m[1].trim()}>>`, "bracket");
  }
  for (const m of text.matchAll(TEXT_AT_RE)) {
    addToken(m[0], "at");
  }
  const bracketNorms = new Set(
    Array.from(tokens.values())
      .filter((t) => t.form === "bracket")
      .map((t) => t.norm),
  );
  for (const m of text.matchAll(TEXT_BARE_RE)) {
    const norm = m[0];
    // bare words only count when they are a known key and the same key
    // was not already written in proper <<KEY>> form
    if (!KEY_BY_NORM.has(norm) || bracketNorms.has(norm)) continue;
    addToken(norm, "bare");
  }

  // plain-text key names: "studentname", "Student Name", "student_name"…
  // Only for keys not already used in a stricter form, names of 6+ chars,
  // whole words, and never inside a <<…>> / @… token that was found above.
  const strictNorms = new Set(Array.from(tokens.values()).map((t) => t.norm));
  const consumed: [number, number][] = [];
  for (const re of [TEXT_BRACKET_RE, TEXT_AT_RE]) {
    for (const m of text.matchAll(re)) {
      consumed.push([m.index ?? 0, (m.index ?? 0) + m[0].length]);
    }
  }
  // only CAPITAL words that really are keys count as consumed — "STUDENT"
  // in "STUDENT ID" must not hide the plain-text match "STUDENT ID"
  for (const m of text.matchAll(TEXT_BARE_RE)) {
    if (KEY_BY_NORM.has(m[0])) {
      consumed.push([m.index ?? 0, (m.index ?? 0) + m[0].length]);
    }
  }
  const overlaps = (s: number, e: number) =>
    consumed.some(([cs, ce]) => s < ce && e > cs);
  for (const [norm, keyName] of KEY_BY_NORM) {
    if (options.plainText === false) break;
    if (norm.length < PLAIN_MIN_LENGTH || strictNorms.has(norm)) continue;
    const re = plainTextRegex(norm);
    for (const m of text.matchAll(re)) {
      const s = m.index ?? 0;
      if (overlaps(s, s + m[0].length)) continue;
      addToken(m[0].replace(/\s+/g, " "), "plain");
      void keyName;
    }
  }

  const controls = new Map<string, EsignControl>();
  const addControl = (
    type: EsignControlType,
    digits: string,
    checked: boolean,
  ) => {
    const letter = type === "checkbox" ? "c" : "t";
    const id = `${letter}${digits}`;
    const marker = checked ? `${id}c` : id;
    const existing = controls.get(marker);
    if (existing) {
      existing.count++;
      return;
    }
    controls.set(marker, {
      marker,
      id,
      type,
      n: parseInt(digits, 10),
      checked,
      count: 1,
    });
  };
  for (const m of text.matchAll(TEXT_CHECKBOX_RE)) {
    addControl("checkbox", m[1], m[2] !== "");
  }
  for (const m of text.matchAll(TEXT_TEXTBOX_RE)) {
    addControl("textbox", m[1], false);
  }

  const tokenList = Array.from(tokens.values());
  const controlList = Array.from(controls.values()).sort((a, b) =>
    a.type === b.type ? a.n - b.n : a.type === "checkbox" ? -1 : 1,
  );

  const matchedKeys = Array.from(
    new Set(
      tokenList
        .filter((t) => t.knownKey)
        .map((t) => t.knownKey as string),
    ),
  );
  const unmatchedKeys = Array.from(
    new Set(tokenList.filter((t) => !t.knownKey).map((t) => t.raw)),
  );

  const checkboxes = controlList.filter((c) => c.type === "checkbox");
  const textboxes = controlList.filter((c) => c.type === "textbox");

  const idUse = new Map<string, number>();
  for (const c of controlList) idUse.set(c.id, (idUse.get(c.id) ?? 0) + c.count);
  const duplicateIds = Array.from(idUse.entries())
    .filter(([, n]) => n > 1)
    .map(([id]) => id);

  const missing = (list: EsignControl[]) => {
    const present = new Set(list.map((c) => c.n));
    const max = list.reduce((m, c) => Math.max(m, c.n), 0);
    const out: number[] = [];
    for (let i = 1; i <= max; i++) if (!present.has(i)) out.push(i);
    return out;
  };

  return {
    tokens: tokenList,
    controls: controlList,
    matchedKeys,
    unmatchedKeys,
    checkboxes,
    textboxes,
    duplicateIds,
    missingNumbers: { checkbox: missing(checkboxes), textbox: missing(textboxes) },
  };
}

/* ───────────────────────── token resolution ───────────────────────── */

export type ResolvedToken =
  | { kind: "placeholder"; key: string }
  | { kind: "text"; text: string };

function findMapping(
  raw: string,
  norm: string,
  mappings: Record<string, string>,
): string | undefined {
  if (mappings[raw]) return mappings[raw];
  const hit = Object.keys(mappings).find(
    (k) => mappings[k] && normalizeKey(k) === norm,
  );
  return hit ? mappings[hit] : undefined;
}

/**
 * Work out what a token should become:
 *   1. user mapping wins (placeholder → canonical key, plain text → text)
 *   2. otherwise a known key from keys.json
 *   3. otherwise null (leave the document untouched there)
 */
export function resolveToken(
  raw: string,
  mappings: Record<string, string> = {},
  visited: Set<string> = new Set(),
): ResolvedToken | null {
  const norm = normalizeKey(raw);
  if (!norm || visited.has(norm)) return null;
  visited.add(norm);

  const mapped = findMapping(raw, norm, mappings);
  if (mapped !== undefined) {
    if (isPlaceholderValue(mapped)) {
      const mNorm = normalizeKey(mapped);
      const known = KEY_BY_NORM.get(mNorm);
      if (known) return { kind: "placeholder", key: placeholderInner(KEYS[known]) };
      const deeper =
        mNorm !== norm ? resolveToken(mapped, mappings, visited) : null;
      return deeper ?? { kind: "placeholder", key: mNorm };
    }
    return { kind: "text", text: mapped.trim() };
  }

  const known = KEY_BY_NORM.get(norm);
  if (known) return { kind: "placeholder", key: placeholderInner(KEYS[known]) };
  return null;
}

export function renderResolved(
  res: ResolvedToken,
  keyOutput: EsignKeyOutput,
): string {
  if (res.kind === "text") return escapeHtml(res.text);
  return keyOutput === "encoded"
    ? `&lt;&lt;${res.key}&gt;&gt;`
    : `<<${res.key}>>`;
}

/* ─────────────────────── tolerant regex builders ─────────────────────── */

/** Anything that may sit between two characters of one token: tags Word
 *  splits runs with, non-breaking spaces, whitespace, or the _ / - people
 *  type inside key names (<<STUDENT_NAME>>). */
const GAP = "(?:<[^>]*>|&nbsp;|&#160;|[\\s_-])*";
/** Tags only (no whitespace) — used for bare words where spaces matter. */
const TAG_GAP = "(?:<[^>]*>)*";
const OPEN2 = `(?:<<|&lt;${GAP}&lt;|&#60;${GAP}&#60;|«|&laquo;|&#171;)`;
const CLOSE2 = `(?:>>|&gt;${GAP}&gt;|&#62;${GAP}&#62;|»|&raquo;|&#187;)`;
const OPEN1 = "(?:<|&lt;|&#60;)";
const CLOSE1 = "(?:>|&gt;|&#62;)";

const joinWith = (s: string, gap: string) =>
  s.split("").map(escapeRegExp).join(gap);

export function bracketRegex(norm: string): RegExp {
  return new RegExp(`${OPEN2}${GAP}${joinWith(norm, GAP)}${GAP}${CLOSE2}`, "gi");
}
export function atRegex(norm: string): RegExp {
  return new RegExp(`@${joinWith(norm, TAG_GAP)}(?![A-Za-z0-9_])`, "gi");
}
export function bareRegex(norm: string): RegExp {
  return new RegExp(
    `(?<![A-Za-z0-9_@])${joinWith(norm, TAG_GAP)}(?![A-Za-z0-9_])`,
    "g",
  );
}
export function freeTextRegex(text: string): RegExp {
  const words = text
    .trim()
    .split(/\s+/)
    .map((w) => joinWith(w, TAG_GAP))
    .join("(?:<[^>]*>|&nbsp;|&#160;|\\s)+");
  return new RegExp(`(?<![A-Za-z0-9])${words}(?![A-Za-z0-9])`, "gi");
}
export function controlRegex(control: EsignControl): RegExp {
  const letter = control.type === "checkbox" ? "c" : "t";
  const digits = joinWith(String(control.id.slice(1)), GAP);
  const flag = control.checked ? `c${GAP}` : "";
  return new RegExp(
    `${OPEN1}${GAP}${letter}${GAP}${digits}${GAP}${flag}${CLOSE1}`,
    "gi",
  );
}

/* ───────────────────────── rewrite machinery ───────────────────────── */

interface RewriteRule {
  regex: RegExp;
  /** returns the replacement for the token text (tags are re-attached automatically) */
  render: (match: string) => string;
  onHit?: () => void;
}

function isInsideTag(chunk: string, offset: number): boolean {
  const lt = chunk.lastIndexOf("<", offset - 1);
  const gt = chunk.lastIndexOf(">", offset - 1);
  return lt > gt;
}

/** Split a matched token into its literal brackets and the part between. */
function splitToken(match: string): {
  open: string;
  inner: string;
  close: string;
} {
  let inner = match;
  let open = "";
  let close = "";
  if (inner.startsWith("<<")) {
    open = "<<";
    inner = inner.slice(2);
  } else if (inner.startsWith("<")) {
    open = "<";
    inner = inner.slice(1);
  }
  if (inner.endsWith(">>")) {
    close = ">>";
    inner = inner.slice(0, -2);
  } else if (inner.endsWith(">")) {
    close = ">";
    inner = inner.slice(0, -1);
  }
  return { open, inner, close };
}

/** Real tags that sat inside a matched token (Word run boundaries). */
function tagsInside(match: string): string {
  return (splitToken(match).inner.match(/<[^>]*>/g) ?? []).join("");
}

/** Human-readable text of a matched token: tags removed, entities decoded. */
function tokenText(match: string): string {
  const { open, inner, close } = splitToken(match);
  return decodeEntities(open + inner.replace(/<[^>]*>/g, "") + close).replace(
    /\s+/g,
    " ",
  );
}

function applyRules(chunk: string, rules: RewriteRule[]): string {
  let out = chunk;
  for (const rule of rules) {
    rule.regex.lastIndex = 0;
    if (!rule.regex.test(out)) continue;
    rule.regex.lastIndex = 0;
    const source = out;
    out = source.replace(rule.regex, (match: string, ...args: unknown[]) => {
      const offset = args[args.length - 2] as number;
      if (isInsideTag(source, offset)) return match;
      rule.onHit?.();
      return rule.render(match) + tagsInside(match);
    });
  }
  return out;
}

/** Run `fn` over every part of the HTML that is not a comment/style/script. */
function rewriteUnprotected(html: string, fn: (chunk: string) => string): string {
  let out = "";
  let last = 0;
  for (const m of html.matchAll(PROTECTED_RE)) {
    const idx = m.index ?? 0;
    out += fn(html.slice(last, idx)) + m[0];
    last = idx + m[0].length;
  }
  return out + fn(html.slice(last));
}

function renderTemplate(template: string, control: EsignControl): string {
  return template
    .replace(/\{id\}/g, control.id)
    .replace(/\{marker\}/g, control.marker)
    .replace(/\{n\}/g, String(control.n))
    .replace(/\{type\}/g, control.type)
    .replace(/\{checked\}/g, control.checked ? "checked" : "");
}

function fixMetaCharset(html: string): string {
  return html.replace(/<meta\b[^>]*>/gi, (tag) =>
    tag.replace(/charset=["']?[\w-]+/i, (m) => m.replace(/[\w-]+$/, "utf-8")),
  );
}

/* ─────────────────────── key rules (shared) ─────────────────────── */

const FORM_ORDER: Record<EsignTokenForm, number> = {
  bracket: 0,
  at: 1,
  bare: 2,
  plain: 3,
};

/** plain-text key inside HTML: tags / whitespace / _ allowed between letters */
export function plainHtmlRegex(raw: string): RegExp {
  // exact spelling of this variant ("Student ID" ≠ "STUDENT ID"); only the
  // separators between letters are flexible
  const letters = raw.replace(/[\s_-]+/g, "");
  // only INLINE tags may sit between the letters (Word / docx-preview split
  // runs into <span>s); a block boundary (</p>, <td> …) ends the match so
  // "STUDENT" at the end of a heading never joins "Name" in the next cell
  const inlineGap =
    "(?:<\\/?(?:span|b|i|u|em|strong|font|sup|sub|small|big|s|strike|o:p|st1:[a-z]+)\\b[^>]*>|&nbsp;|&#160;|[ \\t\\u00a0_-])*";
  const loose = letters.split("").map(escapeRegExp).join(inlineGap);
  return new RegExp(`(?<![A-Za-z0-9])${loose}(?![A-Za-z0-9])`, "g");
}

function regexForToken(form: EsignTokenForm, norm: string, raw: string): RegExp {
  if (form === "bracket") return bracketRegex(norm);
  if (form === "at") return atRegex(norm);
  if (form === "plain") return plainHtmlRegex(raw);
  return bareRegex(norm);
}

function regexForMappingKey(key: string): RegExp {
  const trimmed = key.trim();
  if (/^(?:<<|&lt;&lt;|«).*(?:>>|&gt;&gt;|»)$/i.test(trimmed)) {
    return bracketRegex(normalizeKey(trimmed));
  }
  if (/^@[A-Za-z0-9_]+$/.test(trimmed)) return atRegex(normalizeKey(trimmed));
  return freeTextRegex(trimmed);
}

/**
 * Every key-ish thing we should look for, longest first, with the regex
 * that finds it in raw HTML. Discovered tokens come first; user mappings
 * whose key was not discovered (free text like "student name") follow.
 */
export interface KeyTargetSpec {
  raw: string;
  norm: string;
  form: EsignTokenForm | "text";
  fromMapping: boolean;
}

/**
 * Regex-free list of everything key-ish to search for, longest first:
 * discovered tokens, then user mappings whose key was not discovered
 * (free text like "student name"). Shared by the HTML and DOCX engines.
 */
export function keyTargetSpecs(
  analysis: EsignAnalysis,
  mappings: Record<string, string>,
): KeyTargetSpec[] {
  const targets: KeyTargetSpec[] = analysis.tokens
    .slice()
    .sort(
      (a, b) =>
        b.norm.length - a.norm.length || FORM_ORDER[a.form] - FORM_ORDER[b.form],
    )
    .map((t) => ({ raw: t.raw, norm: t.norm, form: t.form, fromMapping: false }));

  // A mapping key only duplicates a discovered token when it is the same
  // *form* — "student name" (free text) must still be searched for even
  // though <<STUDENTNAME>> was found in bracket form.
  const discovered = new Set(analysis.tokens.map((t) => `${t.form}:${t.norm}`));
  const formOfMappingKey = (k: string): EsignTokenForm | "text" => {
    const t = k.trim();
    if (/^(?:<<|&lt;&lt;|«).*(?:>>|&gt;&gt;|»)$/i.test(t)) return "bracket";
    if (/^@[A-Za-z0-9_]+$/.test(t)) return "at";
    if (/^[A-Z][A-Z0-9_]{3,}$/.test(t)) return "bare";
    return "text";
  };
  const extra: KeyTargetSpec[] = Object.keys(mappings)
    .filter(
      (k) =>
        k.trim() &&
        mappings[k] &&
        !discovered.has(`${formOfMappingKey(k)}:${normalizeKey(k)}`),
    )
    .sort((a, b) => b.length - a.length)
    .map((k) => ({
      raw: k,
      norm: normalizeKey(k),
      form: formOfMappingKey(k),
      fromMapping: true,
    }));

  return [...targets, ...extra];
}

function collectKeyTargets(
  analysis: EsignAnalysis,
  mappings: Record<string, string>,
): { raw: string; norm: string; regex: RegExp; fromMapping: boolean }[] {
  return keyTargetSpecs(analysis, mappings).map((t) => ({
    raw: t.raw,
    norm: t.norm,
    fromMapping: t.fromMapping,
    regex: t.fromMapping
      ? regexForMappingKey(t.raw)
      : regexForToken(t.form as EsignTokenForm, t.norm, t.raw),
  }));
}

/* ───────────────────────────── transform ───────────────────────────── */

export function transformEsignHtml(
  html: string,
  analysis: EsignAnalysis,
  options: EsignReplaceOptions = {},
): EsignReplaceResult {
  const mappings = options.mappings ?? {};
  const excluded = new Set(options.exclude ?? []);
  const keyOutput: EsignKeyOutput = options.keyOutput ?? "encoded";
  const replaceKeys = options.replaceKeys ?? true;
  const replaceControls = options.replaceControls ?? true;
  const templates: EsignTemplates = { ...DEFAULT_TEMPLATES, ...options.templates };

  const stats: EsignReplaceStats = { keys: {}, controls: {}, skipped: [] };
  const rules: RewriteRule[] = [];

  if (replaceKeys) {
    for (const target of collectKeyTargets(analysis, mappings)) {
      if (excluded.has(target.raw)) continue;
      const resolved = resolveToken(target.raw, mappings);
      if (!resolved) {
        if (!target.fromMapping) stats.skipped.push(target.raw);
        continue;
      }
      const replacement = renderResolved(resolved, keyOutput);
      const statKey =
        resolved.kind === "placeholder" ? resolved.key : `"${resolved.text}"`;
      rules.push({
        regex: target.regex,
        render: () => replacement,
        onHit: () => {
          stats.keys[statKey] = (stats.keys[statKey] ?? 0) + 1;
        },
      });
    }
  }

  if (replaceControls) {
    for (const control of analysis.controls) {
      const template =
        control.type === "textbox"
          ? templates.textbox
          : control.checked
            ? templates.checkboxChecked
            : templates.checkbox;
      const replacement = renderTemplate(template, control);
      rules.push({
        regex: controlRegex(control),
        render: () => replacement,
        onHit: () => {
          stats.controls[control.marker] =
            (stats.controls[control.marker] ?? 0) + 1;
        },
      });
    }
  }

  // Pure string rewrite: no DOM parsing / re-serialisation, so everything
  // that is not a matched token (tags, styles, spacing, line endings) is
  // returned exactly as uploaded.
  const out = rewriteUnprotected(html, (chunk) => applyRules(chunk, rules));
  return { html: out, stats };
}

/* ───────────────────────────── highlight ───────────────────────────── */

export const ESIGN_HIGHLIGHT_CSS = `
  mark.es-key-green  { background:#bbf7d0; color:#14532d; border-bottom:2px solid #22c55e; }
  mark.es-key-yellow { background:#fef08a; color:#713f12; border-bottom:2px solid #eab308; }
  mark.es-key-blue   { background:#dbeafe; color:#1e3a8a; border-bottom:2px solid #3b82f6; }
  mark.es-key-plain  { background:#ffedd5; color:#7c2d12; border-bottom:2px solid #f97316; }
  mark.es-key-excluded { background:#fff7ed; color:#9a3412; border-bottom:2px dashed #fdba74; font-weight:500; }
  mark.es-ctrl-checkbox { background:#e0f2fe; color:#0c4a6e; border:1px solid #0ea5e9; }
  mark.es-ctrl-textbox  { background:#ede9fe; color:#4c1d95; border:1px solid #8b5cf6; }
  mark[class^="es-"] { border-radius:3px; padding:0 3px; font-weight:600; font-family:Consolas,monospace; font-size:0.9em; white-space:nowrap; }
  input[type="checkbox"] { outline:2px solid #0ea5e9; outline-offset:2px; }
  input[type="text"] { outline:2px solid #8b5cf6; outline-offset:1px; }
`;

/**
 * Preview-only: wrap every key / marker in a coloured <mark> so the user
 * can see what was detected. Structure-preserving (same tag re-attachment
 * as the real transform), never used for the downloaded file.
 */
export function highlightEsignHtml(
  html: string,
  analysis: EsignAnalysis,
  mappings: Record<string, string> = {},
  exclude: string[] = [],
): string {
  const rules: RewriteRule[] = [];
  const excluded = new Set(exclude);

  for (const target of collectKeyTargets(analysis, mappings)) {
    const token = analysis.tokens.find((t) => t.raw === target.raw);
    const resolved = resolveToken(target.raw, mappings);
    const cls = excluded.has(target.raw)
      ? "es-key-excluded"
      : token?.form === "plain"
        ? "es-key-plain"
        : token?.knownKey
          ? "es-key-green"
          : resolved
            ? "es-key-blue"
            : "es-key-yellow";
    rules.push({
      regex: target.regex,
      render: (match) =>
        `<mark class="${cls}">${escapeHtml(tokenText(match))}</mark>`,
    });
  }

  for (const control of analysis.controls) {
    rules.push({
      regex: controlRegex(control),
      render: () =>
        `<mark class="es-ctrl-${control.type}">&lt;${control.marker}&gt;</mark>`,
    });
  }

  return rewriteUnprotected(html, (chunk) => applyRules(chunk, rules));
}
