import { type NextRequest, NextResponse } from "next/server";
import keys from "../../../keys.json";

/* ──────────────────────────────────────────────────────────────────────
 * HTML/HTM key-replacement API
 *
 * Why a separate endpoint?
 *   The existing /api/replace handles both DOCX and HTML in one route.
 *   HTML files often store placeholders in two equivalent forms:
 *     • literal:  <<PARENTORCURRENTCITYSTATEZIP>>
 *     • encoded:  &lt;&lt;PARENTORCURRENTCITYSTATEZIP&gt;&gt;
 *   The existing route only matches the literal form in HTML (the
 *   encoded-form regex is gated to non-HTML, where the source is
 *   docx XML, not HTML).  This endpoint exists so HTML/HTM files can
 *   have BOTH forms detected and rewritten correctly, without touching
 *   any of the existing DOCX/HTML logic in /api/replace.
 *
 * Behaviour (mirrors /api/replace semantics for consistency):
 *   1. User-supplied keyMappings are honoured first.
 *   2. If a mapping value is itself a <<…>> / &lt;&lt;…&gt;&gt; placeholder
 *      that exists in keys.json, the keys.json value (or the user's
 *      mapping when present) is used.
 *   3. The replacement is always emitted in the SAME form as the
 *      matched placeholder — literal `<<…>>` stays literal, encoded
 *      `&lt;&lt;…&gt;&gt;` stays encoded — so the HTML structure of
 *      the file is preserved.
 * ────────────────────────────────────────────────────────────────────── */

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Strip <<>>, &lt;&lt;>>&gt;, @, whitespace, uppercase → "CAMPUSNAME" */
function normalizeKey(raw: string): string {
  return raw
    .replace(/^<</, "")
    .replace(/>>$/, "")
    .replace(/^&lt;&lt;/, "")
    .replace(/&gt;&gt;$/, "")
    .replace(/^@/, "")
    .replace(/\s+/g, "")
    .toUpperCase()
    .trim();
}

const typedKeys = keys as Record<string, string>;
const knownKeysList = Object.keys(typedKeys);

function isKnownKeyToken(token: string): boolean {
  if (typedKeys[token]) return true;
  const norm = normalizeKey(token);
  return knownKeysList.some((k) => normalizeKey(k) === norm);
}

function isPlaceholder(v: string): boolean {
  if (v.startsWith("<<") && v.endsWith(">>")) return true;
  if (v.startsWith("&lt;&lt;") && v.endsWith("&gt;&gt;")) return true;
  return false;
}

/**
 * Recursively resolve a token to its final value.
 * Mirrors the contract used by /api/replace so that downstream callers
 * (e.g. highlighting) see consistent values.
 */
function resolveValueFully(
  raw: string,
  keyMappings: Record<string, string>,
  visited = new Set<string>(),
): string | null {
  if (visited.has(raw)) return null;
  visited.add(raw);

  let resolved: string | null = null;

  // 1. User mapping wins
  if (keyMappings[raw]) {
    resolved = keyMappings[raw];
  } else {
    const norm = normalizeKey(raw);
    const mappingHit = Object.keys(keyMappings).find(
      (k) => normalizeKey(k) === norm,
    );
    if (mappingHit) resolved = keyMappings[mappingHit];
  }

  // 2. Fallback to keys.json
  if (resolved === null) {
    if (typedKeys[raw]) resolved = typedKeys[raw];
    else {
      const norm = normalizeKey(raw);
      const hit = knownKeysList.find((k) => normalizeKey(k) === norm);
      if (hit) resolved = typedKeys[hit];
    }
  }

  if (resolved === null) return null;

  // 3. If still a placeholder, recurse
  if (isPlaceholder(resolved)) {
    if (normalizeKey(resolved) === normalizeKey(raw)) return null;
    const inner = resolveValueFully(resolved, keyMappings, visited);
    return inner ?? resolved;
  }

  return resolved;
}

/** Escape a value that will be inserted into HTML as plain text. */
function escapeHtmlText(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File;
    const keyMappingsStr = formData.get("keyMappings") as string;
    const foundKeysStr = formData.get("foundKeys") as string;
    const unmatchedKeysStr = formData.get("unmatchedKeys") as string;

    const keyMappings: Record<string, string> = keyMappingsStr
      ? JSON.parse(keyMappingsStr)
      : {};
    const foundKeys: string[] = foundKeysStr ? JSON.parse(foundKeysStr) : [];
    const unmatchedKeys: string[] = unmatchedKeysStr
      ? JSON.parse(unmatchedKeysStr)
      : [];

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith(".html") && !fileName.endsWith(".htm")) {
      return NextResponse.json(
        { error: "This endpoint only handles .html / .htm files" },
        { status: 400 },
      );
    }

    const htmlContent = await file.text();

    // Build the working set of tokens: anything the client told us about
    // (found + unmatched + mapping keys) PLUS every key defined in
    // keys.json that could appear in the file.  We sort longest-first so
    // e.g. <<PARENTORCURRENTCITYSTATEZIP>> is matched before <<PARENT>>.
    const tokens = Array.from(
      new Set([
        ...foundKeys,
        ...unmatchedKeys,
        ...Object.keys(keyMappings),
        ...knownKeysList,
      ]),
    ).sort((a, b) => normalizeKey(b).length - normalizeKey(a).length);

    let out = htmlContent;
    const replacedKeys: string[] = [];

    for (const token of tokens) {
      // Resolve what the user (or keys.json) wants this token to become.
      const value = resolveValueFully(token, keyMappings);
      if (value === null) continue;

      // Build the two regexes once per token (with whitespace tolerance,
      // so <<PARENT OR CURRENT CITY STATE ZIP>> also matches).
      const literalInner = normalizeKey(token)
        .split("")
        .map(escapeRegExp)
        .join("[\\s]*");
      const reLiteral = new RegExp(`<<\\s*${literalInner}\\s*>>`, "gi");
      const reEncoded = new RegExp(
        `&lt;&lt;\\s*${literalInner}\\s*&gt;&gt;`,
        "gi",
      );

      // Decide what to write in place of the matched placeholder.
      //   • If the resolved value is itself a placeholder, keep it as-is
      //     (so a mapping like <<GRADE>> → <<STUDENTGRADE>> round-trips).
      //   • If the resolved value is plain text, emit it HTML-escaped
      //     so the result is always safe to drop into a text node.
      const replacement = isPlaceholder(value) ? value : escapeHtmlText(value);

      if (reLiteral.test(out)) {
        reLiteral.lastIndex = 0;
        out = out.replace(reLiteral, replacement);
        if (!replacedKeys.includes(normalizeKey(token))) {
          replacedKeys.push(normalizeKey(token));
        }
      }

      if (reEncoded.test(out)) {
        reEncoded.lastIndex = 0;
        out = out.replace(reEncoded, replacement);
        if (!replacedKeys.includes(normalizeKey(token))) {
          replacedKeys.push(normalizeKey(token));
        }
      }
    }

    return new NextResponse(out, {
      headers: {
        "Content-Type": "text/html; charset=utf-8",
        "Content-Disposition": `attachment; filename="${file.name}"`,
        // Expose the list of replaced keys so the client can re-analyse
        // without needing to parse the response body.
        "X-Replaced-Keys": JSON.stringify(replacedKeys),
      },
    });
  } catch (error) {
    console.error("Error processing HTML document:", error);
    return NextResponse.json(
      { error: "Failed to process HTML document" },
      { status: 500 },
    );
  }
}
