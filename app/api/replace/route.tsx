import { type NextRequest, NextResponse } from "next/server";
import PizZip from "pizzip";
import keys from "../../../keys.json";

function escapeRegExp(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeXml(s: string) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Strip <<>>, @, spaces, lowercase → "CAMPUSNAME" */
function normalizeKey(raw: string): string {
  return raw
    .replace(/^<</, "")
    .replace(/>>$/, "")
    .replace(/^@/, "")
    .replace(/\s+/g, "")
    .toUpperCase()
    .trim();
}

const typedKeys = keys as Record<string, string>;
const knownKeysList = Object.keys(typedKeys);

/**
 * Recursively resolve a token to its final value.
 */
function resolveValueFully(
  raw: string,
  keyMappings: Record<string, string>,
  visited = new Set<string>(),
): string | null {
  if (visited.has(raw)) return null;
  visited.add(raw);

  let resolved: string | null = null;

  // 1. Direct key lookup in keys.json
  if (typedKeys[raw]) resolved = typedKeys[raw];
  else {
    const norm = normalizeKey(raw);
    const hit = knownKeysList.find((k) => normalizeKey(k) === norm);
    if (hit) resolved = typedKeys[hit];
  }

  // 2. Custom mapping (direct or normalized)
  if (resolved === null && keyMappings[raw]) {
    resolved = keyMappings[raw];
  }
  if (resolved === null) {
    const norm = normalizeKey(raw);
    const mappingHit = Object.keys(keyMappings).find(
      (k) => normalizeKey(k) === norm,
    );
    if (mappingHit) resolved = keyMappings[mappingHit];
  }

  if (resolved === null) return null;

  // 3. If the resolved value is still a placeholder, resolve it recursively
  if (resolved.startsWith("<<") && resolved.endsWith(">>")) {
    const innerResolved = resolveValueFully(resolved, keyMappings, visited);
    return innerResolved ?? resolved;
  }

  return resolved;
}

/** Strip outer << >> or &lt;&lt; &gt;&gt; from a value */
function unwrapPlaceholder(val: string): string {
  if (val.startsWith("&lt;&lt;") && val.endsWith("&gt;&gt;")) {
    return val.slice(8, -8);
  }
  if (val.startsWith("<<") && val.endsWith(">>")) {
    return val.slice(2, -2);
  }
  return val;
}

/**
 * Apply Pattern A/B/C/D token replacement against a plain (already-escaped) text string.
 * This is the same matching logic as before, but now operates on a paragraph's
 * MERGED text instead of the raw XML file — so tokens split across multiple
 * <w:r> runs (e.g. "(SCHOOL NAME)" split into "(SCHOOL " + "NAME)") still match.
 */
function applyTokenReplacements(
  text: string,
  allTokens: string[],
  keyMappings: Record<string, string>,
  isHtml: boolean,
): string {
  let out = text;

  allTokens.forEach((token) => {
    const norm = normalizeKey(token);
    let replaced = false;

    const value = isHtml ? null : resolveValueFully(token, keyMappings);
    if (!isHtml && !value) return;

    const escaped = isHtml ? "" : escapeXml(value as string);
    const htmlReplaceValue = `<<${norm}>>`;

    // =====================
    // Pattern A: <<...>>  (literal + XML-encoded)
    // =====================
    if (norm.length > 0 && !replaced) {
      try {
        const inner = norm.split("").map(escapeRegExp).join("[\\s]*");

        const reLiteral = new RegExp(`<<\\s*${inner}\\s*>>`, "gi");
        if (reLiteral.test(out)) {
          out = out.replace(
            reLiteral,
            isHtml ? htmlReplaceValue : `<<${unwrapPlaceholder(escaped)}>>`,
          );
          replaced = true;
        }

        if (!replaced && !isHtml) {
          const reEncoded = new RegExp(
            `&lt;&lt;\\s*${inner}\\s*&gt;&gt;`,
            "gi",
          );
          if (reEncoded.test(out)) {
            out = out.replace(
              reEncoded,
              `&lt;&lt;${unwrapPlaceholder(escaped)}&gt;&gt;`,
            );
            replaced = true;
          }
        }
      } catch (_) {}
    }

    // =====================
    // Pattern B: @token
    // =====================
    if (!replaced && /^[A-Za-z0-9_]+$/.test(token.replace(/^@/, ""))) {
      try {
        const re = new RegExp(
          `@${escapeRegExp(token.replace(/^@/, ""))}`,
          "gi",
        );
        if (re.test(out)) {
          out = out.replace(re, isHtml ? htmlReplaceValue : escaped);
          replaced = true;
        }
      } catch (_) {}
    }

    // =====================
    // Pattern C: ALL CAPS (e.g., CAMPUSNAME)
    // =====================
    if (!replaced && /^[A-Z][A-Z0-9_]{3,}$/.test(norm)) {
      try {
        const re = new RegExp(
          `(?<![A-Za-z0-9_])${escapeRegExp(norm)}(?![A-Za-z0-9_])`,
          "g",
        );
        if (re.test(out)) {
          out = out.replace(re, isHtml ? htmlReplaceValue : escaped);
          replaced = true;
        }
      } catch (_) {}
    }

    // =====================
    // Pattern D: Free text (exact word boundary match)
    // =====================
    if (!replaced) {
      const freeText = token
        .replace(/^<</, "")
        .replace(/>>$/, "")
        .replace(/^@/, "")
        .trim();
      if (freeText.length > 0) {
        try {
          // allow whitespace variance between words (handles split-run spacing quirks)
          const innerFree = freeText
            .split(/\s+/)
            .map((w) => w.split("").map(escapeRegExp).join("[\\s]*"))
            .join("[\\s]+");
          const re = new RegExp(
            `(?<![A-Za-z0-9])${innerFree}(?![A-Za-z0-9])`,
            "gi",
          );
          if (re.test(out)) {
            out = out.replace(re, isHtml ? htmlReplaceValue : escaped);
            replaced = true;
          }
        } catch (_) {}
      }
    }
  });

  return out;
}

/**
 * Merge all <w:t> nodes within each <w:p>...</w:p> paragraph into one logical
 * string, run token replacement against that merged string, then write the
 * result into the FIRST <w:t> node and empty out the rest. This fixes cases
 * where Word splits a single visible phrase (e.g. "(SCHOOL NAME)") across
 * multiple runs, which previously made it invisible to regex matching.
 */
function replaceTokensInDocxXml(
  xml: string,
  allTokens: string[],
  keyMappings: Record<string, string>,
): string {
  // Normalize self-closing <w:t/> to <w:t></w:t> so they're uniformly handled.
  // IMPORTANT: must not match on <w:tab/>, <w:tblPr/>, <w:trPr/>, <w:tcPr/>,
  // <w:titlePgBelow/>, etc. — anything starting with "w:t" but NOT the exact
  // <w:t> tag. The (?![a-zA-Z]) lookahead prevents matching those substrings.
  let normalized = xml.replace(/<w:t(?![a-zA-Z])([^>]*)\/>/g, "<w:t$1></w:t>");

  normalized = normalized.replace(
    /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g,
    (paragraph) => {
      const wtRegex = /<w:t(?![a-zA-Z])([^>]*)>([\s\S]*?)<\/w:t>/g;
      const matches = Array.from(paragraph.matchAll(wtRegex));

      if (matches.length === 0) return paragraph;

      const fullText = matches.map((m) => m[2]).join("");
      const newFullText = applyTokenReplacements(
        fullText,
        allTokens,
        keyMappings,
        false,
      );

      if (newFullText === fullText) return paragraph;

      let result = "";
      let lastIndex = 0;

      matches.forEach((m, i) => {
        const matchIndex = m.index ?? 0;
        result += paragraph.slice(lastIndex, matchIndex);

        let attrs = m[1] || "";
        if (i === 0 && !/xml:space=/.test(attrs)) {
          attrs = `${attrs} xml:space="preserve"`;
        }

        const content = i === 0 ? newFullText : "";
        result += `<w:t${attrs}>${content}</w:t>`;

        lastIndex = matchIndex + m[0].length;
      });

      result += paragraph.slice(lastIndex);
      return result;
    },
  );

  return normalized;
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
    const isHtmlFile = fileName.endsWith(".html") || fileName.endsWith(".htm");

    const allTokens = [
      ...new Set([...foundKeys, ...unmatchedKeys, ...Object.keys(keyMappings)]),
    ].sort((a, b) => b.length - a.length);

    if (isHtmlFile) {
      // Handle HTML/HTM files (unchanged — HTML text isn't split across runs)
      const htmlContent = await file.text();
      let processedHtml = applyTokenReplacements(
        htmlContent,
        allTokens,
        keyMappings,
        true,
      );

      // Post-processing: Replace remaining empty <> placeholders
      const emptyBracketPatterns = [
        />\s*<>\s*</g,
        />\s*<\s*>\s*</g,
        /<\s*>\s*/g,
      ];

      let keyIndex = 0;
      for (const pattern of emptyBracketPatterns) {
        if (keyIndex >= allTokens.length) break;

        processedHtml = processedHtml.replace(pattern, (match) => {
          if (keyIndex < allTokens.length) {
            const token = allTokens[keyIndex];
            const norm = normalizeKey(token);
            keyIndex++;

            if (pattern.source.includes("><")) {
              return `><<${norm}>><`;
            }
            return `<<${norm}>>`;
          }
          return match;
        });
      }

      return new NextResponse(processedHtml, {
        headers: {
          "Content-Type": "text/html; charset=utf-8",
          "Content-Disposition": `attachment; filename="${file.name}"`,
        },
      });
    } else {
      // Handle DOCX files — now with paragraph-level run merging
      const arrayBuffer = await file.arrayBuffer();
      const zip = new PizZip(arrayBuffer);

      const xmlFiles = Object.keys(zip.files).filter((f) =>
        f.match(/word\/(document|header\d*|footer\d*)\.xml/),
      );

      xmlFiles.forEach((xmlPath) => {
        const xml = zip.files[xmlPath].asText();
        const newXml = replaceTokensInDocxXml(xml, allTokens, keyMappings);
        zip.file(xmlPath, newXml);
      });

      const buffer = zip.generate({
        type: "nodebuffer",
        compression: "DEFLATE",
      });

      return new NextResponse(new Uint8Array(buffer), {
        headers: {
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "Content-Disposition": `attachment; filename="${file.name}"`,
        },
      });
    }
  } catch (error) {
    console.error("Error processing document:", error);
    return NextResponse.json(
      { error: "Failed to process document" },
      { status: 500 },
    );
  }
}
