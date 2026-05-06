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
      // Handle HTML/HTM files
      const htmlContent = await file.text();
      let processedHtml = htmlContent;

      console.log("[v0] HTML file tokens to process:", allTokens);
      console.log("[v0] HTML content length:", htmlContent.length);

      allTokens.forEach((token) => {
        const norm = normalizeKey(token);
        let replaced = false;

        // For HTML files, replace plain text keys with placeholder format <<KEY>>
        const htmlReplaceValue = `<<${norm}>>`;
        
        console.log("[v0] Processing HTML token:", token, "→ norm:", norm, "→ replace with:", htmlReplaceValue);

        // =====================
        // Pattern A: <<...>>
        // =====================
        if (norm.length > 0 && !replaced) {
          try {
            const inner = norm.split("").map(escapeRegExp).join("[\\s]*");
            const reLiteral = new RegExp(`<<\\s*${inner}\\s*>>`, "gi");
            if (reLiteral.test(processedHtml)) {
              processedHtml = processedHtml.replace(reLiteral, htmlReplaceValue);
              replaced = true;
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
            if (re.test(processedHtml)) {
              processedHtml = processedHtml.replace(re, htmlReplaceValue);
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
            console.log("[v0] Testing Pattern C (ALL CAPS) for:", norm);
            if (re.test(processedHtml)) {
              console.log("[v0] Pattern C matched! Replacing with:", htmlReplaceValue);
              processedHtml = processedHtml.replace(re, htmlReplaceValue);
              replaced = true;
            } else {
              console.log("[v0] Pattern C did not match for:", norm);
            }
          } catch (e) {
            console.log("[v0] Pattern C error:", e);
          }
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
              const re = new RegExp(
                `(?<![A-Za-z0-9])${escapeRegExp(freeText)}(?![A-Za-z0-9])`,
                "gi",
              );
              if (re.test(processedHtml)) {
                processedHtml = processedHtml.replace(re, htmlReplaceValue);
                replaced = true;
              }
            } catch (_) {}
          }
        }
      });

      return new NextResponse(processedHtml, {
        headers: {
          "Content-Type": "text/html; charset=utf-8",
          "Content-Disposition": `attachment; filename="${file.name}"`,
        },
      });
    } else {
      // Handle DOCX files (original logic)
      const arrayBuffer = await file.arrayBuffer();
      const zip = new PizZip(arrayBuffer);

      const xmlFiles = Object.keys(zip.files).filter((f) =>
        f.match(/word\/(document|header\d*|footer\d*)\.xml/),
      );

      xmlFiles.forEach((xmlPath) => {
        let xml = zip.files[xmlPath].asText();

        allTokens.forEach((token) => {
          const value = resolveValueFully(token, keyMappings);
          if (!value) return;

          const escaped = escapeXml(value);
          const norm = normalizeKey(token);
          let replaced = false;

          // =====================
          // Pattern A: <<...>>  — Handle BOTH literal AND XML-encoded brackets
          // =====================
          if (norm.length > 0 && !replaced) {
            try {
              const inner = norm.split("").map(escapeRegExp).join("[\\s]*");
              const finalValue = unwrapPlaceholder(escaped);

              // 1. Try literal angle brackets: <<TOKEN>>
              const reLiteral = new RegExp(`<<\\s*${inner}\\s*>>`, "gi");
              if (reLiteral.test(xml)) {
                xml = xml.replace(reLiteral, `<<${finalValue}>>`);
                replaced = true;
              }

              // 2. Try XML-encoded angle brackets: &lt;&lt;TOKEN&gt;&gt;
              if (!replaced) {
                const reEncoded = new RegExp(
                  `&lt;&lt;\\s*${inner}\\s*&gt;&gt;`,
                  "gi",
                );
                if (reEncoded.test(xml)) {
                  xml = xml.replace(reEncoded, `&lt;&lt;${finalValue}&gt;&gt;`);
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
              if (re.test(xml)) {
                xml = xml.replace(re, escaped);
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
              if (re.test(xml)) {
                xml = xml.replace(re, escaped);
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
                const re = new RegExp(
                  `(?<![A-Za-z0-9])${escapeRegExp(freeText)}(?![A-Za-z0-9])`,
                  "gi",
                );
                if (re.test(xml)) {
                  xml = xml.replace(re, escaped);
                  replaced = true;
                }
              } catch (_) {}
            }
          }
        });

        zip.file(xmlPath, xml);
      });

      const buffer = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });

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
