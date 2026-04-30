import { type NextRequest, NextResponse } from "next/server";
import PizZip from "pizzip";
import keys from "../../../keys.json";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File;

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    const arrayBuffer = await file.arrayBuffer();
    const zip = new PizZip(arrayBuffer);

    // Collect text from main doc + headers + footers
    let allText = "";
    const xmlFiles = Object.keys(zip.files).filter((f) =>
      f.match(/word\/(document|header\d*|footer\d*)\.xml/),
    );

    xmlFiles.forEach((xmlPath) => {
      const xml = zip.file(xmlPath)?.asText();
      if (xml) {
        // First, extract <<…>> tokens with spaces preserved BEFORE stripping tags
        const doubleBracketMatches = xml.match(/<<[^<>]+>>/g) || [];
        if (doubleBracketMatches.length) {
          allText += " " + doubleBracketMatches.join(" ");
        }

        const textContent = xml
          .replace(/<[^>]*>/g, " ")
          .replace(/\s+/g, " ")
          .trim();
        allText += " " + textContent;
      }
    });

    const text = allText.trim();

    const matchedKeys: string[] = [];
    const unmatchedKeys: string[] = [];

    const knownKeys = Object.keys(keys as Record<string, string>);

    // ── Helper: try to match a raw token against known keys ────────────────
    function normalizeToken(raw: string): string {
      return raw
        .replace(/^<</, "")
        .replace(/>>$/, "")
        .replace(/^@/, "")
        .replace(/\s+/g, "") // strip spaces so <<FIRST NAME>> → FIRSTNAME
        .toUpperCase()
        .trim();
    }

    function matchToken(raw: string): { matched: boolean; key: string } {
      const normalized = normalizeToken(raw);
      const hit = knownKeys.find(
        (k) => k.replace(/\s+/g, "").toUpperCase() === normalized,
      );
      if (hit) return { matched: true, key: hit };
      return { matched: false, key: raw };
    }

    // ── 1. <<…>> patterns (handles spaces, mixed case) ────────────────────
    // Re-extract from combined text (catches tokens we appended above)
    const doubleBracketPattern = /<<([^<>]+)>>/g;
    let m: RegExpExecArray | null;
    while ((m = doubleBracketPattern.exec(text)) !== null) {
      const raw = m[0]; // e.g. <<First Name>>
      const { matched, key } = matchToken(raw);
      if (matched) {
        if (!matchedKeys.includes(key)) matchedKeys.push(key);
      } else {
        if (!unmatchedKeys.includes(raw)) unmatchedKeys.push(raw);
      }
    }

    // ── 2. @placeholder patterns ──────────────────────────────────────────
    const atKeyPattern = /@[a-zA-Z0-9_]+/g;
    const atKeys = text.match(atKeyPattern) || [];
    atKeys.forEach((raw) => {
      const { matched, key } = matchToken(raw);
      if (matched) {
        if (!matchedKeys.includes(key)) matchedKeys.push(key);
      } else {
        if (!unmatchedKeys.includes(raw)) unmatchedKeys.push(raw);
      }
    });

    // ── 3. Plain UPPERCASE tokens (≥4 chars) ─────────────────────────────
    const uppercasePattern = /\b[A-Z]{4,}[A-Z0-9_]*\b/g;
    const uppercaseKeys = text.match(uppercasePattern) || [];
    uppercaseKeys.forEach((raw) => {
      const { matched, key } = matchToken(raw);
      if (matched) {
        if (!matchedKeys.includes(key)) matchedKeys.push(key);
      } else {
        if (!unmatchedKeys.includes(raw)) unmatchedKeys.push(raw);
      }
    });

    // ── 4. Deduplicate: remove from unmatched if already matched ──────────
    const finalUnmatched = unmatchedKeys.filter(
      (u) => !matchedKeys.some((m) => normalizeToken(m) === normalizeToken(u)),
    );

    return NextResponse.json({
      matchedKeys: [...new Set(matchedKeys)],
      unmatchedKeys: [...new Set(finalUnmatched)],
    });
  } catch (error) {
    console.error("Error analyzing document:", error);
    return NextResponse.json(
      { error: "Failed to analyze document" },
      { status: 500 },
    );
  }
}
