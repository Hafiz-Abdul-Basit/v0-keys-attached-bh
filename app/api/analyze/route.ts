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
      f.match(/word\/(document|header\d*|footer\d*)\.xml/)
    );

    xmlFiles.forEach((xmlPath) => {
      const xml = zip.file(xmlPath)?.asText();
      if (xml) {
        const textContent = xml
          .replace(/<[^>]*>/g, " ")
          .replace(/\s+/g, " ")
          .trim();
        allText += " " + textContent;
      }
    });

    const text = allText.trim();
    console.log("[v0] Extracted text (body + headers/footers):", text);

    const matchedKeys: string[] = [];
    const unmatchedKeys: string[] = [];

    // 1. Detect @placeholders
    const atKeyPattern = /@[a-zA-Z0-9_]+/g;
    const atKeys = text.match(atKeyPattern) || [];

    atKeys.forEach((key) => {
      const lowerKey = key.toLowerCase();
      const matched = Object.keys(keys).find(
        (knownKey) => knownKey.toLowerCase() === lowerKey
      );

      if (matched) {
        matchedKeys.push(matched);
      } else {
        unmatchedKeys.push(key);
      }
    });

    // 2. Detect plain UPPERCASE KEYS (≥4 chars, numbers allowed)
    const uppercasePattern = /\b[A-Z]{4,}[A-Z0-9_]*\b/g;
    const uppercaseKeys = text.match(uppercasePattern) || [];

    uppercaseKeys.forEach((key) => {
      const matched = Object.keys(keys).find(
        (knownKey) => knownKey.toLowerCase() === key.toLowerCase()
      );

      if (matched) {
        matchedKeys.push(matched);
      } else {
        unmatchedKeys.push(key);
      }
    });

    return NextResponse.json({
      matchedKeys: [...new Set(matchedKeys)],
      unmatchedKeys: [...new Set(unmatchedKeys)],
    });
  } catch (error) {
    console.error("Error analyzing document:", error);
    return NextResponse.json(
      { error: "Failed to analyze document" },
      { status: 500 }
    );
  }
}
