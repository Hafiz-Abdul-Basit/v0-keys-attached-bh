import { type NextRequest, NextResponse } from "next/server";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import keys from "../../../keys.json";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File;

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    // Read the file as array buffer
    const arrayBuffer = await file.arrayBuffer();
    const zip = new PizZip(arrayBuffer);

    // Extract text from all XML parts of the document
    let allText = "";

    // Try to extract with docxtemplater first
    try {
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true
      });
      doc.render();
      allText += doc.getFullText() + " ";
    } catch (error) {
      console.log(
        "[v0] Error with docxtemplater, trying direct XML extraction"
      );
    }

    // Also extract directly from document.xml to catch table content
    try {
      const documentXml = zip.file("word/document.xml")?.asText();
      if (documentXml) {
        const textContent = documentXml
          .replace(/<[^>]*>/g, " ") // strip XML tags
          .replace(/\s+/g, " ") // normalize whitespace
          .trim();
        allText += " " + textContent;
      }
    } catch (error) {
      console.log("[v0] Could not extract from document.xml");
    }

    const text = allText.trim();
    console.log("[v0] Full extracted text:", text);

    const matchedKeys: string[] = [];
    const unmatchedKeys: string[] = [];

    // First, find all known keys that exist in the document
    Object.keys(keys).forEach((key) => {
      if (text.includes(key)) {
        matchedKeys.push(key);
      }
    });

    // Detect ONLY uppercase keys (4+ characters)
    const uppercasePattern = /\b[A-Z]{4,}[A-Z0-9]*\b/g;
    const allUppercaseKeys = text.match(uppercasePattern) || [];

    // Remove duplicates and sort by length (longest first)
    const uniqueUppercaseKeys = [...new Set(allUppercaseKeys)].sort(
      (a, b) => b.length - a.length
    );

    console.log("[v0] All uppercase keys found:", uniqueUppercaseKeys);

    // Now check each uppercase key against our known keys
    uniqueUppercaseKeys.forEach((key) => {
      // Skip if already matched
      if (matchedKeys.includes(key)) return;

      // Skip if it's a subset of an already matched key
      // BUT only if it's not a real key in keys.json
      if (
        matchedKeys.some(
          (matchedKey) =>
            matchedKey.includes(key) &&
            matchedKey !== key &&
            !uniqueUppercaseKeys.includes(key)
        )
      ) {
        return;
      }

      // Skip if it's just a number or trivial acronym
      if (
        key.match(/^\d+$/) ||
        key.match(/^[A-Z]{1,3}$/) ||
        ["THE", "AND", "FOR", "WITH", "THIS", "THAT", "HAVE", "FROM"].includes(
          key
        )
      ) {
        return;
      }

      // If not a known key, add to unmatched
      if (!keys[key as keyof typeof keys]) {
        unmatchedKeys.push(key);
      }
    });

    // Filter out any keys that are partial matches of longer unmatched keys
    const finalUnmatchedKeys = unmatchedKeys.filter((key) => {
      return !unmatchedKeys.some(
        (otherKey) =>
          otherKey !== key &&
          otherKey.length > key.length &&
          otherKey.includes(key)
      );
    });

    console.log("[v0] Found matching keys:", matchedKeys);
    console.log("[v0] Found unmatched keys:", finalUnmatchedKeys);

    return NextResponse.json({
      matchedKeys,
      unmatchedKeys: finalUnmatchedKeys
    });
  } catch (error) {
    console.error("Error analyzing document:", error);
    return NextResponse.json(
      { error: "Failed to analyze document" },
      { status: 500 }
    );
  }
}
