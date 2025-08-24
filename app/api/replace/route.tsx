import { type NextRequest, NextResponse } from "next/server";
import PizZip from "pizzip";
import keys from "../../../keys.json";

function escapeRegExp(string: string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File;
    const keyMappingsStr = formData.get("keyMappings") as string;
    const foundKeysStr = formData.get("foundKeys") as string;
    const unmatchedKeysStr = formData.get("unmatchedKeys") as string;

    const keyMappings = keyMappingsStr ? JSON.parse(keyMappingsStr) : {};
    const foundKeys = foundKeysStr ? JSON.parse(foundKeysStr) : [];
    const unmatchedKeys = unmatchedKeysStr ? JSON.parse(unmatchedKeysStr) : [];

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    // Read the file as array buffer
    const arrayBuffer = await file.arrayBuffer();
    const zip = new PizZip(arrayBuffer);

    // Get the main document content
    const documentXml = zip.files["word/document.xml"]?.asText();
    if (!documentXml) {
      return NextResponse.json(
        { error: "Invalid document format" },
        { status: 400 }
      );
    }

    let updatedXml = documentXml;

    // Replace keys in order from longest to shortest to prevent partial replacements
    const allKeysToReplace = [...foundKeys, ...unmatchedKeys]
      .filter((key) => keys[key as keyof typeof keys] || keyMappings[key])
      .sort((a, b) => b.length - a.length); // Longest first

    console.log("[v0] Keys to replace (sorted):", allKeysToReplace);

    allKeysToReplace.forEach((key: string) => {
      let finalValue = "";

      if (keys[key as keyof typeof keys]) {
        finalValue = keys[key as keyof typeof keys];
      } else if (keyMappings[key]) {
        finalValue = keyMappings[key];
      } else {
        return; // Skip if no mapping exists
      }

      const escapedValue = finalValue
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");

      // More precise regex to avoid partial matches
      const regex = new RegExp(
        `(?<![A-Za-z0-9_])${escapeRegExp(key)}(?![A-Za-z0-9_])`,
        "g"
      );

      const matches = updatedXml.match(regex);
      if (matches && matches.length > 0) {
        console.log(
          `[v0] Replacing ${key} with ${escapedValue} (${matches.length} occurrences)`
        );
        updatedXml = updatedXml.replace(regex, escapedValue);
      }
    });

    // Update the document with the modified XML
    zip.file("word/document.xml", updatedXml);

    // Generate the updated document
    const buffer = zip.generate({
      type: "nodebuffer",
      compression: "DEFLATE"
    });

    // Convert Node.js Buffer -> Uint8Array
    const uint8Array = new Uint8Array(buffer);

    return new NextResponse(uint8Array, {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${file.name}"`
      }
    });
  } catch (error) {
    console.error("Error processing document:", error);
    return NextResponse.json(
      { error: "Failed to process document" },
      { status: 500 }
    );
  }
}
