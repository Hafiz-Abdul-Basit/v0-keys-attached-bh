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

    const arrayBuffer = await file.arrayBuffer();
    const zip = new PizZip(arrayBuffer);

    // Target body + header + footer XMLs
    const xmlFiles = Object.keys(zip.files).filter((f) =>
      f.match(/word\/(document|header\d*|footer\d*)\.xml/)
    );

    const allKeysToReplace = [...foundKeys, ...unmatchedKeys].sort(
      (a, b) => b.length - a.length
    );

    xmlFiles.forEach((xmlPath) => {
      let updatedXml = zip.files[xmlPath].asText();

      allKeysToReplace.forEach((key: string) => {
        let finalValue = "";

        if (keys[key as keyof typeof keys]) {
          finalValue = keys[key as keyof typeof keys];
        } else {
          const ciMatch = Object.keys(keys).find(
            (knownKey) => knownKey.toLowerCase() === key.toLowerCase()
          );
          if (ciMatch) {
            finalValue = keys[ciMatch as keyof typeof keys];
          } else if (keyMappings[key]) {
            finalValue = keyMappings[key];
          }
        }

        if (!finalValue) return;

        const escapedValue = finalValue
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");

        const regex = new RegExp(
          `(?<![A-Za-z0-9_])${escapeRegExp(key)}(?![A-Za-z0-9_])`,
          "g"
        );

        updatedXml = updatedXml.replace(regex, escapedValue);
      });

      zip.file(xmlPath, updatedXml);
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
  } catch (error) {
    console.error("Error processing document:", error);
    return NextResponse.json(
      { error: "Failed to process document" },
      { status: 500 }
    );
  }
}
