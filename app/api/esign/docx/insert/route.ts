import { type NextRequest, NextResponse } from "next/server";
import {
  insertTypedTags,
  type TagInsert,
} from "../../../../../lib/esign/docx-template";

/* ──────────────────────────────────────────────────────────────────────
 * Esign-Templates / Customized — add a tag yourself
 *
 * Form fields:
 *   file     the .docx
 *   inserts  JSON [{ afterText, type: checkbox|checkboxChecked|textbox, all? }]
 *
 * Writes an unnumbered <c> / <cc> / <t> right after the named text, inside
 * the same run (formatting untouched), and returns the new .docx. The
 * client re-analyses it; the tag is then numbered like any typed tag.
 * ────────────────────────────────────────────────────────────────────── */

const TYPES = new Set(["checkbox", "checkboxChecked", "textbox"]);

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }
    if (!file.name.toLowerCase().endsWith(".docx")) {
      return NextResponse.json(
        { error: "This endpoint only handles .docx files" },
        { status: 400 },
      );
    }

    let inserts: TagInsert[] = [];
    try {
      const raw = formData.get("inserts");
      const parsed = typeof raw === "string" ? JSON.parse(raw) : [];
      inserts = (Array.isArray(parsed) ? parsed : []).filter(
        (x) =>
          x &&
          typeof x.afterText === "string" &&
          x.afterText.trim() &&
          TYPES.has(x.type),
      );
    } catch {
      inserts = [];
    }
    if (!inserts.length) {
      return NextResponse.json({ error: "No valid inserts" }, { status: 400 });
    }

    const { bytes, inserted } = insertTypedTags(await file.arrayBuffer(), inserts);
    const body = new Blob([bytes as Uint8Array<ArrayBuffer>]);
    return new NextResponse(body, {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${encodeURIComponent(file.name)}"`,
        "X-Esign-Inserted": JSON.stringify(inserted),
      },
    });
  } catch (error) {
    console.error("Error inserting tags:", error);
    return NextResponse.json({ error: "Failed to insert tags" }, { status: 500 });
  }
}
