import { type NextRequest, NextResponse } from "next/server";
import {
  buildEsignDocx,
  type DocxAssignment,
} from "../../../../../lib/esign/docx-template";

/* ──────────────────────────────────────────────────────────────────────
 * Esign-Templates / Customized — build the mapped .docx
 *
 * Form fields:
 *   file         the client's .docx
 *   assignments  JSON { "<candidate id>": { type, id } }
 *                type: checkbox | checkboxChecked | textbox | ignore
 *                id:   "c1" / "t2" …
 *   mappings     JSON { "<<raw token>>" | "free text": "<<KEY>>" | "text" }
 *   options      JSON { replaceKeys }
 *
 * Returns the .docx with every assigned candidate turned into its
 * <c1> / <c1c> / <t1> marker text and keys rewritten to <<KEY>>.
 * The HTML conversion happens in the browser afterwards.
 * ────────────────────────────────────────────────────────────────────── */

function parseJson<T>(raw: FormDataEntryValue | null, fallback: T): T {
  if (typeof raw !== "string" || !raw.trim()) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

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

    const assignments = parseJson<Record<string, DocxAssignment>>(
      formData.get("assignments"),
      {},
    );
    const mappings = parseJson<Record<string, string>>(
      formData.get("mappings"),
      {},
    );
    const options = parseJson<{ replaceKeys?: boolean; exclude?: string[] }>(
      formData.get("options"),
      {},
    );

    const { bytes, stats } = buildEsignDocx(await file.arrayBuffer(), {
      assignments,
      mappings,
      exclude: Array.isArray(options.exclude) ? options.exclude : [],
      replaceKeys: options.replaceKeys ?? true,
    });

    const body = new Blob([bytes as Uint8Array<ArrayBuffer>]);
    // Keep the header small: per-candidate marker map is not needed by the
    // client (it already knows the assignments); a count is enough.
    const headerStats = {
      keys: stats.keys,
      skipped: stats.skipped,
      markerCount: Object.keys(stats.markers).length,
    };
    return new NextResponse(body, {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${encodeURIComponent(file.name)}"`,
        "X-Esign-Stats": encodeURIComponent(JSON.stringify(headerStats)),
      },
    });
  } catch (error) {
    console.error("Error building esign docx:", error);
    return NextResponse.json(
      { error: "Failed to build .docx template" },
      { status: 500 },
    );
  }
}
