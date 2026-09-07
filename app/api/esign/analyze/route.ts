import { type NextRequest, NextResponse } from "next/server";
import {
  analyzeEsignHtml,
  decodeHtmlBytes,
} from "../../../../lib/esign/html-template";

/* ──────────────────────────────────────────────────────────────────────
 * Esign-Templates / Customized — analyze
 *
 * Accepts one .html/.htm file (converted from Word) and reports the
 * placeholder keys and checkbox/textbox markers it contains.
 * Independent of /api/analyze (the DOCX pipeline).
 * ────────────────────────────────────────────────────────────────────── */

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    const name = file.name.toLowerCase();
    if (!name.endsWith(".html") && !name.endsWith(".htm")) {
      return NextResponse.json(
        { error: "Esign templates must be .html or .htm files" },
        { status: 400 },
      );
    }

    const { html, charset } = decodeHtmlBytes(await file.arrayBuffer());
    const analysis = analyzeEsignHtml(html);

    return NextResponse.json({ ...analysis, charset });
  } catch (error) {
    console.error("Error analyzing esign template:", error);
    return NextResponse.json(
      { error: "Failed to analyze esign template" },
      { status: 500 },
    );
  }
}
