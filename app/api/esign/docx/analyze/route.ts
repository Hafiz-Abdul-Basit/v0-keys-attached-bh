import { type NextRequest, NextResponse } from "next/server";
import { analyzeEsignDocx } from "../../../../../lib/esign/docx-template";

/* ──────────────────────────────────────────────────────────────────────
 * Esign-Templates / Customized — analyze a client's .docx
 *
 * Reports placeholder keys, markers already typed (<c1>, <t1> …) and
 * every checkbox / textbox *candidate* (content controls, form fields,
 * box symbols, "[ ]", underscores, pictures, shapes) with a suggested
 * type so the user can map them. Independent of /api/analyze.
 * ────────────────────────────────────────────────────────────────────── */

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

    const analysis = analyzeEsignDocx(await file.arrayBuffer());
    return NextResponse.json(analysis);
  } catch (error) {
    console.error("Error analyzing esign docx:", error);
    return NextResponse.json(
      { error: "Failed to analyze .docx template" },
      { status: 500 },
    );
  }
}
