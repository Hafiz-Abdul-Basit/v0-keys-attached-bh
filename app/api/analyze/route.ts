import { type NextRequest, NextResponse } from "next/server";
import { legacyAnalyze } from "../../../lib/esign/legacy-adapter";

/* ──────────────────────────────────────────────────────────────────────
 * DocX Key Replacer — analyze
 *
 * Same request / response contract as before:
 *   in : multipart form with `file` (.docx / .html / .htm)
 *   out: { matchedKeys: string[], unmatchedKeys: string[] }
 *
 * Now backed by the shared esign engine (lib/esign/legacy-adapter.ts):
 * tokens split across Word runs, &lt;&lt;KEY&gt;&gt; in HTML, «KEY»,
 * @KEY and CAPITAL keys are all detected; ordinary capitalised words are
 * no longer reported as "unmatched".
 * ────────────────────────────────────────────────────────────────────── */

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    const bytes = new Uint8Array(await file.arrayBuffer());
    const { matchedKeys, unmatchedKeys } = legacyAnalyze(bytes, file.name);

    return NextResponse.json({ matchedKeys, unmatchedKeys });
  } catch (error) {
    console.error("Error analyzing document:", error);
    return NextResponse.json(
      { error: "Failed to analyze document" },
      { status: 500 },
    );
  }
}
