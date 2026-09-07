import { type NextRequest, NextResponse } from "next/server";
import { legacyReplace } from "../../../lib/esign/legacy-adapter";

/* ──────────────────────────────────────────────────────────────────────
 * DocX Key Replacer — replace
 *
 * Same request / response contract as before:
 *   in : multipart form with `file`, `keyMappings` (JSON), and the
 *        `foundKeys` / `unmatchedKeys` lists the UI still sends (the
 *        engine re-analyses the file itself, so they are not needed)
 *   out: the rewritten .docx (or .html/.htm) as an attachment
 *
 * Now backed by the shared esign engine (lib/esign/legacy-adapter.ts).
 * The important difference from the old implementation: text is spliced
 * inside the existing runs, so bold / italic / colour / size of every run
 * in a paragraph survive a replacement. HTML files are written back in
 * their original charset with browser-safe &lt;&lt;KEY&gt;&gt; keys.
 * ────────────────────────────────────────────────────────────────────── */

function parseMappings(raw: FormDataEntryValue | null): Record<string, string> {
  if (typeof raw !== "string" || !raw.trim()) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? (parsed as Record<string, string>) : {};
  } catch {
    return {};
  }
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    const keyMappings = parseMappings(formData.get("keyMappings"));

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    const bytes = new Uint8Array(await file.arrayBuffer());
    const result = legacyReplace(bytes, file.name, keyMappings);

    const body = new Blob([result.bytes as Uint8Array<ArrayBuffer>]);
    return new NextResponse(body, {
      headers: {
        "Content-Type": result.contentType,
        "Content-Disposition": `attachment; filename="${encodeURIComponent(file.name)}"`,
        "X-Replaced-Keys": encodeURIComponent(JSON.stringify(result.replaced)),
        "X-Skipped-Keys": encodeURIComponent(JSON.stringify(result.skipped)),
      },
    });
  } catch (error) {
    console.error("Error processing document:", error);
    return NextResponse.json(
      { error: "Failed to process document" },
      { status: 500 },
    );
  }
}
