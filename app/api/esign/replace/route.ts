import { type NextRequest, NextResponse } from "next/server";
import {
  analyzeEsignHtml,
  decodeHtmlBytes,
  encodeHtmlForCharset,
  transformEsignHtml,
  type EsignKeyOutput,
  type EsignTemplates,
} from "../../../../lib/esign/html-template";

/* ──────────────────────────────────────────────────────────────────────
 * Esign-Templates / Customized — replace
 *
 * Form fields:
 *   file      .html / .htm converted from Word
 *   mappings  JSON { "<<raw token>>" | "free text": "<<KEY>>" | "text" }
 *   options   JSON { keyOutput, replaceKeys, replaceControls, templates }
 *
 * Returns the rewritten HTML (UTF-8). Replacement counts are exposed in
 * the X-Esign-Stats header. Independent of /api/replace.
 * ────────────────────────────────────────────────────────────────────── */

interface ReplaceOptionsBody {
  keyOutput?: EsignKeyOutput;
  replaceKeys?: boolean;
  replaceControls?: boolean;
  templates?: Partial<EsignTemplates>;
  exclude?: string[];
}

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

    const name = file.name.toLowerCase();
    if (!name.endsWith(".html") && !name.endsWith(".htm")) {
      return NextResponse.json(
        { error: "Esign templates must be .html or .htm files" },
        { status: 400 },
      );
    }

    const mappings = parseJson<Record<string, string>>(
      formData.get("mappings"),
      {},
    );
    const options = parseJson<ReplaceOptionsBody>(formData.get("options"), {});

    const { html, charset } = decodeHtmlBytes(await file.arrayBuffer());
    const analysis = analyzeEsignHtml(html);
    const { html: out, stats } = transformEsignHtml(html, analysis, {
      mappings,
      exclude: Array.isArray(options.exclude) ? options.exclude : [],
      keyOutput: options.keyOutput === "literal" ? "literal" : "encoded",
      replaceKeys: options.replaceKeys ?? true,
      replaceControls: options.replaceControls ?? true,
      templates: options.templates,
    });

    // Write the file back in the charset it arrived with (Word exports
    // windows-1252) so the download differs from the upload only where a
    // token was replaced.
    const encoded = encodeHtmlForCharset(out, charset);

    const body = new Blob([encoded.bytes as Uint8Array<ArrayBuffer>]);

    return new NextResponse(body, {
      headers: {
        "Content-Type": `text/html; charset=${encoded.charset}`,
        "Content-Disposition": `attachment; filename="${encodeURIComponent(file.name)}"`,
        "X-Esign-Stats": encodeURIComponent(JSON.stringify(stats)),
        "X-Esign-Charset": encoded.charset,
      },
    });
  } catch (error) {
    console.error("Error processing esign template:", error);
    return NextResponse.json(
      { error: "Failed to process esign template" },
      { status: 500 },
    );
  }
}
