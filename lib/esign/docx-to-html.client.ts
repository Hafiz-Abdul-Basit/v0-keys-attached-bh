"use client";
/* ──────────────────────────────────────────────────────────────────────
 * Esign-Templates — DOCX → standalone HTML (browser only)
 *
 * Uses docx-preview (already a dependency, high-fidelity rendering of
 * fonts, tables, page layout) to render the .docx into a detached
 * element, then serialises it into a single self-contained .htm:
 * document CSS inlined in <style>, pictures inlined as data: URIs.
 * ────────────────────────────────────────────────────────────────────── */

function escapeHtmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

const OVERRIDE_CSS = `
/* esign: plain white page, no viewer chrome */
.docx-wrapper { background: #ffffff !important; padding: 0 !important; }
.docx-wrapper > section.docx { box-shadow: none !important; margin-bottom: 0 !important; }
`;

export async function convertDocxToHtml(
  data: Blob | ArrayBuffer | Uint8Array,
  title: string,
): Promise<string> {
  const { renderAsync } = await import("docx-preview");

  const host = document.createElement("div");
  host.style.cssText =
    "position:absolute;left:-100000px;top:0;width:1200px;visibility:hidden;";
  const body = document.createElement("div");
  const styles = document.createElement("div");
  host.appendChild(styles);
  host.appendChild(body);
  document.body.appendChild(host);

  try {
    const blob =
      data instanceof Blob
        ? data
        : new Blob([data as Uint8Array<ArrayBuffer>]);
    await renderAsync(blob, body, styles, {
      className: "docx",
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: true,
      useBase64URL: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      renderChanges: false,
      renderComments: false,
      trimXmlDeclaration: true,
    });

    const css = Array.from(styles.querySelectorAll("style"))
      .map((s) => s.textContent ?? "")
      .join("\n");

    return [
      "<!DOCTYPE html>",
      '<html lang="en">',
      "<head>",
      '<meta charset="utf-8">',
      '<meta name="viewport" content="width=device-width, initial-scale=1">',
      `<title>${escapeHtmlText(title)}</title>`,
      "<style>",
      css,
      OVERRIDE_CSS,
      "</style>",
      "</head>",
      "<body>",
      body.innerHTML,
      "</body>",
      "</html>",
    ].join("\n");
  } finally {
    host.remove();
  }
}
