"use client";
import { useCallback, useEffect, useRef, useState } from "react";
import {
  analyzeEsignHtml,
  highlightEsignHtml,
  ESIGN_HIGHLIGHT_CSS,
} from "@/lib/esign/html-template";

interface EsignPreviewProps {
  /** Decoded HTML of the template (original or replaced) */
  html: string;
  /** Current key mappings — used to colour "mapped" tokens blue */
  mappings?: Record<string, string>;
  /** Tokens the user excluded from replacement — shown struck through */
  exclude?: string[];
  /** Wrap detected keys / markers in coloured marks */
  highlight?: boolean;
}

const BASE_CSS = `
  html, body { background: #fff; color: #000; }
  body { margin: 12px; }
  img { max-width: 100%; }
  /* docx-preview draws real 8.5×11in pages; keep that layout (columns,
     indents, floating boxes stay where Word put them) and let the parent
     zoom the whole page to the pane width instead of reflowing it */
  .docx-wrapper { background: #fff !important; padding: 0 !important; }
  .docx-wrapper > section.docx {
    margin: 0 auto 14px auto !important; box-shadow: none !important;
    min-height: 0 !important; border-bottom: 1px dashed #d1d5db;
  }
`;

/** width of the docx-preview page in CSS px, from its inline style */
function pageWidthPx(page: HTMLElement): number {
  const m = page.style.width.match(/^([\d.]+)(pt|px|in|cm|mm)$/);
  if (!m) return page.offsetWidth || 816;
  const v = parseFloat(m[1]);
  switch (m[2]) {
    case "pt":
      return (v * 96) / 72;
    case "in":
      return v * 96;
    case "cm":
      return (v * 96) / 2.54;
    case "mm":
      return (v * 96) / 25.4;
    default:
      return v;
  }
}

/**
 * Renders the template inside a sandboxed iframe so the Word-generated
 * <style> blocks cannot leak into the app, and real <input> controls show
 * up exactly as a browser would display them. A Word page rendered by
 * docx-preview is zoomed down to fit the pane width.
 */
export function EsignPreview({
  html,
  mappings = {},
  exclude = [],
  highlight = true,
}: EsignPreviewProps) {
  const mappingsKey = JSON.stringify(mappings);
  const excludeKey = JSON.stringify(exclude);
  const [doc, setDoc] = useState("");
  const iframeRef = useRef<HTMLIFrameElement>(null);

  /** zoom the docx-preview page so it fits the iframe width (max 1:1) */
  const fit = useCallback(() => {
    const frame = iframeRef.current;
    const cdoc = frame?.contentDocument;
    const body = cdoc?.body as (HTMLElement & { style: CSSStyleDeclaration & { zoom?: string } }) | undefined;
    if (!frame || !cdoc || !body) return;
    const page = cdoc.querySelector("section.docx") as HTMLElement | null;
    if (!page) {
      body.style.zoom = "";
      return;
    }
    const avail = frame.clientWidth - 44; // body margins + a scrollbar
    const zoom = Math.min(1, avail / pageWidthPx(page));
    body.style.zoom = zoom > 0 && Number.isFinite(zoom) ? zoom.toFixed(3) : "";
  }, []);

  useEffect(() => {
    // small debounce: mappings change on every keystroke
    const handle = setTimeout(() => {
      let body = html;
      if (highlight) {
        try {
          const analysis = analyzeEsignHtml(html);
          body = highlightEsignHtml(
            html,
            analysis,
            JSON.parse(mappingsKey),
            JSON.parse(excludeKey),
          );
        } catch (err) {
          console.error("Esign preview highlight failed:", err);
        }
      }
      const style = `<style>${BASE_CSS}${ESIGN_HIGHLIGHT_CSS}</style>`;
      if (/<head\b[^>]*>/i.test(body)) {
        body = body.replace(/<head\b[^>]*>/i, (m) => m + style);
      } else {
        body = style + body;
      }
      setDoc(body);
    }, 250);
    return () => clearTimeout(handle);
  }, [html, mappingsKey, excludeKey, highlight]);

  useEffect(() => {
    const frame = iframeRef.current;
    if (!frame || typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(() => fit());
    ro.observe(frame);
    return () => ro.disconnect();
  }, [fit]);

  return (
    <iframe
      ref={iframeRef}
      title="Esign template preview"
      srcDoc={doc}
      sandbox="allow-same-origin"
      onLoad={fit}
      className="w-full h-full border-0 bg-white rounded-md"
    />
  );
}
