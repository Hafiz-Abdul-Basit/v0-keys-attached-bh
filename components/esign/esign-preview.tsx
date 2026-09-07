"use client";
import { useEffect, useState } from "react";
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
  body { margin: 16px; }
  img { max-width: 100%; }
`;

/**
 * Renders the template inside a sandboxed iframe so the Word-generated
 * <style> blocks cannot leak into the app, and real <input> controls show
 * up exactly as a browser would display them.
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

  return (
    <iframe
      title="Esign template preview"
      srcDoc={doc}
      sandbox="allow-same-origin"
      className="w-full h-full border-0 bg-white rounded-md"
    />
  );
}
