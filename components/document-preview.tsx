"use client";
import React, { useRef, useEffect } from "react";

interface DocumentPreviewProps {
  file: File;
  matchedKeys?: string[];
  unmatchedKeys?: string[];
  isReplaced?: boolean;
}

export function DocumentPreview({
  file,
  matchedKeys = [],
  unmatchedKeys = [],
  isReplaced = false,
}: DocumentPreviewProps) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const loadPreview = async () => {
      if (!containerRef.current) return;
      try {
        const fileName = file.name.toLowerCase();
        const isHtmlFile = fileName.endsWith(".html") || fileName.endsWith(".htm");

        if (isHtmlFile) {
          // Handle HTML/HTM files
          const htmlContent = await file.text();
          containerRef.current.innerHTML = htmlContent;
          highlightKeys(
            containerRef.current,
            matchedKeys,
            unmatchedKeys,
            isReplaced,
          );
        } else {
          // Handle DOCX files
          const { renderAsync } = await import("docx-preview");
          containerRef.current.innerHTML = "";
          await renderAsync(file, containerRef.current, undefined, {
            className: "docx-wrapper",
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            ignoreFonts: false,
            breakPages: true,
            ignoreLastRenderedPageBreak: true,
          });

          highlightKeys(
            containerRef.current,
            matchedKeys,
            unmatchedKeys,
            isReplaced,
          );
        }
      } catch (error) {
        console.error("Error rendering document:", error);
        if (containerRef.current) {
          containerRef.current.innerHTML =
            '<p class="text-red-500">Error loading document preview</p>';
        }
      }
    };
    loadPreview();
  }, [file, matchedKeys, unmatchedKeys, isReplaced]);

  return (
    <div className="h-full overflow-auto">
      <div
        ref={containerRef}
        className="docx-preview-container"
        style={{
          minHeight: "100%",
          padding: "1rem",
          backgroundColor: "white",
          color: "black",
        }}
      />
      <style jsx global>{`
        .docx-wrapper {
          background: white;
          padding: 20px;
          margin: 0 auto;
          box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
          line-height: 1.5 !important;
          word-break: break-word;
          white-space: normal !important;
          position: static !important;
        }
        .docx-wrapper p,
        .docx-wrapper div {
          margin: 0 0 0.5em 0;
          line-height: 1.5 !important;
          white-space: normal !important;
          position: static !important;
        }
        .docx-wrapper table {
          border-collapse: collapse;
          width: 100%;
        }
        .docx-wrapper table td,
        .docx-wrapper table th {
          border: 1px solid #ddd;
          padding: 8px;
          white-space: normal !important;
        }
        .key-highlight-yellow {
          background-color: #fef08a;
          color: #713f12;
          border-radius: 3px;
          padding: 0 2px;
          font-weight: 600;
          border-bottom: 2px solid #eab308;
        }
        .key-highlight-green {
          background-color: #bbf7d0;
          color: #14532d;
          border-radius: 3px;
          padding: 0 2px;
          font-weight: 600;
          border-bottom: 2px solid #22c55e;
        }
      `}</style>
    </div>
  );
}

/* ── helpers ── */

function escapeRegExp(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function escapeHtml(s: string) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildPatterns(key: string): { re: RegExp; cls: string }[] {
  const bare = key
    .replace(/^<</, "")
    .replace(/>>$/, "")
    .replace(/^@/, "")
    .trim();
  const stripped = bare.replace(/\s+/g, "");
  const results: { re: RegExp; cls: string }[] = [];

  // 1. <<…>> wrapper — matches literal << >> AND XML-encoded &lt;&lt; &gt;&gt;
  if (stripped.length > 0) {
    try {
      const inner = stripped.split("").map(escapeRegExp).join("[\\s]*");
      const patternStr = `(?:<<|&lt;&lt;)[\\s]*${inner}[\\s]*(?:>>|&gt;&gt;)`;
      results.push({
        re: new RegExp(patternStr, "gi"),
        cls: "",
      });
    } catch (e) {
      console.error("   Pattern 1 error:", e);
    }
  }

  // 2. Strict ALLCAPS word (≥4 chars, no lowercase at all)
  if (/^[A-Z][A-Z0-9_]{3,}$/.test(stripped)) {
    try {
      const patternStr = `(?<![A-Za-z0-9_])${escapeRegExp(stripped)}(?![A-Za-z0-9_])`;
      results.push({
        re: new RegExp(patternStr, "g"),
        cls: "",
      });
    } catch (e) {
      console.error("   Pattern 2 error:", e);
    }
  }

  return results;
}

function highlightKeys(
  root: HTMLElement,
  matchedKeys: string[],
  unmatchedKeys: string[],
  isReplaced: boolean,
) {
  const rules: { pattern: RegExp; cssClass: string }[] = [];

  // Matched (known placeholders) → GREEN
  matchedKeys.forEach((key) =>
    buildPatterns(key).forEach(({ re }) =>
      rules.push({ pattern: re, cssClass: "key-highlight-green" }),
    ),
  );

  // Unmatched (unknown placeholders) → YELLOW
  unmatchedKeys.forEach((key) =>
    buildPatterns(key).forEach(({ re }) =>
      rules.push({ pattern: re, cssClass: "key-highlight-yellow" }),
    ),
  );

  if (rules.length) {
    walkTextNodes(root, rules);
  }

  // ── FALLBACK: catch any <<...>> placeholders the detector missed ──
  if (!isReplaced) {
    // Before replace: catch lowercase/unknown placeholders missed by scanner → YELLOW
    const fallbackRules = [
      {
        pattern: /&lt;&lt;[a-z][a-z0-9_\s]*&gt;&gt;/gi,
        cssClass: "key-highlight-yellow",
      },
      { pattern: /<<[a-z][a-z0-9_\s]*>>/gi, cssClass: "key-highlight-yellow" },
    ];
    walkTextNodes(root, fallbackRules);
  } else {
    // After replace: catch ANY remaining placeholder → YELLOW (unknown)
    const fallbackRules = [
      {
        pattern: /&lt;&lt;[a-z0-9_][a-z0-9_\s]*&gt;&gt;/gi,
        cssClass: "key-highlight-yellow",
      },
      {
        pattern: /<<[a-z0-9_][a-z0-9_\s]*>>/gi,
        cssClass: "key-highlight-yellow",
      },
    ];
    walkTextNodes(root, fallbackRules);
  }
}

function walkTextNodes(
  node: Node,
  rules: { pattern: RegExp; cssClass: string }[],
) {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent || "";
    if (!text.trim()) return;

    let html = escapeHtml(text);
    let modified = false;

    // Log only if text contains something interesting (angle brackets or known words)
    if (
      text.includes("<") ||
      text.includes("address") ||
      text.includes("CAMPUS")
    ) {
    }

    for (const { pattern, cssClass } of rules) {
      pattern.lastIndex = 0;
      const hasMatch = pattern.test(html);
      pattern.lastIndex = 0; // reset after test

      if (
        hasMatch &&
        (text.includes("<") ||
          text.includes("address") ||
          text.includes("CAMPUS"))
      ) {
      }

      const next = html.replace(pattern, (m) => {
        if (
          text.includes("<") ||
          text.includes("address") ||
          text.includes("CAMPUS")
        ) {
        }
        modified = true;
        return `<mark class="${cssClass}">${m}</mark>`;
      });

      if (next !== html) {
        html = next;
        modified = true;
      }
    }

    if (modified) {
      const span = document.createElement("span");
      span.innerHTML = html;
      node.parentNode?.replaceChild(span, node);
      if (
        text.includes("<") ||
        text.includes("address") ||
        text.includes("CAMPUS")
      ) {
      }
    }
    return;
  }

  if (
    node.nodeName === "SCRIPT" ||
    node.nodeName === "STYLE" ||
    node.nodeName === "MARK"
  )
    return;

  Array.from(node.childNodes).forEach((c) => walkTextNodes(c, rules));
}
