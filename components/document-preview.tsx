"use client";

import { useEffect, useRef } from "react";

interface DocumentPreviewProps {
  file: File;
}

export function DocumentPreview({ file }: DocumentPreviewProps) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const loadDocxPreview = async () => {
      if (!containerRef.current) return;

      try {
        console.log(
          "[v0] Loading document preview for file:",
          file.name,
          "size:",
          file.size
        );

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

        console.log("[v0] Document preview rendered successfully");
      } catch (error) {
        console.error("Error rendering document:", error);
        if (containerRef.current) {
          containerRef.current.innerHTML =
            '<p class="text-red-500">Error loading document preview</p>';
        }
      }
    };

    loadDocxPreview();
  }, [file]);

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

        /* Ensure paragraphs and divs don't overlap */
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
      `}</style>
    </div>
  );
}
