"use client"

import { useEffect, useRef } from "react"

interface DocumentPreviewProps {
  file: File
}

export function DocumentPreview({ file }: DocumentPreviewProps) {
  const containerRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const loadDocxPreview = async () => {
      if (!containerRef.current) return

      try {
        console.log("[v0] Loading document preview for file:", file.name, "size:", file.size)

        // Dynamic import of docx-preview
        const { renderAsync } = await import("docx-preview")

        // Clear previous content
        containerRef.current.innerHTML = ""

        // Render the document
        await renderAsync(file, containerRef.current, undefined, {
          className: "docx-wrapper",
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          ignoreLastRenderedPageBreak: true,
          experimental: false,
          trimXmlDeclaration: true,
          useBase64URL: false,
          useMathMLPolyfill: false,
          showChanges: false,
          debug: false,
        })

        console.log("[v0] Document preview rendered successfully")
      } catch (error) {
        console.error("Error rendering document:", error)
        if (containerRef.current) {
          containerRef.current.innerHTML = '<p class="text-red-500">Error loading document preview</p>'
        }
      }
    }

    loadDocxPreview()
  }, [file])

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
          box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }
        .docx-wrapper table {
          border-collapse: collapse;
          width: 100%;
        }
        .docx-wrapper table td, .docx-wrapper table th {
          border: 1px solid #ddd;
          padding: 8px;
        }
      `}</style>
    </div>
  )
}
