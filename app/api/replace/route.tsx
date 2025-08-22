import { type NextRequest, NextResponse } from "next/server"
import PizZip from "pizzip"
import keys from "../../../keys.json"

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData()
    const file = formData.get("file") as File
    const keyMappingsStr = formData.get("keyMappings") as string
    const foundKeysStr = formData.get("foundKeys") as string
    const unmatchedKeysStr = formData.get("unmatchedKeys") as string

    const keyMappings = keyMappingsStr ? JSON.parse(keyMappingsStr) : {}
    const foundKeys = foundKeysStr ? JSON.parse(foundKeysStr) : []
    const unmatchedKeys = unmatchedKeysStr ? JSON.parse(unmatchedKeysStr) : []

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 })
    }

    // Read the file as array buffer
    const arrayBuffer = await file.arrayBuffer()
    const zip = new PizZip(arrayBuffer)

    // Get the main document content
    const documentXml = zip.files["word/document.xml"]?.asText()
    if (!documentXml) {
      return NextResponse.json({ error: "Invalid document format" }, { status: 400 })
    }

    let updatedXml = documentXml

    foundKeys.forEach((key: string) => {
      if (keys[key as keyof typeof keys]) {
        const finalValue = keys[key as keyof typeof keys]
        // Escape XML special characters in the replacement value
        const escapedValue = finalValue
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&apos;")

        // Replace all occurrences of the plain text key with the XML-escaped replacement value
        const regex = new RegExp(`\\b${key}\\b`, "g")
        updatedXml = updatedXml.replace(regex, escapedValue)
        console.log(`[v0] Replaced ${key} with ${escapedValue}`)
      }
    })

    unmatchedKeys.forEach((key: string) => {
      if (keyMappings[key]) {
        const mappedValue = keyMappings[key]
        // Escape XML special characters in the replacement value
        const escapedValue = mappedValue
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&apos;")

        // Replace all occurrences of the plain text key with the XML-escaped replacement value
        const regex = new RegExp(`\\b${key}\\b`, "g")
        updatedXml = updatedXml.replace(regex, escapedValue)
        console.log(`[v0] Mapped ${key} to ${escapedValue}`)
      }
    })

    // Update the document with the modified XML
    zip.file("word/document.xml", updatedXml)

    console.log("[v0] Document processed successfully with XML-safe replacement")

    // Generate the updated document
    const buffer = zip.generate({
      type: "nodebuffer",
      compression: "DEFLATE",
    })

    console.log("[v0] Document buffer generated, size:", buffer.length)

    // Return the file
    return new NextResponse(buffer, {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${file.name}"`,
      },
    })
  } catch (error) {
    console.error("Error processing document:", error)
    return NextResponse.json({ error: "Failed to process document" }, { status: 500 })
  }
}
