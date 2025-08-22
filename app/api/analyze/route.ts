import { type NextRequest, NextResponse } from "next/server"
import PizZip from "pizzip"
import Docxtemplater from "docxtemplater"
import keys from "../../../keys.json"

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData()
    const file = formData.get("file") as File

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 })
    }

    // Read the file as array buffer
    const arrayBuffer = await file.arrayBuffer()
    const zip = new PizZip(arrayBuffer)

    // Create docxtemplater instance
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    })

    // Render the document
    doc.render()

    // Get the document text
    const text = doc.getFullText()

    const matchedKeys: string[] = []
    const unmatchedKeys: string[] = []

    Object.keys(keys).forEach((key) => {
      if (text.includes(key)) {
        matchedKeys.push(key)
      }
    })

    const uppercaseWords = text.match(/\b[A-Z][A-Z0-9_]{1,}\b/g) || []
    const uniqueUppercaseWords = [...new Set(uppercaseWords)]

    const commonWords = [
      "THE",
      "AND",
      "FOR",
      "ARE",
      "BUT",
      "NOT",
      "YOU",
      "ALL",
      "CAN",
      "HER",
      "WAS",
      "ONE",
      "OUR",
      "HAD",
      "BUT",
      "WORDS",
      "USE",
      "EACH",
      "WHICH",
      "SHE",
      "HOW",
      "ITS",
      "OIL",
      "SIT",
      "NOW",
      "FIND",
      "LONG",
      "DOWN",
      "DAY",
      "DID",
      "GET",
      "HAS",
      "HIM",
      "HIS",
      "HOW",
      "ITS",
      "MAY",
      "NEW",
      "NOW",
      "OLD",
      "SEE",
      "TWO",
      "WHO",
      "BOY",
      "DID",
      "ITS",
      "LET",
      "OLD",
      "PUT",
      "SAY",
      "SHE",
      "TOO",
      "USE",
    ]

    uniqueUppercaseWords.forEach((word) => {
      if (
        !Object.keys(keys).includes(word) &&
        !matchedKeys.includes(word) &&
        !commonWords.includes(word) &&
        word.length > 2
      ) {
        // Only include words longer than 2 characters
        unmatchedKeys.push(word)
      }
    })

    console.log("[v0] Found matching keys:", matchedKeys)
    console.log("[v0] Found unmatched keys:", unmatchedKeys)

    return NextResponse.json({ matchedKeys, unmatchedKeys })
  } catch (error) {
    console.error("Error analyzing document:", error)
    return NextResponse.json({ error: "Failed to analyze document" }, { status: 500 })
  }
}
