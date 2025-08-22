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

    // Extract text from all XML parts of the document
    let allText = ""

    // Get main document text
    try {
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      })
      doc.render()
      allText += doc.getFullText() + " "
    } catch (error) {
      console.log("[v0] Error with docxtemplater, trying direct XML extraction")
    }

    // Also extract directly from document.xml to catch table content
    try {
      const documentXml = zip.file("word/document.xml")?.asText()
      if (documentXml) {
        // Remove XML tags and extract text content
        const textContent = documentXml
          .replace(/<[^>]*>/g, " ") // Remove XML tags
          .replace(/\s+/g, " ") // Normalize whitespace
          .trim()
        allText += " " + textContent
      }
    } catch (error) {
      console.log("[v0] Could not extract from document.xml")
    }

    const text = allText.trim()

    const matchedKeys: string[] = []
    const unmatchedKeys: string[] = []

    Object.keys(keys).forEach((key) => {
      if (text.includes(key)) {
        matchedKeys.push(key)
      }
    })

    const uppercaseWords = text.match(/\b[A-Z][A-Z0-9_]*[A-Z0-9]\b/g) || []
    const singleUppercaseWords = text.match(/\b[A-Z]{4,}\b/g) || []
    const underscoreWords = text.match(/\b[A-Z]+_[A-Z0-9_]*\b/g) || []
    const mixedCaseWords = text.match(/\b[A-Z]+[A-Z0-9]*[A-Z]+\b/g) || []

    const allUppercaseWords = [...uppercaseWords, ...singleUppercaseWords, ...underscoreWords, ...mixedCaseWords]
    const uniqueUppercaseWords = [...new Set(allUppercaseWords)]

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
      "MAY",
      "NEW",
      "OLD",
      "SEE",
      "TWO",
      "WHO",
      "BOY",
      "LET",
      "PUT",
      "SAY",
      "TOO",
      "UNITED",
      "INDEPENDENT",
      "SCHOOL",
      "DISTRICT",
      "ATTENDANCE",
      "IMPROVEMENT",
      "PLAN",
      "STUDENT",
      "PARENT",
      "GUARDIAN",
      "CAMPUS",
      "DATE",
      "TIME",
      "YEAR",
      "MONTH",
      "WEEK",
      "MINUTES",
      "HOURS",
      "DAYS",
      "ACADEMIC",
      "SUCCESS",
      "CRITICAL",
      "INDICATOR",
      "PASADENA",
      "MISSION",
      "WORK",
      "PARENTS",
      "HELP",
      "ENSURE",
      "ATTEND",
      "REGULARLY",
      "HIGHLY",
      "IMPACTS",
      "SOCIAL",
      "EMOTIONAL",
      "LEARNING",
      "LOST",
      "INSTRUCTIONAL",
      "DUE",
      "UNEXCUSED",
      "ABSENCES",
      "STATED",
      "ABOVE",
      "INCLUDE",
      "EXCUSED",
      "ISS",
      "DPS",
      "ISD",
    ]

    uniqueUppercaseWords.forEach((word) => {
      if (
        !Object.keys(keys).includes(word) &&
        !matchedKeys.includes(word) &&
        !commonWords.includes(word) &&
        word.length > 3 && // Keep minimum length at 4 characters
        !word.match(/^\d+$/) && // Exclude pure numbers
        !word.match(/^[A-Z]{1,2}$/) // Exclude very short abbreviations
      ) {
        unmatchedKeys.push(word)
      }
    })

    console.log("[v0] Document text sample:", text.substring(0, 500))
    console.log("[v0] All uppercase words found:", uniqueUppercaseWords)
    console.log("[v0] Found matching keys:", matchedKeys)
    console.log("[v0] Found unmatched keys:", unmatchedKeys)

    return NextResponse.json({ matchedKeys, unmatchedKeys })
  } catch (error) {
    console.error("Error analyzing document:", error)
    return NextResponse.json({ error: "Failed to analyze document" }, { status: 500 })
  }
}
