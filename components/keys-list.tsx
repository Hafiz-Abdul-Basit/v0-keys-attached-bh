"use client"

import { useEffect, useState } from "react"
import { Card, CardContent } from "@/components/ui/card"
import { Badge } from "@/components/ui/badge"
import { Input } from "@/components/ui/input"
import { Search } from "lucide-react"

interface KeysData {
  [key: string]: string
}

interface KeysListProps {
  matchedKeys?: string[]
  unmatchedKeys?: string[]
  onKeyUpdate?: (key: string, value: string) => void
}

export function KeysList({ matchedKeys = [], unmatchedKeys = [], onKeyUpdate }: KeysListProps) {
  const [keys, setKeys] = useState<KeysData>({})
  const [loading, setLoading] = useState(true)
  const [searchTerm, setSearchTerm] = useState("")
  const [keyMappings, setKeyMappings] = useState<{ [key: string]: string }>({})

  useEffect(() => {
    const loadKeys = async () => {
      try {
        const response = await fetch("/keys.json")
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`)
        }
        const data = await response.json()
        setKeys(data)
      } catch (error) {
        console.error("Error loading keys:", error)
      } finally {
        setLoading(false)
      }
    }

    loadKeys()
  }, [])

  const filteredKeys = Object.entries(keys).filter(
    ([key, placeholder]) =>
      key.toLowerCase().includes(searchTerm.toLowerCase()) ||
      placeholder.toLowerCase().includes(searchTerm.toLowerCase()),
  )

  const unmatchedEntries = unmatchedKeys
    .filter((key) => key.toLowerCase().includes(searchTerm.toLowerCase()))
    .map((key) => [key, keyMappings[key] || `<<${key}>>`] as [string, string])

  const allEntries = [...filteredKeys, ...unmatchedEntries]

  const sortedKeys = allEntries.sort(([keyA], [keyB]) => {
    const aMatched = matchedKeys.includes(keyA)
    const bMatched = matchedKeys.includes(keyB)
    const aUnmatched = unmatchedKeys.includes(keyA)
    const bUnmatched = unmatchedKeys.includes(keyB)

    if (aMatched && !bMatched && !bUnmatched) return -1
    if (!aMatched && !aUnmatched && bMatched) return 1
    if (aUnmatched && !bUnmatched && !bMatched) return -1
    if (!aUnmatched && !aMatched && bUnmatched) return 1

    return keyA.localeCompare(keyB)
  })

  const handleKeyMappingChange = (documentKey: string, placeholderKey: string) => {
    setKeyMappings((prev) => ({ ...prev, [documentKey]: placeholderKey }))
    if (onKeyUpdate) {
      onKeyUpdate(documentKey, placeholderKey)
    }
  }

  if (loading) {
    return <div className="text-center text-muted-foreground">Loading keys...</div>
  }

  return (
    <div className="space-y-4 h-full flex flex-col">
      <div className="relative">
        <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-muted-foreground" />
        <Input
          placeholder="Search placeholders..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          className="pl-10"
        />
      </div>

      <div className="flex gap-2 text-sm">
        {matchedKeys.length > 0 && (
          <Badge variant="default" className="bg-green-600">
            {matchedKeys.length} matched
          </Badge>
        )}
        {unmatchedKeys.length > 0 && (
          <Badge variant="secondary" className="bg-orange-100 text-orange-800 dark:bg-orange-900 dark:text-orange-200">
            {unmatchedKeys.length} need mapping
          </Badge>
        )}
      </div>

      <div className="space-y-2 flex-1 overflow-auto">
        {sortedKeys.map(([key, placeholder]) => {
          const isMatched = matchedKeys.includes(key)
          const isUnmatched = unmatchedKeys.includes(key)
          return (
            <Card
              key={key}
              className={`p-3 hover:bg-accent/50 transition-colors ${
                isMatched
                  ? "border-green-500 bg-green-50 dark:bg-green-950"
                  : isUnmatched
                    ? "border-orange-500 bg-orange-50 dark:bg-orange-950"
                    : ""
              }`}
            >
              <CardContent className="p-0">
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <div className="font-medium text-sm">{isUnmatched ? `${key} (in document)` : key}</div>
                    <div className="flex gap-1">
                      {isMatched && (
                        <Badge variant="default" className="text-xs bg-green-600">
                          Found
                        </Badge>
                      )}
                      {isUnmatched && (
                        <Badge variant="secondary" className="text-xs bg-orange-600 text-white">
                          Map
                        </Badge>
                      )}
                    </div>
                  </div>

                  <div className="text-xs font-mono bg-muted p-2 rounded border">{placeholder}</div>

                  {isUnmatched && (
                    <div className="space-y-1">
                      <label className="text-xs text-muted-foreground">Map to placeholder:</label>
                      <Input
                        placeholder={`e.g., <<CORRECT_PLACEHOLDER_NAME>>`}
                        value={keyMappings[key] || ""}
                        onChange={(e) => handleKeyMappingChange(key, e.target.value)}
                        className="text-xs h-8"
                      />
                    </div>
                  )}

                  {(key.toLowerCase().includes("date") || key.toLowerCase().includes("time")) &&
                    onKeyUpdate &&
                    !isUnmatched && (
                      <div className="space-y-1">
                        <label className="text-xs text-muted-foreground">Update date:</label>
                        <Input
                          type="date"
                          className="text-xs h-8"
                          onChange={(e) => onKeyUpdate(key, e.target.value)}
                          placeholder="Update date"
                        />
                      </div>
                    )}
                </div>
              </CardContent>
            </Card>
          )
        })}
      </div>
    </div>
  )
}
