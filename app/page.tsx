"use client"

import type React from "react"
import { useState, useRef, useEffect } from "react"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Upload, Download, X, FileText, Plus, Settings, Map } from "lucide-react"
import { DocumentPreview } from "@/components/document-preview"
import { KeysList } from "@/components/keys-list"
import { downloadFile } from "@/lib/download-helper"
import { Badge } from "@/components/ui/badge"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"

interface FileData {
  file: File
  matchedKeys: string[]
  unmatchedKeys: string[]
  replacedFile?: File
}

export default function Home() {
  const [uploadedFiles, setUploadedFiles] = useState<FileData[]>([])
  const [selectedFileIndex, setSelectedFileIndex] = useState<number>(0)
  const [isProcessing, setIsProcessing] = useState(false)
  const [isUploading, setIsUploading] = useState(false)
  const [keyMappings, setKeyMappings] = useState<{ [key: string]: string }>({})
  const [previewKey, setPreviewKey] = useState(0)
  const [isDragOver, setIsDragOver] = useState(false)
  const [showPlaceholderManager, setShowPlaceholderManager] = useState(false)
  const [newPlaceholderKey, setNewPlaceholderKey] = useState("")
  const [newPlaceholderValue, setNewPlaceholderValue] = useState("")
  const [customPlaceholders, setCustomPlaceholders] = useState<{ [key: string]: string }>({})
  const fileInputRef = useRef<HTMLInputElement>(null)

  // Auto-open placeholder manager when unmatched keys are found
  useEffect(() => {
    if (uniqueUnmatchedKeys.length > 0 && !showPlaceholderManager) {
      setShowPlaceholderManager(true)
      
      // Auto-populate with suggested mappings for unmatched keys
      const newMappings = { ...customPlaceholders }
      uniqueUnmatchedKeys.forEach((key) => {
        if (!newMappings[key]) {
          // Auto-suggest a placeholder format for unmatched keys only
          newMappings[key] = `<<${key}>>`
        }
      })
      setCustomPlaceholders(newMappings)
    }
  }, [uploadedFiles.length])

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragOver(true)
  }

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragOver(false)
  }

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragOver(false)

    const files = Array.from(e.dataTransfer.files)
    const docxFiles = files.filter(
      (file) => file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    if (docxFiles.length === 0) {
      alert("Please upload .docx files only")
      return
    }

    processFiles(docxFiles)
  }

  const processFiles = async (files: File[]) => {
    setIsUploading(true)
    const newFileData: FileData[] = []

    for (const file of files) {
      try {
        const formData = new FormData()
        formData.append("file", file)

        const response = await fetch("/api/analyze", {
          method: "POST",
          body: formData,
        })

        if (response.ok) {
          const { matchedKeys: foundKeys, unmatchedKeys: unfoundKeys } = await response.json()
          newFileData.push({
            file,
            matchedKeys: foundKeys,
            unmatchedKeys: unfoundKeys || [],
          })
          console.log(`[v0] Analyzed ${file.name} - Found matching keys:`, foundKeys)
        }
      } catch (error) {
        console.error(`Error analyzing document ${file.name}:`, error)
      }
    }

    setUploadedFiles((prev) => [...prev, ...newFileData])
    if (uploadedFiles.length === 0 && newFileData.length > 0) {
      setSelectedFileIndex(0)
    }
    setIsUploading(false)
  }

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files
    if (!files) return

    const docxFiles = Array.from(files).filter(
      (file) => file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    if (docxFiles.length === 0) {
      alert("Please upload .docx files only")
      return
    }

    await processFiles(docxFiles)
  }

  const removeFile = (index: number) => {
    setUploadedFiles((prev) => prev.filter((_, i) => i !== index))
    if (selectedFileIndex >= uploadedFiles.length - 1) {
      setSelectedFileIndex(Math.max(0, uploadedFiles.length - 2))
    }
  }

  const handleKeyUpdate = (key: string, value: string) => {
    setKeyMappings((prev) => ({
      ...prev,
      [key]: value,
    }))
  }

  const handleAddPlaceholder = () => {
    if (newPlaceholderKey.trim() && newPlaceholderValue.trim()) {
      const formattedKey = newPlaceholderKey.toUpperCase().trim()
      const formattedValue = newPlaceholderValue.startsWith("<<")
        ? newPlaceholderValue
        : `<<${newPlaceholderValue.toUpperCase().trim()}>>`

      setCustomPlaceholders((prev) => ({
        ...prev,
        [formattedKey]: formattedValue,
      }))

      setNewPlaceholderKey("")
      setNewPlaceholderValue("")
    }
  }

  const handleRemovePlaceholder = (key: string) => {
    setCustomPlaceholders((prev) => {
      const updated = { ...prev }
      delete updated[key]
      return updated
    })
  }

  const handleReplaceKeys = async () => {
    if (uploadedFiles.length === 0) return

    setIsProcessing(true)
    try {
      const updatedFiles = [...uploadedFiles]

      for (let i = 0; i < updatedFiles.length; i++) {
        const fileData = updatedFiles[i]
        const formData = new FormData()
        formData.append("file", fileData.file)
        const allKeyMappings = { ...keyMappings, ...customPlaceholders }
        formData.append("keyMappings", JSON.stringify(allKeyMappings))
        formData.append("foundKeys", JSON.stringify(fileData.matchedKeys))
        formData.append("unmatchedKeys", JSON.stringify(fileData.unmatchedKeys))

        const response = await fetch("/api/replace", {
          method: "POST",
          body: formData,
        })

        if (!response.ok) {
          throw new Error(`Failed to process document ${fileData.file.name}`)
        }

        const blob = await response.blob()
        const timestamp = Date.now()
        const replacedFile = new File([blob], fileData.file.name, {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        })

        updatedFiles[i] = { ...fileData, replacedFile }
      }

      setUploadedFiles(updatedFiles)
      setPreviewKey((prev) => prev + 1)
      console.log("[v0] All files processed successfully")
    } catch (error) {
      console.error("Error replacing keys:", error)
      alert("Error replacing keys. Please try again.")
    } finally {
      setIsProcessing(false)
    }
  }

  const handleDownload = async () => {
    const processedFiles = uploadedFiles.filter((f) => f.replacedFile)

    if (processedFiles.length === 0) {
      alert("Please replace keys first before downloading")
      return
    }

    try {
      if (processedFiles.length === 1) {
        const fileData = processedFiles[0]
        const blob = new Blob([fileData.replacedFile!], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        })
        downloadFile(blob, fileData.file.name)
      } else {
        const JSZip = (await import("jszip")).default
        const zip = new JSZip()

        processedFiles.forEach((fileData) => {
          zip.file(fileData.file.name, fileData.replacedFile!)
        })

        const zipBlob = await zip.generateAsync({ type: "blob" })
        downloadFile(zipBlob, "processed_documents.zip")
      }
    } catch (error) {
      console.error("Error downloading documents:", error)
      alert("Error downloading documents. Please try again.")
    }
  }

  const handleShowPlaceholderManager = () => {
    setShowPlaceholderManager(!showPlaceholderManager)
  }

  const currentFile = uploadedFiles[selectedFileIndex]
  const allMatchedKeys = uploadedFiles.flatMap((f) => f.matchedKeys)
  const allUnmatchedKeys = uploadedFiles.flatMap((f) => f.unmatchedKeys)
  const uniqueMatchedKeys = [...new Set(allMatchedKeys)]
  const uniqueUnmatchedKeys = [...new Set(allUnmatchedKeys)]

  return (
    <div className="min-h-screen bg-background">
      {/* Top Bar */}
      <div className="border-b bg-card p-4">
        <div className="flex items-center justify-between">
          <h1 className="text-2xl font-bold">DocX Key Replacer</h1>
          <div className="flex items-center gap-2">
            <Button 
              onClick={handleShowPlaceholderManager} 
              variant={showPlaceholderManager ? "default" : "outline"}
              className="flex items-center gap-2"
              disabled={uploadedFiles.length === 0}
            >
              <Map className="h-4 w-4" />
              {showPlaceholderManager ? "Hide Mappings" : "Manage Mappings"}
              {uniqueUnmatchedKeys.length > 0 && (
                <Badge variant="secondary" className="ml-1">
                  {uniqueUnmatchedKeys.length}
                </Badge>
              )}
            </Button>
            <Button onClick={() => fileInputRef.current?.click()} className="flex items-center gap-2">
              <Upload className="h-4 w-4" />
              Upload .docx Files
            </Button>
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".docx"
            multiple
            onChange={handleFileUpload}
            className="hidden"
          />
        </div>
      </div>

      {/* Placeholder Manager Panel */}
      {showPlaceholderManager && (
        <div className="border-b bg-muted/30 p-4">
          <Card>
            <CardHeader>
              <CardTitle className="text-lg flex items-center gap-2">
                <Map className="h-5 w-5" />
                Placeholder Mapping
                {uniqueUnmatchedKeys.length > 0 && (
                  <Badge variant="destructive" className="ml-2">
                    {uniqueUnmatchedKeys.length} unmatched keys
                  </Badge>
                )}
              </CardTitle>
              <p className="text-sm text-muted-foreground">
                Map keys found in your documents to the correct placeholders
              </p>
            </CardHeader>
            <CardContent>
              <div className="space-y-6">
                {uniqueUnmatchedKeys.length > 0 && (
                  <div className="space-y-4">
                    <div>
                      <h4 className="font-medium text-sm mb-2">
                        Unmatched Keys Found in Documents:
                      </h4>
                      <p className="text-xs text-muted-foreground mb-3">
                        These keys were found in your documents but don't match any known placeholders.
                        Map them to the correct placeholder format.
                      </p>
                    </div>
                    <div className="grid gap-3 max-h-64 overflow-y-auto p-1">
                      {uniqueUnmatchedKeys.map((key) => (
                        <div key={key} className="flex items-center gap-2 bg-background p-3 rounded-md border">
                          <div className="flex-1">
                            <Label htmlFor={`key-${key}`} className="text-xs text-muted-foreground mb-1 block">
                              Key found in document:
                            </Label>
                            <div className="text-sm font-mono p-2 bg-muted rounded">
                              {key}
                            </div>
                          </div>
                          <div className="flex-1">
                            <Label htmlFor={`value-${key}`} className="text-xs text-muted-foreground mb-1 block">
                              Map to placeholder:
                            </Label>
                            <Input
                              id={`value-${key}`}
                              value={customPlaceholders[key] || ""}
                              onChange={(e) => {
                                setCustomPlaceholders((prev) => ({
                                  ...prev,
                                  [key]: e.target.value,
                                }))
                              }}
                              className="font-mono text-sm"
                              placeholder={`<<${key}>>`}
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                <div className="border-t pt-4">
                  <h4 className="font-medium text-sm mb-3">Add Custom Mapping:</h4>
                  <div className="grid grid-cols-5 gap-3">
                    <div className="col-span-2">
                      <Label htmlFor="new-key" className="text-xs text-muted-foreground mb-1 block">
                        Key found in document:
                      </Label>
                      <Input
                        id="new-key"
                        placeholder="e.g., ABSENCESDATESLOSS"
                        value={newPlaceholderKey}
                        onChange={(e) => setNewPlaceholderKey(e.target.value.toUpperCase())}
                      />
                    </div>
                    <div className="col-span-2">
                      <Label htmlFor="new-value" className="text-xs text-muted-foreground mb-1 block">
                        Map to placeholder:
                      </Label>
                      <Input
                        id="new-value"
                        placeholder="e.g., <<ABSENCESDATESLOSS>>"
                        value={newPlaceholderValue}
                        onChange={(e) => setNewPlaceholderValue(e.target.value)}
                      />
                    </div>
                    <div className="flex items-end">
                      <Button onClick={handleAddPlaceholder} className="flex items-center gap-2 w-full">
                        <Plus className="h-4 w-4" />
                        Add
                      </Button>
                    </div>
                  </div>
                </div>

                {Object.keys(customPlaceholders).filter((key) => !uniqueUnmatchedKeys.includes(key)).length > 0 && (
                  <div className="space-y-2 border-t pt-4">
                    <h4 className="font-medium text-sm">
                      Custom Mappings:
                      <Badge variant="outline" className="ml-2 text-xs">
                        {Object.keys(customPlaceholders).filter((key) => !uniqueUnmatchedKeys.includes(key)).length}
                      </Badge>
                    </h4>
                    <div className="grid gap-2 max-h-32 overflow-y-auto">
                      {Object.entries(customPlaceholders)
                        .filter(([key]) => !uniqueUnmatchedKeys.includes(key))
                        .map(([key, value]) => (
                          <div key={key} className="flex items-center gap-2 bg-background p-3 rounded border">
                            <div className="flex-1">
                              <div className="text-sm font-mono">
                                <span className="text-blue-600 font-semibold">{key}</span>
                                <span className="mx-2 text-muted-foreground">→</span>
                                <span>{value}</span>
                              </div>
                            </div>
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() => handleRemovePlaceholder(key)}
                              className="h-8 w-8 p-0 hover:bg-destructive hover:text-destructive-foreground"
                            >
                              <X className="h-4 w-4" />
                            </Button>
                          </div>
                        ))}
                    </div>
                  </div>
                )}
              </div>
            </CardContent>
          </Card>
        </div>
      )}

      {/* File List Bar */}
      {uploadedFiles.length > 0 && (
        <div className="border-b bg-muted/30 p-2">
          <div className="flex items-center gap-2 overflow-x-auto">
            <span className="text-sm font-medium text-muted-foreground whitespace-nowrap">Files:</span>
            {uploadedFiles.map((fileData, index) => (
              <div
                key={index}
                className={`flex items-center gap-2 px-3 py-1 rounded-md border cursor-pointer transition-colors ${
                  selectedFileIndex === index ? "bg-primary text-primary-foreground" : "bg-background hover:bg-muted"
                }`}
                onClick={() => setSelectedFileIndex(index)}
              >
                <FileText className="h-3 w-3" />
                <span className="text-xs truncate max-w-32">{fileData.file.name}</span>
                {fileData.replacedFile && (
                  <Badge variant="secondary" className="text-xs px-1 py-0">
                    ✓
                  </Badge>
                )}
                <Button
                  variant="ghost"
                  size="sm"
                  className="h-4 w-4 p-0 hover:bg-destructive hover:text-destructive-foreground"
                  onClick={(e) => {
                    e.stopPropagation()
                    removeFile(index)
                  }}
                >
                  <X className="h-3 w-3" />
                </Button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Main Content */}
      <div className="flex h-[calc(100vh-180px)]">
        {/* Left Panel - Document Preview */}
        <div className="flex-1 border-r">
          <div className="h-full overflow-auto p-4">
            <Card className="h-full">
              <CardHeader>
                <CardTitle>
                  Document Preview
                  {currentFile && (
                    <span className="text-sm font-normal text-muted-foreground ml-2">({currentFile.file.name})</span>
                  )}
                </CardTitle>
              </CardHeader>
              <CardContent className="h-[calc(100%-80px)]">
                {isUploading ? (
                  <div className="flex h-full items-center justify-center">
                    <div className="text-center space-y-6">
                      <div className="relative">
                        {/* Outer rotating ring */}
                        <div className="w-20 h-20 mx-auto relative">
                          <div className="absolute inset-0 border-4 border-gray-200 rounded-full"></div>
                          <div className="absolute inset-0 border-4 border-blue-500 border-t-transparent rounded-full animate-spin"></div>

                          {/* Inner pulsing document icon */}
                          <div className="absolute inset-0 flex items-center justify-center">
                            <div className="bg-white rounded-lg p-2 shadow-sm animate-pulse">
                              <FileText className="h-8 w-8 text-blue-500" />
                            </div>
                          </div>
                        </div>

                        {/* Progress dots */}
                        <div className="flex justify-center space-x-1 mt-4">
                          <div
                            className="w-2 h-2 bg-blue-500 rounded-full animate-bounce"
                            style={{ animationDelay: "0ms" }}
                          ></div>
                          <div
                            className="w-2 h-2 bg-blue-500 rounded-full animate-bounce"
                            style={{ animationDelay: "150ms" }}
                          ></div>
                          <div
                            className="w-2 h-2 bg-blue-500 rounded-full animate-bounce"
                            style={{ animationDelay: "300ms" }}
                          ></div>
                        </div>
                      </div>

                      <div className="space-y-2">
                        <h3 className="text-lg font-semibold text-gray-900">Processing Documents</h3>
                        <p className="text-sm text-gray-600">Analyzing document structure and extracting keys...</p>
                        <div className="text-xs text-gray-500 bg-gray-50 px-3 py-1 rounded-full inline-block">
                          Please wait while we process your files
                        </div>
                      </div>
                    </div>
                  </div>
                ) : currentFile?.replacedFile ? (
                  <DocumentPreview
                    key={`replaced-${previewKey}-${selectedFileIndex}`}
                    file={currentFile.replacedFile}
                  />
                ) : currentFile ? (
                  <DocumentPreview key={`original-${currentFile.file.name}`} file={currentFile.file} />
                ) : (
                  <div
                    className={`flex h-full items-center justify-center text-muted-foreground border-2 border-dashed rounded-lg transition-colors ${
                      isDragOver ? "border-primary bg-primary/5" : "border-muted-foreground/25"
                    }`}
                    onDragOver={handleDragOver}
                    onDragLeave={handleDragLeave}
                    onDrop={handleDrop}
                    onClick={() => fileInputRef.current?.click()}
                  >
                    <div className="text-center cursor-pointer">
                      <Upload className={`mx-auto h-12 w-12 mb-4 ${isDragOver ? "text-primary" : ""}`} />
                      <p className="text-lg mb-2">
                        {isDragOver ? "Drop your .docx files here" : "Upload .docx files to preview them here"}
                      </p>
                      <p className="text-sm text-muted-foreground">Drag and drop files or click to browse</p>
                    </div>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>
        </div>

        {/* Right Panel - Keys List */}
        <div className="w-[520px]">
          <div className="h-full overflow-auto p-4">
            <Card className="h-full">
              <CardHeader>
                <CardTitle className="flex items-center justify-between">
                  Available Placeholders
                  <Badge variant="outline" className="text-xs">
                    {uniqueMatchedKeys.length + uniqueUnmatchedKeys.length} unique keys
                  </Badge>
                </CardTitle>
              </CardHeader>
              <CardContent className="h-[calc(100%-80px)]">
                <KeysList
                  matchedKeys={uniqueMatchedKeys}
                  unmatchedKeys={uniqueUnmatchedKeys}
                  onKeyUpdate={handleKeyUpdate}
                  customMappings={customPlaceholders}
                />
              </CardContent>
            </Card>
          </div>
        </div>
      </div>

      {/* Bottom Bar */}
      <div className="border-t bg-card p-4">
        <div className="flex justify-center gap-4">
          <Button
            onClick={handleReplaceKeys}
            disabled={
              uploadedFiles.length === 0 ||
              isProcessing ||
              isUploading ||
              (uniqueMatchedKeys.length === 0 && uniqueUnmatchedKeys.length === 0)
            }
            className="flex items-center gap-2"
          >
            {isProcessing
              ? "Processing All Files..."
              : `Replace Keys in ${uploadedFiles.length} File${uploadedFiles.length !== 1 ? "s" : ""}`}
          </Button>
          <Button
            onClick={handleDownload}
            disabled={uploadedFiles.filter((f) => f.replacedFile).length === 0}
            className="flex items-center gap-2"
            variant="outline"
          >
            <Download className="h-4 w-4" />
            Download {uploadedFiles.filter((f) => f.replacedFile).length > 1 ? "ZIP" : "File"}
          </Button>
        </div>
      </div>
    </div>
  )
}
