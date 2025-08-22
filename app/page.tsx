"use client"

import type React from "react"

import { useState, useRef } from "react"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Upload, Download, X, FileText } from "lucide-react"
import { DocumentPreview } from "@/components/document-preview"
import { KeysList } from "@/components/keys-list"
import { downloadFile } from "@/lib/download-helper"
import { Badge } from "@/components/ui/badge"

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
  const [keyMappings, setKeyMappings] = useState<{ [key: string]: string }>({})
  const [previewKey, setPreviewKey] = useState(0)
  const [isDragOver, setIsDragOver] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

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

  const handleReplaceKeys = async () => {
    if (uploadedFiles.length === 0) return

    setIsProcessing(true)
    try {
      const updatedFiles = [...uploadedFiles]

      for (let i = 0; i < updatedFiles.length; i++) {
        const fileData = updatedFiles[i]
        const formData = new FormData()
        formData.append("file", fileData.file)
        formData.append("keyMappings", JSON.stringify(keyMappings))
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
        // Single file download
        const fileData = processedFiles[0]
        const blob = new Blob([fileData.replacedFile!], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        })
        downloadFile(blob, fileData.file.name)
      } else {
        // Multiple files - create ZIP
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
          <Button onClick={() => fileInputRef.current?.click()} className="flex items-center gap-2">
            <Upload className="h-4 w-4" />
            Upload .docx Files
          </Button>
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
                {currentFile?.replacedFile ? (
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
              (uniqueMatchedKeys.length === 0 && uniqueUnmatchedKeys.length === 0)
            }
            className="flex items-center gap-2 bg-transparent"
            variant="outline"
          >
            {isProcessing
              ? "Processing All Files..."
              : `Replace Keys in ${uploadedFiles.length} File${uploadedFiles.length !== 1 ? "s" : ""}`}
          </Button>
          <Button
            onClick={handleDownload}
            disabled={uploadedFiles.filter((f) => f.replacedFile).length === 0}
            className="flex items-center gap-2"
          >
            <Download className="h-4 w-4" />
            Download {uploadedFiles.filter((f) => f.replacedFile).length > 1 ? "ZIP" : "File"}
          </Button>
        </div>
      </div>
    </div>
  )
}
