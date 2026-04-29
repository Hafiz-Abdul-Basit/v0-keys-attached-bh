"use client";
import { Toaster, toast } from "react-hot-toast";
import type React from "react";
import { useState, useRef, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import {
  Upload,
  Download,
  X,
  FileText,
  Plus,
  Map,
  ArrowLeft,
  CheckCircle2,
  AlertCircle,
  ClipboardCopy,
} from "lucide-react";
import { DocumentPreview } from "@/components/document-preview";
import { KeysList } from "@/components/keys-list";
import { downloadFile } from "@/lib/download-helper";
import { Badge } from "@/components/ui/badge";
import { Label } from "@/components/ui/label";
import * as Checkbox from "@radix-ui/react-checkbox";
import { CheckIcon } from "@radix-ui/react-icons";
import { PlaceholderAutocomplete } from "@/components/Placeholderautocomplete";
import keysData from "@/keys.json";

function extractKeys(data: unknown): string[] {
  if (Array.isArray(data)) {
    return data.map((item) =>
      typeof item === "string" ? item : String(item?.key ?? item?.name ?? item),
    );
  }
  if (data && typeof data === "object") {
    const obj = data as Record<string, unknown>;
    if (Array.isArray(obj.keys)) return extractKeys(obj.keys);
    if (Array.isArray(obj.placeholders)) return extractKeys(obj.placeholders);
    return Object.keys(obj);
  }
  return [];
}

const AVAILABLE_KEYS: string[] = extractKeys(keysData);

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
  const [customPlaceholders, setCustomPlaceholders] = useState<{
    [key: string]: string;
  }>({});
  const [selectedUnmatched, setSelectedUnmatched] = useState<string[]>([]);
  // Track processing progress per file
  const [processingIndex, setProcessingIndex] = useState<number>(-1);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const currentUniqueUnmatched = [...new Set(uploadedFiles.flatMap((f) => f.unmatchedKeys))]

    setCustomPlaceholders((prev) => {
      const updated = { ...prev };
      let hasNew = false;
      currentUniqueUnmatched.forEach((key) => {
        if (!updated[key]) {
          const autoMatch = AVAILABLE_KEYS.find(
            (k) => k.toLowerCase() === key.toLowerCase(),
          );
          updated[key] = autoMatch ? `<<${autoMatch}>>` : "";
          hasNew = true;
        }
      });
      return hasNew ? updated : prev;
    });

    setSelectedUnmatched(currentUniqueUnmatched);
  }, [uploadedFiles]);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragOver(true)
  }

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragOver(false)
  }

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const files = Array.from(e.dataTransfer.files);
    const docxFiles = files.filter(
      (file) =>
        file.type ===
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    if (docxFiles.length === 0) {
      toast.error("Please upload .docx files only");
      return;
    }
    processFiles(docxFiles);
  };

  const processFiles = async (files: File[]) => {
    setIsUploading(true);
    const newFileData: FileData[] = [];
    const toastId = toast.loading(
      `Analyzing ${files.length} file${files.length > 1 ? "s" : ""}…`,
    );

    for (const file of files) {
      try {
        const formData = new FormData();
        formData.append("file", file);
        const response = await fetch("/api/analyze", {
          method: "POST",
          body: formData,
        });
        if (response.ok) {
          const { matchedKeys: foundKeys, unmatchedKeys: unfoundKeys } = await response.json()
          newFileData.push({
            file,
            matchedKeys: foundKeys,
            unmatchedKeys: unfoundKeys || [],
          });
        }
      } catch (error) {
        console.error(`Error analyzing document ${file.name}:`, error);
        toast.error(`Failed to analyze ${file.name}`);
      }
    }

    setUploadedFiles((prev) => [...prev, ...newFileData])
    if (uploadedFiles.length === 0 && newFileData.length > 0) {
      setSelectedFileIndex(0)
    }
    setIsUploading(false);

    const totalMatched = newFileData.reduce(
      (s, f) => s + f.matchedKeys.length,
      0,
    );
    const totalUnmatched = newFileData.reduce(
      (s, f) => s + f.unmatchedKeys.length,
      0,
    );

    toast.dismiss(toastId);
    if (totalUnmatched > 0) {
      toast(
        `${newFileData.length} file${newFileData.length > 1 ? "s" : ""} loaded · ${totalMatched} matched · ${totalUnmatched} need mapping`,
        { icon: "⚠️", duration: 4000 },
      );
      // Auto-open mapping panel when there are unmatched keys
      setShowPlaceholderManager(true);
    } else {
      toast.success(
        `${newFileData.length} file${newFileData.length > 1 ? "s" : ""} loaded · ${totalMatched} keys matched`,
        { duration: 3000 },
      );
    }
  };

  const handleFileUpload = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const files = event.target.files;
    if (!files) return;
    const docxFiles = Array.from(files).filter(
      (file) =>
        file.type ===
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    if (docxFiles.length === 0) {
      toast.error("Please upload .docx files only");
      return;
    }
    await processFiles(docxFiles);
    // Reset input so the same file can be re-uploaded
    event.target.value = "";
  };

  const removeFile = (index: number) => {
    const name = uploadedFiles[index].file.name;
    setUploadedFiles((prev) => prev.filter((_, i) => i !== index));
    if (selectedFileIndex >= uploadedFiles.length - 1) {
      setSelectedFileIndex(Math.max(0, uploadedFiles.length - 2))
    }
    toast(`Removed ${name}`, { icon: "🗑️", duration: 2000 });
  };

  const handleKeyUpdate = (key: string, value: string) => {
    setKeyMappings((prev) => ({ ...prev, [key]: value }));
  };

  const handleAddPlaceholder = () => {
    if (!newPlaceholderKey.trim() || !newPlaceholderValue.trim()) {
      toast.error("Both key and placeholder value are required");
      return;
    }
    const formattedKey = newPlaceholderKey.toUpperCase().trim();
    const formattedValue = newPlaceholderValue.startsWith("<<")
      ? newPlaceholderValue
      : `<<${newPlaceholderValue.toUpperCase().trim()}>>`;

    if (customPlaceholders[formattedKey] !== undefined) {
      toast.error(`Mapping for "${formattedKey}" already exists`);
      return;
    }

    setCustomPlaceholders((prev) => ({
      ...prev,
      [formattedKey]: formattedValue,
    }));
    setNewPlaceholderKey("");
    setNewPlaceholderValue("");
    toast.success(`Mapping added: ${formattedKey} → ${formattedValue}`);
  };

  const handleRemovePlaceholder = (key: string) => {
    setCustomPlaceholders((prev) => {
      const updated = { ...prev };
      delete updated[key];
      return updated;
    });
    toast(`Removed mapping for ${key}`, { icon: "🗑️", duration: 2000 });
  };

  const handleReplaceKeys = async () => {
    if (uploadedFiles.length === 0) return

    // Warn if unmatched keys still have no mapping filled in
    const unmappedStillEmpty = uniqueUnmatchedKeys.filter(
      (k) => !customPlaceholders[k],
    );
    if (unmappedStillEmpty.length > 0) {
      toast(
        `${unmappedStillEmpty.length} key${unmappedStillEmpty.length > 1 ? "s" : ""} still have no mapping and will be skipped`,
        { icon: "⚠️", duration: 4000 },
      );
    }

    setIsProcessing(true);
    const toastId = toast.loading("Processing files…");
    try {
      const updatedFiles = [...uploadedFiles];
      for (let i = 0; i < updatedFiles.length; i++) {
        setProcessingIndex(i);
        const fileData = updatedFiles[i];
        const formData = new FormData();
        formData.append("file", fileData.file);
        const allKeyMappings = { ...keyMappings, ...customPlaceholders };
        formData.append("keyMappings", JSON.stringify(allKeyMappings));
        formData.append("foundKeys", JSON.stringify(fileData.matchedKeys));
        formData.append(
          "unmatchedKeys",
          JSON.stringify(fileData.unmatchedKeys),
        );
        const response = await fetch("/api/replace", {
          method: "POST",
          body: formData,
        });
        if (!response.ok) {
          throw new Error(`Failed to process document ${fileData.file.name}`)
        }
        const blob = await response.blob();
        const replacedFile = new File([blob], fileData.file.name, {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        updatedFiles[i] = { ...fileData, replacedFile };
      }
      setUploadedFiles(updatedFiles);
      setPreviewKey((prev) => prev + 1);
      toast.dismiss(toastId);
      toast.success(
        `${updatedFiles.length} file${updatedFiles.length > 1 ? "s" : ""} processed successfully!`,
        { duration: 3000 },
      );
    } catch (error) {
      console.error("Error replacing keys:", error);
      toast.dismiss(toastId);
      toast.error("Error replacing keys. Please try again.");
    } finally {
      setIsProcessing(false);
      setProcessingIndex(-1);
    }
  }

  const handleDownload = async () => {
    const processedFiles = uploadedFiles.filter((f) => f.replacedFile);
    if (processedFiles.length === 0) {
      toast.error("Please replace keys first before downloading");
      return;
    }
    try {
      if (processedFiles.length === 1) {
        const fileData = processedFiles[0]
        const blob = new Blob([fileData.replacedFile!], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        downloadFile(blob, fileData.file.name);
        toast.success(`Downloaded ${fileData.file.name}`);
      } else {
        const toastId = toast.loading("Zipping files…");
        const JSZip = (await import("jszip")).default;
        const zip = new JSZip();
        processedFiles.forEach((fileData) => {
          zip.file(fileData.file.name, fileData.replacedFile!);
        });
        const zipBlob = await zip.generateAsync({ type: "blob" });
        downloadFile(zipBlob, "processed_documents.zip");
        toast.dismiss(toastId);
        toast.success(`Downloaded ${processedFiles.length} files as ZIP`);
      }
    } catch (error) {
      console.error("Error downloading documents:", error);
      toast.error("Error downloading documents. Please try again.");
    }
  }

  const currentFile = uploadedFiles[selectedFileIndex];
  const allMatchedKeys = uploadedFiles.flatMap((f) => f.matchedKeys);
  const allUnmatchedKeys = uploadedFiles.flatMap((f) => f.unmatchedKeys);
  const uniqueMatchedKeys = [...new Set(allMatchedKeys)];
  const uniqueUnmatchedKeys = [...new Set(allUnmatchedKeys)];
  const totalUniqueKeys = uniqueMatchedKeys.length + uniqueUnmatchedKeys.length;

  // How many unmatched keys already have a value filled in
  const mappedCount = uniqueUnmatchedKeys.filter(
    (k) => customPlaceholders[k],
  ).length;
  const allMapped =
    uniqueUnmatchedKeys.length === 0 ||
    mappedCount === uniqueUnmatchedKeys.length;

  return (
    <>
      {/* Toast container */}
      <Toaster
        position="bottom-right"
        toastOptions={{ style: { fontSize: "13px" } }}
      />

      <div className="min-h-screen bg-background flex flex-col">
        {/* ── Top Bar ── */}
        <div className="border-b bg-card p-4 flex-shrink-0">
          <div className="flex items-center justify-between">
            <h1 className="text-2xl font-bold">DocX Key Replacer</h1>
            <div className="flex items-center gap-2">
              <Button
                variant={showPlaceholderManager ? "default" : "outline"}
                className="flex items-center gap-2"
                disabled={uploadedFiles.length === 0}
                onClick={() => setShowPlaceholderManager((v) => !v)}
              >
                {showPlaceholderManager ? (
                  <>
                    <ArrowLeft className="h-4 w-4" />
                    Back to Placeholders
                  </>
                ) : (
                  <>
                    <Map className="h-4 w-4" />
                    Manage Mappings
                    {uniqueUnmatchedKeys.length > 0 && (
                      <Badge variant="secondary" className="ml-1">
                        {uniqueUnmatchedKeys.length}
                      </Badge>
                    )}
                  </>
                )}
              </Button>

              <Button
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-2"
              >
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

        {/* ── File List Bar ── */}
        {uploadedFiles.length > 0 && (
          <div className="border-b bg-muted/30 p-2 flex-shrink-0">
            <div className="flex items-center gap-2 overflow-x-auto">
              <span className="text-sm font-medium text-muted-foreground whitespace-nowrap">
                Files:
              </span>
              {uploadedFiles.map((fileData, index) => (
                <div
                  key={index}
                  className={`flex items-center gap-2 px-3 py-1 rounded-md border cursor-pointer transition-colors ${
                    selectedFileIndex === index
                      ? "bg-primary text-primary-foreground"
                      : "bg-background hover:bg-muted"
                  }`}
                  onClick={() => setSelectedFileIndex(index)}
                >
                  <FileText className="h-3 w-3 flex-shrink-0" />
                  <span className="text-xs truncate max-w-32">
                    {fileData.file.name}
                  </span>
                  {/* Processing spinner for this specific file */}
                  {isProcessing && processingIndex === index && (
                    <span className="text-xs animate-spin">⚙️</span>
                  )}
                  {/* Done check */}
                  {fileData.replacedFile && !isProcessing && (
                    <Badge variant="secondary" className="text-xs px-1 py-0">
                      ✓
                    </Badge>
                  )}
                  <Button
                    variant="ghost"
                    size="sm"
                    className="h-4 w-4 p-0 hover:bg-destructive hover:text-destructive-foreground"
                    onClick={(e) => {
                      e.stopPropagation();
                      removeFile(index);
                    }}
                  >
                    <X className="h-3 w-3" />
                  </Button>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── Main Content ── */}
        <div className="flex h-[calc(100vh-180px)]">
          {/* ── Left: Document Preview (never changes) ── */}
          <div className="flex-1 border-r">
            <div className="h-full overflow-auto p-4">
              <Card className="h-full">
                <CardHeader>
                  <CardTitle>
                    Document Preview
                    {currentFile && (
                      <span className="text-sm font-normal text-muted-foreground ml-2">
                        ({currentFile.file.name})
                      </span>
                    )}
                  </CardTitle>
                </CardHeader>
                <CardContent className="h-[calc(100%-80px)]">
                  {isUploading ? (
                    <div className="flex h-full items-center justify-center">
                      <div className="text-center space-y-6">
                        <div className="relative">
                          <div className="w-20 h-20 mx-auto relative">
                            <div className="absolute inset-0 border-4 border-gray-200 rounded-full" />
                            <div className="absolute inset-0 border-4 border-blue-500 border-t-transparent rounded-full animate-spin" />
                            <div className="absolute inset-0 flex items-center justify-center">
                              <div className="bg-white rounded-lg p-2 shadow-sm animate-pulse">
                                <FileText className="h-8 w-8 text-blue-500" />
                              </div>
                            </div>
                          </div>
                          <div className="flex justify-center space-x-1 mt-4">
                            <div
                              className="w-2 h-2 bg-blue-500 rounded-full animate-bounce"
                              style={{ animationDelay: "0ms" }}
                            />
                            <div
                              className="w-2 h-2 bg-blue-500 rounded-full animate-bounce"
                              style={{ animationDelay: "150ms" }}
                            />
                            <div
                              className="w-2 h-2 bg-blue-500 rounded-full animate-bounce"
                              style={{ animationDelay: "300ms" }}
                            />
                          </div>
                        </div>
                        <div className="space-y-2">
                          <h3 className="text-lg font-semibold text-gray-900">
                            Processing Documents
                          </h3>
                          <p className="text-sm text-gray-600">
                            Analyzing document structure and extracting keys...
                          </p>
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
                    <DocumentPreview
                      key={`original-${currentFile.file.name}`}
                      file={currentFile.file}
                    />
                  ) : (
                    <div
                      className={`flex h-full items-center justify-center text-muted-foreground border-2 border-dashed rounded-lg transition-colors ${
                        isDragOver
                          ? "border-primary bg-primary/5"
                          : "border-muted-foreground/25"
                      }`}
                      onDragOver={handleDragOver}
                      onDragLeave={handleDragLeave}
                      onDrop={handleDrop}
                      onClick={() => fileInputRef.current?.click()}
                    >
                      <div className="text-center cursor-pointer">
                        <Upload
                          className={`mx-auto h-12 w-12 mb-4 ${isDragOver ? "text-primary" : ""}`}
                        />
                        <p className="text-lg mb-2">
                          {isDragOver
                            ? "Drop your .docx files here"
                            : "Upload .docx files to preview them here"}
                        </p>
                        <p className="text-sm text-muted-foreground">
                          Drag and drop files or click to browse
                        </p>
                      </div>
                    </div>
                  )}
                </CardContent>
              </Card>
            </div>
          </div>

          {/* ── Right Panel: swaps between Available Placeholders ↔ Mapping Manager ── */}
          <div className="w-[520px] flex-shrink-0">
            <div className="h-full overflow-auto p-4">
              <Card className="h-full flex flex-col">
                {/* ── AVAILABLE PLACEHOLDERS VIEW ── */}
                {!showPlaceholderManager && (
                  <>
                    <CardHeader className="flex-shrink-0">
                      <CardTitle className="flex items-center justify-between flex-wrap gap-2">
                        <span>Available Placeholders</span>
                        {/* ── Keys count row ── */}
                        <div className="flex items-center gap-2 flex-wrap">
                          <Badge variant="outline" className="text-xs gap-1">
                            {totalUniqueKeys} unique keys
                          </Badge>
                          {uniqueMatchedKeys.length > 0 && (
                            <Badge
                              variant="secondary"
                              className="text-xs gap-1 bg-green-100 text-green-700 border-green-200"
                            >
                              <CheckCircle2 className="h-3 w-3" />
                              {uniqueMatchedKeys.length} matched
                            </Badge>
                          )}
                          {uniqueUnmatchedKeys.length > 0 && (
                            <Badge
                              variant="secondary"
                              className="text-xs gap-1 bg-orange-100 text-orange-700 border-orange-200"
                            >
                              <AlertCircle className="h-3 w-3" />
                              {uniqueUnmatchedKeys.length} need mapping
                            </Badge>
                          )}
                        </div>
                      </CardTitle>

                      {/* Mapping progress bar — only shown when there are unmatched keys */}
                      {uniqueUnmatchedKeys.length > 0 && (
                        <div className="mt-2">
                          <div className="flex items-center justify-between text-xs text-gray-500 mb-1">
                            <span>Mapping progress</span>
                            <span>
                              {mappedCount}/{uniqueUnmatchedKeys.length} mapped
                            </span>
                          </div>
                          <div className="w-full h-1.5 bg-gray-100 rounded-full overflow-hidden">
                            <div
                              className={`h-full rounded-full transition-all duration-500 ${
                                allMapped ? "bg-green-500" : "bg-orange-400"
                              }`}
                              style={{
                                width: `${uniqueUnmatchedKeys.length === 0 ? 100 : (mappedCount / uniqueUnmatchedKeys.length) * 100}%`,
                              }}
                            />
                          </div>
                        </div>
                      )}
                    </CardHeader>
                    <CardContent className="flex-1 overflow-auto">
                      <KeysList
                        matchedKeys={uniqueMatchedKeys}
                        unmatchedKeys={uniqueUnmatchedKeys}
                        onKeyUpdate={handleKeyUpdate}
                        customMappings={customPlaceholders}
                      />
                    </CardContent>
                  </>
                )}

                {/* ── MAPPING MANAGER VIEW ── */}
                {showPlaceholderManager && (
                  <>
                    <CardHeader className="flex-shrink-0 border-b pb-3">
                      <CardTitle className="flex items-center gap-2 text-base">
                        <Map className="h-4 w-4 text-gray-500" />
                        Placeholder Mapping
                        {uniqueUnmatchedKeys.length > 0 && (
                          <Badge variant="destructive" className="ml-1 text-xs">
                            {uniqueUnmatchedKeys.length} unmatched
                          </Badge>
                        )}
                      </CardTitle>

                      {/* Progress bar inside mapping panel too */}
                      {uniqueUnmatchedKeys.length > 0 && (
                        <div className="mt-3">
                          <div className="flex items-center justify-between text-xs text-gray-500 mb-1">
                            <span>Mapping progress</span>
                            <span
                              className={
                                allMapped ? "text-green-600 font-medium" : ""
                              }
                            >
                              {mappedCount}/{uniqueUnmatchedKeys.length} mapped
                              {allMapped && " ✓"}
                            </span>
                          </div>
                          <div className="w-full h-1.5 bg-gray-100 rounded-full overflow-hidden">
                            <div
                              className={`h-full rounded-full transition-all duration-500 ${
                                allMapped ? "bg-green-500" : "bg-orange-400"
                              }`}
                              style={{
                                width: `${(mappedCount / uniqueUnmatchedKeys.length) * 100}%`,
                              }}
                            />
                          </div>
                        </div>
                      )}
                    </CardHeader>

                    <CardContent className="flex-1 overflow-y-auto space-y-6">
                      {/* ── Unmatched Keys ── */}
                      {uniqueUnmatchedKeys.length > 0 && (
                        <div className="space-y-4">
                          <div>
                            <h4 className="font-medium text-sm mb-1">
                              Unmatched Keys Found in Documents:
                            </h4>
                            <p className="text-xs text-gray-500">
                              These keys don&apos;t match any known placeholder.
                              Pick a suggestion or type a custom value.
                            </p>
                          </div>

                          {/* Select All */}
                          <div className="flex items-center gap-2">
                            <Checkbox.Root
                              id="select-all"
                              checked={
                                selectedUnmatched.length ===
                                uniqueUnmatchedKeys.length
                              }
                              onCheckedChange={(checked) =>
                                setSelectedUnmatched(
                                  checked ? [...uniqueUnmatchedKeys] : [],
                                )
                              }
                              className="w-5 h-5 border border-gray-300 rounded data-[state=checked]:bg-blue-600 data-[state=checked]:text-white flex items-center justify-center flex-shrink-0"
                            >
                              <Checkbox.Indicator>
                                <CheckIcon className="w-4 h-4" />
                              </Checkbox.Indicator>
                            </Checkbox.Root>
                            <Label
                              htmlFor="select-all"
                              className="text-sm cursor-pointer"
                            >
                              Select All
                            </Label>
                          </div>

                          <div className="grid gap-3">
                            {uniqueUnmatchedKeys
                              .filter(
                                (key) => customPlaceholders[key] !== undefined,
                              )
                              .map((key) => {
                                const isMapped = !!customPlaceholders[key];
                                return (
                                  <div
                                    key={key}
                                    className={`flex items-start gap-3 p-3 rounded-lg border transition-colors ${
                                      isMapped
                                        ? "bg-green-50 border-green-200"
                                        : "bg-gray-50 border-gray-200"
                                    }`}
                                  >
                                    <Checkbox.Root
                                      id={`select-${key}`}
                                      checked={selectedUnmatched.includes(key)}
                                      onCheckedChange={(checked) =>
                                        setSelectedUnmatched((prev) =>
                                          checked
                                            ? [...prev, key]
                                            : prev.filter((k) => k !== key),
                                        )
                                      }
                                      className="w-5 h-5 border border-gray-300 rounded mt-1 data-[state=checked]:bg-blue-600 data-[state=checked]:text-white flex items-center justify-center flex-shrink-0"
                                    >
                                      <Checkbox.Indicator>
                                        <CheckIcon className="w-4 h-4" />
                                      </Checkbox.Indicator>
                                    </Checkbox.Root>

                                    <div className="flex-1 space-y-2 min-w-0">
                                      <div>
                                        <Label className="text-xs text-gray-500 mb-1 flex items-center gap-1">
                                          Key found in document:
                                          {/* Copy key to clipboard */}
                                          <button
                                            type="button"
                                            onClick={() => {
                                              navigator.clipboard.writeText(
                                                key,
                                              );
                                              toast("Copied to clipboard", {
                                                icon: "📋",
                                                duration: 1500,
                                              });
                                            }}
                                            className="text-gray-400 hover:text-gray-600 transition-colors ml-auto"
                                            title="Copy key"
                                          >
                                            <ClipboardCopy className="h-3 w-3" />
                                          </button>
                                        </Label>
                                        <div className="text-sm font-mono bg-white p-2 rounded break-all border border-gray-200 flex items-center justify-between gap-2">
                                          <span>{key}</span>
                                          {isMapped && (
                                            <CheckCircle2 className="h-4 w-4 text-green-500 flex-shrink-0" />
                                          )}
                                        </div>
                                      </div>
                                      <div>
                                        <Label className="text-xs text-gray-500 mb-1 block">
                                          Map to placeholder:
                                        </Label>
                                        <PlaceholderAutocomplete
                                          id={`value-${key}`}
                                          value={customPlaceholders[key] || ""}
                                          onChange={(val: string) =>
                                            setCustomPlaceholders((prev) => ({
                                              ...prev,
                                              [key]: val,
                                            }))
                                          }
                                          placeholder={`e.g., <<${key}>>`}
                                          availableKeys={AVAILABLE_KEYS}
                                        />
                                      </div>
                                    </div>
                                  </div>
                                );
                              })}
                          </div>

                          <Button
                            onClick={() => {
                              const newCustom = { ...customPlaceholders };
                              const removedCount = uniqueUnmatchedKeys.filter(
                                (key) => !selectedUnmatched.includes(key),
                              ).length;
                              if (removedCount > 0) {
                                uniqueUnmatchedKeys.forEach((key) => {
                                  if (!selectedUnmatched.includes(key))
                                    delete newCustom[key];
                                });
                                setCustomPlaceholders(newCustom);
                                setSelectedUnmatched((prev) =>
                                  prev.filter(
                                    (key) => newCustom[key] !== undefined,
                                  ),
                                );
                                toast(
                                  `Removed ${removedCount} unselected key${removedCount > 1 ? "s" : ""}`,
                                  {
                                    icon: "🗑️",
                                    duration: 2500,
                                  },
                                );
                              } else {
                                toast("No unselected keys to remove", {
                                  icon: "ℹ️",
                                });
                              }
                            }}
                            variant="destructive"
                            disabled={uniqueUnmatchedKeys.length === 0}
                            className="w-full mt-1"
                          >
                            Remove Unselected
                          </Button>
                        </div>
                      )}

                      {/* ── Add Custom Mapping ── */}
                      <div className="border-t pt-4 space-y-3">
                        <h4 className="font-medium text-sm">
                          Add Custom Mapping:
                        </h4>
                        <div className="space-y-3">
                          <div>
                            <Label
                              htmlFor="new-key"
                              className="text-xs text-gray-500 mb-1 block"
                            >
                              Key found in document:
                            </Label>
                            <PlaceholderAutocomplete
                              id="new-key"
                              value={newPlaceholderKey}
                              onChange={(val: string) => {
                                const stripped = val
                                  .replace(/^<</, "")
                                  .replace(/>>$/, "")
                                  .toUpperCase();
                                setNewPlaceholderKey(stripped);
                              }}
                              placeholder="e.g., ABSENCESDATESLOSS"
                              availableKeys={AVAILABLE_KEYS}
                            />
                          </div>
                          <div>
                            <Label
                              htmlFor="new-value"
                              className="text-xs text-gray-500 mb-1 block"
                            >
                              Map to placeholder:
                            </Label>
                            <PlaceholderAutocomplete
                              id="new-value"
                              value={newPlaceholderValue}
                              onChange={setNewPlaceholderValue}
                              placeholder="e.g., <<ABSENCESDATESLOSS>>"
                              availableKeys={AVAILABLE_KEYS}
                            />
                          </div>
                          <Button
                            onClick={handleAddPlaceholder}
                            className="flex items-center gap-2 w-full"
                          >
                            <Plus className="h-4 w-4" />
                            Add Mapping
                          </Button>
                        </div>
                      </div>

                      {/* ── Custom Mappings List ── */}
                      {Object.keys(customPlaceholders).filter(
                        (key) => !uniqueUnmatchedKeys.includes(key),
                      ).length > 0 && (
                        <div className="space-y-2 border-t pt-4">
                          <h4 className="font-medium text-sm">
                            Custom Mappings:
                            <Badge variant="outline" className="ml-2 text-xs">
                              {
                                Object.keys(customPlaceholders).filter(
                                  (key) => !uniqueUnmatchedKeys.includes(key),
                                ).length
                              }
                            </Badge>
                          </h4>
                          <div className="grid gap-2">
                            {Object.entries(customPlaceholders)
                              .filter(
                                ([key]) => !uniqueUnmatchedKeys.includes(key),
                              )
                              .map(([key, value]) => (
                                <div
                                  key={key}
                                  className="flex items-center gap-2 p-3 bg-gray-50 rounded-lg border border-gray-200"
                                >
                                  <div className="flex-1 min-w-0">
                                    <div className="text-sm font-mono break-all">
                                      <span className="text-blue-600 font-semibold">
                                        {key}
                                      </span>
                                      <span className="mx-2 text-gray-500">
                                        →
                                      </span>
                                      <span>{value}</span>
                                    </div>
                                  </div>
                                  <Button
                                    variant="ghost"
                                    size="sm"
                                    onClick={() => handleRemovePlaceholder(key)}
                                    className="h-8 w-8 p-0 hover:bg-red-100 hover:text-red-600 flex-shrink-0"
                                  >
                                    <X className="h-4 w-4" />
                                  </Button>
                                </div>
                              ))}
                          </div>
                        </div>
                      )}
                    </CardContent>
                  </>
                )}
              </Card>
            </div>
          </div>
        </div>

        {/* ── Bottom Bar ── */}
        <div className="border-t bg-card p-4 flex-shrink-0">
          <div className="flex justify-center gap-4">
            <Button
              onClick={handleReplaceKeys}
              disabled={
                uploadedFiles.length === 0 ||
                isProcessing ||
                isUploading ||
                (uniqueMatchedKeys.length === 0 &&
                  uniqueUnmatchedKeys.length === 0)
              }
              className="flex items-center gap-2"
            >
              {isProcessing
                ? `Processing file ${processingIndex + 1} of ${uploadedFiles.length}…`
                : `Replace Keys in ${uploadedFiles.length} File${uploadedFiles.length !== 1 ? "s" : ""}`}
            </Button>
            <Button
              onClick={handleDownload}
              disabled={
                uploadedFiles.filter((f) => f.replacedFile).length === 0
              }
              className="flex items-center gap-2"
              variant="outline"
            >
              <Download className="h-4 w-4" />
              Download{" "}
              {uploadedFiles.filter((f) => f.replacedFile).length > 1
                ? "ZIP"
                : "File"}
            </Button>
          </div>
        </div>
      </div>
    </>
  );
}
