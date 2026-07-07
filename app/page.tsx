"use client";
import { Toaster, toast } from "react-hot-toast";
import { ReplacementAnimation } from "@/components/replacement-animation";
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
  RotateCcw,
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
      typeof item === "string"
        ? item
        : String(
            (item as Record<string, unknown>)?.key ??
              (item as Record<string, unknown>)?.name ??
              item,
          ),
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
  file: File;
  matchedKeys: string[];
  unmatchedKeys: string[];
  replacedFile?: File;
  replacedMatchedKeys?: string[];
  replacedUnmatchedKeys?: string[];
}

export default function Home() {
  const [uploadedFiles, setUploadedFiles] = useState<FileData[]>([]);
  const [selectedFileIndex, setSelectedFileIndex] = useState<number>(0);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [keyMappings, setKeyMappings] = useState<Record<string, string>>({});
  const [previewKey, setPreviewKey] = useState(0);
  const [isDragOver, setIsDragOver] = useState(false);
  const [showPlaceholderManager, setShowPlaceholderManager] = useState(false);
  const [newPlaceholderKey, setNewPlaceholderKey] = useState("");
  const [newPlaceholderValue, setNewPlaceholderValue] = useState("");
  const [customPlaceholders, setCustomPlaceholders] = useState<
    Record<string, string>
  >({});
  const [selectedUnmatched, setSelectedUnmatched] = useState<string[]>([]);
  const [processingIndex, setProcessingIndex] = useState<number>(-1);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const currentUniqueUnmatched = [
      ...new Set(uploadedFiles.flatMap((f) => f.unmatchedKeys)),
    ];

    setCustomPlaceholders((prev) => {
      const updated = { ...prev };
      let hasNew = false;
      currentUniqueUnmatched.forEach((key) => {
        if (updated[key] === undefined) {
          const norm = key
            .replace(/^<</, "")
            .replace(/>>$/, "")
            .replace(/\s+/g, "")
            .toUpperCase();
          const autoMatch = AVAILABLE_KEYS.find(
            (k) => k.replace(/\s+/g, "").toUpperCase() === norm,
          );
          updated[key] = autoMatch ? `<<${autoMatch}>>` : "";
          hasNew = true;
        }
      });
      return hasNew ? updated : prev;
    });

    setSelectedUnmatched(currentUniqueUnmatched);
  }, [uploadedFiles]);

  /* ── reset ── */
  const handleReset = () => {
    setUploadedFiles([]);
    setSelectedFileIndex(0);
    setKeyMappings({});
    setCustomPlaceholders({});
    setSelectedUnmatched([]);
    setShowPlaceholderManager(false);
    setNewPlaceholderKey("");
    setNewPlaceholderValue("");
    setPreviewKey((p) => p + 1);
    toast("All data reset", { icon: "🔄", duration: 2000 });
  };

  /* ── file helpers ── */
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };
  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  };
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const docxFiles = Array.from(e.dataTransfer.files).filter(
      (f) =>
        f.type ===
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    if (!docxFiles.length) {
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
        const res = await fetch("/api/analyze", {
          method: "POST",
          body: formData,
        });
        if (res.ok) {
          const { matchedKeys, unmatchedKeys } = await res.json();
          newFileData.push({
            file,
            matchedKeys,
            unmatchedKeys: unmatchedKeys || [],
          });
        }
      } catch (err) {
        console.error(`Error analyzing ${file.name}:`, err);
        toast.error(`Failed to analyze ${file.name}`);
      }
    }

    setUploadedFiles((prev) => [...prev, ...newFileData]);
    if (uploadedFiles.length === 0 && newFileData.length > 0)
      setSelectedFileIndex(0);
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
    } else {
      toast.success(
        `${newFileData.length} file${newFileData.length > 1 ? "s" : ""} loaded · ${totalMatched} keys matched`,
        { duration: 3000 },
      );
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files) return;
    const docxFiles = Array.from(e.target.files).filter(
      (f) =>
        f.type ===
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    if (!docxFiles.length) {
      toast.error("Please upload .docx files only");
      return;
    }
    await processFiles(docxFiles);
    e.target.value = "";
  };

  const removeFile = (index: number) => {
    const name = uploadedFiles[index].file.name;
    setUploadedFiles((prev) => prev.filter((_, i) => i !== index));
    if (selectedFileIndex >= uploadedFiles.length - 1)
      setSelectedFileIndex(Math.max(0, uploadedFiles.length - 2));
    toast(`Removed ${name}`, { icon: "🗑️", duration: 2000 });
  };

  const handleKeyUpdate = (key: string, value: string) =>
    setKeyMappings((prev) => ({ ...prev, [key]: value }));

  /* ── custom placeholder helpers ── */
  const handleAddPlaceholder = () => {
    if (!newPlaceholderKey.trim() || !newPlaceholderValue.trim()) {
      toast.error("Both text and placeholder value are required");
      return;
    }
    const rawKey = newPlaceholderKey.trim();
    const formattedValue = newPlaceholderValue.trim().startsWith("<<")
      ? newPlaceholderValue.trim()
      : `<<${newPlaceholderValue.trim().toUpperCase()}>>`;

    if (customPlaceholders[rawKey] !== undefined) {
      toast.error(`Mapping for "${rawKey}" already exists`);
      return;
    }
    setCustomPlaceholders((prev) => ({ ...prev, [rawKey]: formattedValue }));
    setNewPlaceholderKey("");
    setNewPlaceholderValue("");
    toast.success(`Mapping added: "${rawKey}" → ${formattedValue}`);
  };

  const handleRemovePlaceholder = (key: string) => {
    setCustomPlaceholders((prev) => {
      const u = { ...prev };
      delete u[key];
      return u;
    });
    toast(`Removed mapping for "${key}"`, { icon: "🗑️", duration: 2000 });
  };

  /* ── Remove SELECTED unmatched keys ── */
  const handleRemoveSelected = () => {
    if (selectedUnmatched.length === 0) {
      toast("No keys selected to remove", { icon: "ℹ️" });
      return;
    }
    const count = selectedUnmatched.length;
    setCustomPlaceholders((prev) => {
      const updated = { ...prev };
      selectedUnmatched.forEach((k) => delete updated[k]);
      return updated;
    });
    setUploadedFiles((prev) =>
      prev.map((f) => ({
        ...f,
        unmatchedKeys: f.unmatchedKeys.filter(
          (k) => !selectedUnmatched.includes(k),
        ),
      })),
    );
    setSelectedUnmatched([]);
    toast(`Removed ${count} selected key${count > 1 ? "s" : ""}`, {
      icon: "🗑️",
      duration: 2500,
    });
  };

  /* ── replace & download ── */
  const handleReplaceKeys = async () => {
    if (!uploadedFiles.length) return;
    const unmappedEmpty = uniqueUnmatchedKeys.filter(
      (k) => !customPlaceholders[k],
    );
    if (unmappedEmpty.length > 0) {
      toast(
        `${unmappedEmpty.length} key${unmappedEmpty.length > 1 ? "s" : ""} have no mapping and will be skipped`,
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
        formData.append(
          "keyMappings",
          JSON.stringify({ ...keyMappings, ...customPlaceholders }),
        );
        formData.append("foundKeys", JSON.stringify(fileData.matchedKeys));
        formData.append(
          "unmatchedKeys",
          JSON.stringify(fileData.unmatchedKeys),
        );
        const res = await fetch("/api/replace", {
          method: "POST",
          body: formData,
        });
        if (!res.ok) throw new Error(`Failed to process ${fileData.file.name}`);
        const blob = await res.blob();
        const replacedFile = new File([blob], fileData.file.name, {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });

        /* ── RE-ANALYZE replaced file for accurate highlighting ── */
        let replacedMatchedKeys: string[] = [];
        let replacedUnmatchedKeys: string[] = [];
        try {
          const analyzeForm = new FormData();
          analyzeForm.append("file", replacedFile);
          const analyzeRes = await fetch("/api/analyze", {
            method: "POST",
            body: analyzeForm,
          });
          if (analyzeRes.ok) {
            const analysis = await analyzeRes.json();
            replacedMatchedKeys = analysis.matchedKeys || [];
            replacedUnmatchedKeys = analysis.unmatchedKeys || [];
          }
        } catch (err) {
          console.error("Re-analysis failed:", err);
          // fallback to original keys if re-analysis fails
          replacedMatchedKeys = fileData.matchedKeys;
          replacedUnmatchedKeys = fileData.unmatchedKeys;
        }

        updatedFiles[i] = {
          ...fileData,
          replacedFile,
          replacedMatchedKeys,
          replacedUnmatchedKeys,
        };
      }
      setUploadedFiles(updatedFiles);
      setPreviewKey((p) => p + 1);
      toast.dismiss(toastId);
      toast.success(
        `${updatedFiles.length} file${updatedFiles.length > 1 ? "s" : ""} processed!`,
        { duration: 3000 },
      );
    } catch (err) {
      console.error(err);
      toast.dismiss(toastId);
      toast.error("Error replacing keys. Please try again.");
    } finally {
      setIsProcessing(false);
      setProcessingIndex(-1);
    }
  };

  const handleDownload = async () => {
    const processedFiles = uploadedFiles.filter((f) => f.replacedFile);
    if (!processedFiles.length) {
      toast.error("Please replace keys first before downloading");
      return;
    }
    try {
      if (processedFiles.length === 1) {
        const fd = processedFiles[0];
        downloadFile(
          new Blob([fd.replacedFile!], {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          }),
          fd.file.name,
        );
        toast.success(`Downloaded ${fd.file.name}`);
      } else {
        const toastId = toast.loading("Zipping files…");
        const JSZip = (await import("jszip")).default;
        const zip = new JSZip();
        processedFiles.forEach((fd) =>
          zip.file(fd.file.name, fd.replacedFile!),
        );
        const zipBlob = await zip.generateAsync({ type: "blob" });
        downloadFile(zipBlob, "processed_documents.zip");
        toast.dismiss(toastId);
        toast.success(`Downloaded ${processedFiles.length} files as ZIP`);
      }
    } catch (err) {
      console.error(err);
      toast.error("Error downloading. Please try again.");
    }
  };

  /* ── derived state ── */
  const currentFile = uploadedFiles[selectedFileIndex];
  const uniqueMatchedKeys = [
    ...new Set(uploadedFiles.flatMap((f) => f.matchedKeys)),
  ];
  const uniqueUnmatchedKeys = [
    ...new Set(uploadedFiles.flatMap((f) => f.unmatchedKeys)),
  ];
  const totalUniqueKeys = uniqueMatchedKeys.length + uniqueUnmatchedKeys.length;
  const mappedCount = uniqueUnmatchedKeys.filter(
    (k) => customPlaceholders[k],
  ).length;
  const allMapped =
    uniqueUnmatchedKeys.length === 0 ||
    mappedCount === uniqueUnmatchedKeys.length;

  /* ── render ── */
  return (
    <>
      <Toaster
        position="bottom-right"
        toastOptions={{ style: { fontSize: "13px" } }}
      />

      <div className="min-h-screen bg-background flex flex-col">
        {/* Top Bar */}
        <div className="border-b bg-card p-4 flex-shrink-0">
          <div className="flex items-center justify-between">
            <h1 className="text-2xl font-bold flex items-center gap-3">
              <span>
                DocX Key Replacer{" "}
                <span className="text-sm text-gray-500">
                  ( Developed by Abdul Basit )
                </span>
              </span>
              {uploadedFiles.length > 0 && (
                <Badge variant="secondary" className="text-xs mt-3">
                  {uploadedFiles.length} file
                  {uploadedFiles.length !== 1 ? "s" : ""} uploaded
                </Badge>
              )}
            </h1>

            <div className="flex items-center gap-2">
              <Button
                variant="outline"
                className="flex items-center gap-2"
                onClick={handleReset}
                disabled={
                  !uploadedFiles.length &&
                  !Object.keys(customPlaceholders).length
                }
              >
                <RotateCcw className="h-4 w-4" />
                Reset
              </Button>
              <Button
                variant={showPlaceholderManager ? "default" : "outline"}
                className="flex items-center gap-2"
                disabled={!uploadedFiles.length}
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

        {/* File List Bar */}
        {uploadedFiles.length > 0 && (
          <div className="border-b bg-muted/30 p-2 flex-shrink-0">
            <div className="flex items-center gap-2 overflow-x-auto">
              <span className="text-sm font-medium text-muted-foreground whitespace-nowrap">
                Files:
              </span>
              {uploadedFiles.map((fileData, index) => (
                <div
                  key={index}
                  title={fileData.file.name}
                  className={`flex items-center gap-2 px-3 py-1 rounded-md border cursor-pointer transition-colors ${selectedFileIndex === index ? "bg-primary text-primary-foreground" : "bg-background hover:bg-muted"}`}
                  onClick={() => setSelectedFileIndex(index)}
                >
                  <FileText className="h-3 w-3 flex-shrink-0" />
                  <span className="text-xs truncate max-w-32">
                    {fileData.file.name}
                  </span>
                  {isProcessing && processingIndex === index && (
                    <span className="text-xs animate-spin">⚙️</span>
                  )}
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

        {/* Main Content */}
        <div className="flex h-[calc(100vh-220px)]">
          {/* Left: Document Preview */}
          <div className="flex-1 border-r">
            <div className="h-full overflow-auto p-4">
              <Card className="h-full">
                <CardHeader>
                  <CardTitle className="flex items-center gap-3 flex-wrap">
                    Document Preview
                    {currentFile && (
                      <span className="text-sm font-normal text-muted-foreground">
                        ({currentFile.file.name})
                      </span>
                    )}
                    {currentFile && (
                      <span className="flex items-center gap-3 text-xs font-normal ml-2">
                        <span className="flex items-center gap-1">
                          <span
                            className="inline-block w-3 h-3 rounded"
                            style={{
                              background: "#fef08a",
                              border: "1px solid #eab308",
                            }}
                          />
                          Unmatched / Unknown
                        </span>
                        <span className="flex items-center gap-1">
                          <span
                            className="inline-block w-3 h-3 rounded"
                            style={{
                              background: "#bbf7d0",
                              border: "1px solid #22c55e",
                            }}
                          />
                          Matched / Known
                        </span>
                      </span>
                    )}
                  </CardTitle>
                </CardHeader>
                <CardContent className="h-[calc(100%-80px)]">
                  {isUploading ? (
                    <div className="flex h-full items-center justify-center">
                      <div className="text-center space-y-6">
                        <div className="w-20 h-20 mx-auto relative">
                          <div className="absolute inset-0 border-4 border-gray-200 rounded-full" />
                          <div className="absolute inset-0 border-4 border-blue-500 border-t-transparent rounded-full animate-spin" />
                          <div className="absolute inset-0 flex items-center justify-center">
                            <div className="bg-white rounded-lg p-2 shadow-sm animate-pulse">
                              <FileText className="h-8 w-8 text-blue-500" />
                            </div>
                          </div>
                        </div>
                        <div className="flex justify-center space-x-1">
                          {[0, 150, 300].map((d) => (
                            <div
                              key={d}
                              className="w-2 h-2 bg-blue-500 rounded-full animate-bounce"
                              style={{ animationDelay: `${d}ms` }}
                            />
                          ))}
                        </div>
                        <p className="text-sm text-gray-600">
                          Analyzing document structure and extracting keys...
                        </p>
                      </div>
                    </div>
                  ) : currentFile?.replacedFile ? (
                    <DocumentPreview
                      key={`replaced-${previewKey}-${selectedFileIndex}`}
                      file={currentFile.replacedFile}
                      matchedKeys={
                        currentFile.replacedMatchedKeys ??
                        currentFile.matchedKeys
                      }
                      unmatchedKeys={
                        currentFile.replacedUnmatchedKeys ??
                        currentFile.unmatchedKeys
                      }
                      isReplaced={true}
                    />
                  ) : currentFile ? (
                    <DocumentPreview
                      key={`original-${currentFile.file.name}`}
                      file={currentFile.file}
                      matchedKeys={currentFile.matchedKeys}
                      unmatchedKeys={currentFile.unmatchedKeys}
                      isReplaced={false}
                    />
                  ) : (
                    <div
                      className={`flex h-full items-center justify-center text-muted-foreground border-2 border-dashed rounded-lg transition-colors cursor-pointer ${isDragOver ? "border-primary bg-primary/5" : "border-muted-foreground/25"}`}
                      onDragOver={handleDragOver}
                      onDragLeave={handleDragLeave}
                      onDrop={handleDrop}
                      onClick={() => fileInputRef.current?.click()}
                    >
                      <div className="text-center">
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

          {/* Right Panel */}
          <div className="w-[520px] flex-shrink-0">
            <div className="h-full overflow-auto p-4">
              <Card className="h-full flex flex-col">
                {/* ── AVAILABLE PLACEHOLDERS VIEW ── */}
                {!showPlaceholderManager && (
                  <>
                    <CardHeader className="flex-shrink-0">
                      <CardTitle className="flex items-center justify-between flex-wrap gap-2">
                        <span>Available Placeholders</span>
                        <div className="flex items-center gap-2 flex-wrap">
                          <Badge className="text-xs bg-black text-white hover:bg-black/90">
                            {AVAILABLE_KEYS.length} total keys
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
                              className={`h-full rounded-full transition-all duration-500 ${allMapped ? "bg-green-500" : "bg-orange-400"}`}
                              style={{
                                width: `${(mappedCount / uniqueUnmatchedKeys.length) * 100}%`,
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
                              className={`h-full rounded-full transition-all duration-500 ${allMapped ? "bg-green-500" : "bg-orange-400"}`}
                              style={{
                                width: `${(mappedCount / uniqueUnmatchedKeys.length) * 100}%`,
                              }}
                            />
                          </div>
                        </div>
                      )}
                    </CardHeader>

                    <CardContent className="flex-1 overflow-y-auto space-y-6">
                      {/* Unmatched Keys section */}
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
                                  uniqueUnmatchedKeys.length &&
                                uniqueUnmatchedKeys.length > 0
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

                          {/* Individual unmatched keys */}
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
                                    className={`flex items-start gap-3 p-3 rounded-lg border transition-colors ${isMapped ? "bg-green-50 border-green-200" : "bg-gray-50 border-gray-200"}`}
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
                                          <button
                                            type="button"
                                            onClick={() => {
                                              navigator.clipboard.writeText(
                                                key,
                                              );
                                              toast("Copied", {
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
                                          placeholder={`e.g., <<${key.replace(/^<</, "").replace(/>>$/, "").replace(/\s+/g, "_").toUpperCase()}>>`}
                                          availableKeys={AVAILABLE_KEYS}
                                        />
                                      </div>
                                    </div>
                                  </div>
                                );
                              })}
                          </div>

                          <Button
                            onClick={handleRemoveSelected}
                            variant="destructive"
                            disabled={selectedUnmatched.length === 0}
                            className="w-full mt-1"
                          >
                            Remove Selected ({selectedUnmatched.length})
                          </Button>
                        </div>
                      )}

                      {/* Add Custom Mapping */}
                      <div className="border-t pt-4 space-y-3">
                        <div>
                          <h4 className="font-medium text-sm">
                            Add Custom Mapping:
                          </h4>
                          <p className="text-xs text-gray-500 mt-1">
                            Map any text in the document — with or without
                            spaces — to a placeholder. e.g.{" "}
                            <span className="font-mono">student name</span>,{" "}
                            <span className="font-mono">John Doe</span>,{" "}
                            <span className="font-mono">
                              &lt;&lt;FIRSTNAME&gt;&gt;
                            </span>
                          </p>
                        </div>
                        <div className="space-y-3">
                          <div>
                            <Label
                              htmlFor="new-key"
                              className="text-xs text-gray-500 mb-1 block"
                            >
                              Text found in document (exact, any format):
                            </Label>
                            <input
                              id="new-key"
                              type="text"
                              value={newPlaceholderKey}
                              onChange={(e) =>
                                setNewPlaceholderKey(e.target.value)
                              }
                              onKeyDown={(e) => {
                                if (e.key === "Enter") handleAddPlaceholder();
                              }}
                              placeholder="e.g.  student name  OR  <<FIRSTNAME>>  OR  @LASTNAME"
                              className="w-full border border-gray-200 rounded-md px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
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
                              placeholder="e.g., <<FIRSTNAME>>"
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

                      {/* Custom Mappings List */}
                      {Object.keys(customPlaceholders).filter(
                        (k) => !uniqueUnmatchedKeys.includes(k),
                      ).length > 0 && (
                        <div className="space-y-2 border-t pt-4">
                          <h4 className="font-medium text-sm">
                            Custom Mappings:
                            <Badge variant="outline" className="ml-2 text-xs">
                              {
                                Object.keys(customPlaceholders).filter(
                                  (k) => !uniqueUnmatchedKeys.includes(k),
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
                                        &quot;{key}&quot;
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

        {/* Bottom Bar */}
        <div className="border-t bg-card p-4 flex-shrink-0">
          <div className="flex justify-center gap-4 mb-2">
            <Button
              onClick={handleReplaceKeys}
              disabled={
                !uploadedFiles.length ||
                isProcessing ||
                isUploading ||
                (uniqueMatchedKeys.length === 0 &&
                  uniqueUnmatchedKeys.length === 0 &&
                  Object.keys(customPlaceholders).length === 0)
              }
              className="flex items-center gap-2"
            >
              {isProcessing
                ? `Processing file ${processingIndex + 1} of ${uploadedFiles.length}…`
                : `Replace Keys in ${uploadedFiles.length} File${uploadedFiles.length !== 1 ? "s" : ""}`}
            </Button>
            <Button
              onClick={handleDownload}
              disabled={!uploadedFiles.filter((f) => f.replacedFile).length}
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
          <div className="text-center text-sm mt-[20px]">
            Developed by Abdul Basit
          </div>
        </div>
      </div>
      {/* ── Replacement Animation Overlay ── */}
      <ReplacementAnimation
        isOpen={isProcessing}
        processingIndex={processingIndex}
        totalFiles={uploadedFiles.length}
        currentFileName={
          processingIndex >= 0
            ? uploadedFiles[processingIndex]?.file.name
            : undefined
        }
        matchedKeys={
          processingIndex >= 0
            ? (uploadedFiles[processingIndex]?.matchedKeys ?? [])
            : []
        }
        unmatchedKeys={
          processingIndex >= 0
            ? (uploadedFiles[processingIndex]?.unmatchedKeys ?? [])
            : []
        }
        customPlaceholders={customPlaceholders}
      />
    </>
  );
}
