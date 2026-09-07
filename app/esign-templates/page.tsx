"use client";
import { Toaster, toast } from "react-hot-toast";
import type React from "react";
import { useEffect, useRef, useState } from "react";
import Link from "next/link";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import {
  Upload,
  Download,
  X,
  FileCode2,
  FileText,
  Plus,
  ArrowLeft,
  CheckCircle2,
  AlertCircle,
  RotateCcw,
  Settings2,
  KeyRound,
  SquareCheck,
  TextCursorInput,
  Eye,
  ClipboardCopy,
  Image as ImageIcon,
  Shapes,
  Wand2,
  Ban,
  ChevronRight,
} from "lucide-react";
import { PlaceholderAutocomplete } from "@/components/Placeholderautocomplete";
import { EsignPreview } from "@/components/esign/esign-preview";
import { downloadFile } from "@/lib/download-helper";
import {
  analyzeEsignHtml,
  decodeHtmlBytes,
  transformEsignHtml,
  DEFAULT_TEMPLATES,
  ESIGN_KNOWN_KEYS,
  type EsignAnalysis,
  type EsignKeyOutput,
  type EsignReplaceStats,
  type EsignTemplates,
} from "@/lib/esign/html-template";
import {
  assignControlIds,
  type CandidateType,
  type DocxAnalysis,
  type DocxCandidate,
} from "@/lib/esign/docx-shared";
import { convertDocxToHtml } from "@/lib/esign/docx-to-html.client";

/* ───────────────────────────── types ───────────────────────────── */

type FileKind = "docx" | "html";

interface EsignFile {
  kind: FileKind;
  file: File;
  /** HTML shown in the preview (the .htm itself, or the .docx rendered to HTML) */
  html: string;
  charset: string;
  analysis: EsignAnalysis;
  /** docx only: candidates + the user's chosen type per candidate */
  candidates: DocxCandidate[];
  types: Record<string, CandidateType>;
  /** outputs after Build */
  replacedHtml?: string;
  replacedBytes?: Uint8Array;
  mappedDocxBytes?: Uint8Array;
  replacedAnalysis?: EsignAnalysis;
  stats?: EsignReplaceStats;
  /** Step 2 (docx only): the mapped .docx saved by Word as .htm and
   *  uploaded back, used instead of the docx-preview conversion */
  linkedHtml?: LinkedHtml;
  /** which converter produced replacedHtml */
  htmlSource?: "docx-preview" | "word";
}

interface LinkedHtml {
  file: File;
  html: string;
  charset: string;
  analysis: EsignAnalysis;
}

interface EsignSettings {
  keyOutput: EsignKeyOutput;
  replaceKeys: boolean;
  replaceControls: boolean;
  templates: EsignTemplates;
}


const DEFAULT_SETTINGS: EsignSettings = {
  // encoded is the only form a browser renders as <<KEY>>; a literal
  // <<KEY>> in HTML is parsed as a tag and shows up as "<>"
  keyOutput: "encoded",
  replaceKeys: true,
  replaceControls: true,
  templates: { ...DEFAULT_TEMPLATES },
};

const SETTINGS_STORAGE_KEY = "esign-template-settings-v2";

function kindOf(f: File): FileKind | null {
  const n = f.name.toLowerCase();
  if (n.endsWith(".docx")) return "docx";
  if (n.endsWith(".html") || n.endsWith(".htm")) return "html";
  return null;
}

function baseName(name: string): string {
  return name.replace(/\.(docx|html?|htm)$/i, "");
}

function loadSettings(): EsignSettings {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) return DEFAULT_SETTINGS;
    const parsed = JSON.parse(raw) as Partial<EsignSettings>;
    return {
      ...DEFAULT_SETTINGS,
      ...parsed,
      templates: { ...DEFAULT_TEMPLATES, ...(parsed.templates ?? {}) },
    };
  } catch {
    return DEFAULT_SETTINGS;
  }
}

function defaultTypes(candidates: DocxCandidate[]): Record<string, CandidateType> {
  const out: Record<string, CandidateType> = {};
  candidates.forEach((c) => {
    // candidates we could not classify stay undecided → shown in "Review"
    if (c.suggested) out[c.id] = c.suggested;
  });
  return out;
}

function parseStatsHeader(res: Response): Record<string, unknown> | undefined {
  try {
    const header = res.headers.get("X-Esign-Stats");
    return header ? JSON.parse(decodeURIComponent(header)) : undefined;
  } catch {
    return undefined;
  }
}

/* ───────────────────────────── page ───────────────────────────── */

export default function EsignTemplatesPage() {
  const [files, setFiles] = useState<EsignFile[]>([]);
  const [selectedIndex, setSelectedIndex] = useState(0);
  const [isUploading, setIsUploading] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingIndex, setProcessingIndex] = useState(-1);
  const [processingStep, setProcessingStep] = useState("");
  const [isDragOver, setIsDragOver] = useState(false);
  const [mappings, setMappings] = useState<Record<string, string>>({});
  /** raw tokens (mostly plain-text matches like "Student Name") the user
   *  decided to leave alone */
  const [excluded, setExcluded] = useState<string[]>([]);
  const [customText, setCustomText] = useState("");
  const [customValue, setCustomValue] = useState("");
  const [previewMode, setPreviewMode] = useState<"original" | "replaced">(
    "original",
  );
  const [showHighlights, setShowHighlights] = useState(true);
  const [settings, setSettings] = useState<EsignSettings>(DEFAULT_SETTINGS);
  const [settingsLoaded, setSettingsLoaded] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const wordHtmlInputRef = useRef<HTMLInputElement>(null);

  /* settings persistence */
  useEffect(() => {
    setSettings(loadSettings());
    setSettingsLoaded(true);
  }, []);
  useEffect(() => {
    if (!settingsLoaded) return;
    try {
      localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
    } catch {
      /* ignore */
    }
  }, [settings, settingsLoaded]);

  /* make sure every unmatched token has a mapping slot */
  useEffect(() => {
    const unmatched = [...new Set(files.flatMap((f) => f.analysis.unmatchedKeys))];
    setMappings((prev) => {
      let changed = false;
      const next = { ...prev };
      unmatched.forEach((k) => {
        if (next[k] === undefined) {
          next[k] = "";
          changed = true;
        }
      });
      return changed ? next : prev;
    });
  }, [files]);

  /* ── derived ── */
  const current = files[selectedIndex];
  const uniqueMatched = [...new Set(files.flatMap((f) => f.analysis.matchedKeys))];
  const uniqueUnmatched = [
    ...new Set(files.flatMap((f) => f.analysis.unmatchedKeys)),
  ];
  const mappedCount = uniqueUnmatched.filter((k) => mappings[k]).length;
  const plainTokens = (() => {
    const seen = new Map<string, { raw: string; knownKey: string; count: number }>();
    files.forEach((f) =>
      f.analysis.tokens
        .filter((t) => t.form === "plain" && t.knownKey)
        .forEach((t) => {
          const e = seen.get(t.raw);
          if (e) e.count += t.count;
          else seen.set(t.raw, { raw: t.raw, knownKey: t.knownKey as string, count: t.count });
        }),
    );
    return Array.from(seen.values());
  })();
  const customEntries = Object.entries(mappings).filter(
    ([k, v]) => !uniqueUnmatched.includes(k) && v,
  );
  const totalControls = files.reduce(
    (s, f) =>
      s +
      f.analysis.controls.length +
      f.candidates.filter((c) => (f.types[c.id] ?? "ignore") !== "ignore").length,
    0,
  );
  const needsReview = files.reduce(
    (s, f) =>
      s +
      f.candidates.filter((c) => c.suggested === null && f.types[c.id] === undefined)
        .length,
    0,
  );
  const processedFiles = files.filter((f) => f.replacedHtml !== undefined);
  /** review list for the simple checklist */
  const currentUndecided = current
    ? current.candidates.filter(
        (c) => c.suggested === null && current.types[c.id] === undefined,
      )
    : [];
  const unmappedKeys = uniqueUnmatched.filter((k) => !mappings[k]);
  const reviewCount = needsReview + unmappedKeys.length;
  const activeMappings = Object.fromEntries(
    Object.entries(mappings).filter(([, v]) => v),
  );

  /* ── upload ── */
  const processFiles = async (incoming: File[]) => {
    setIsUploading(true);
    const toastId = toast.loading(
      `Analyzing ${incoming.length} file${incoming.length > 1 ? "s" : ""}…`,
    );
    const added: EsignFile[] = [];

    for (const file of incoming) {
      const kind = kindOf(file);
      if (!kind) continue;
      try {
        const fd = new FormData();
        fd.append("file", file);

        if (kind === "docx") {
          const res = await fetch("/api/esign/docx/analyze", {
            method: "POST",
            body: fd,
          });
          if (!res.ok) throw new Error(`analyze failed (${res.status})`);
          const analysis = (await res.json()) as DocxAnalysis;
          let html: string;
          try {
            html = await convertDocxToHtml(file, file.name);
          } catch (convErr) {
            console.error(`docx-preview could not render ${file.name}:`, convErr);
            throw new Error(
              `${file.name} could not be rendered in the browser (${(convErr as Error).message})`,
            );
          }
          added.push({
            kind,
            file,
            html,
            charset: "docx",
            analysis,
            candidates: analysis.candidates,
            types: defaultTypes(analysis.candidates),
          });
        } else {
          const bytes = await file.arrayBuffer();
          const { html, charset } = decodeHtmlBytes(bytes);
          const res = await fetch("/api/esign/analyze", {
            method: "POST",
            body: fd,
          });
          if (!res.ok) throw new Error(`analyze failed (${res.status})`);
          const analysis = (await res.json()) as EsignAnalysis;
          added.push({
            kind,
            file,
            html,
            charset,
            analysis,
            candidates: [],
            types: {},
          });
        }
      } catch (err) {
        console.error(`Error analyzing ${file.name}:`, err);
        toast.error(
          (err as Error)?.message?.includes("rendered")
            ? (err as Error).message
            : `Failed to analyze ${file.name}`,
          { duration: 6000 },
        );
      }
    }

    setFiles((prev) => {
      // select the first newly added file when nothing was loaded before
      if (prev.length === 0 && added.length > 0) setSelectedIndex(0);
      return [...prev, ...added];
    });
    setPreviewMode("original");
    setIsUploading(false);
    toast.dismiss(toastId);

    if (added.length) {
      const keys = added.reduce((s, f) => s + f.analysis.matchedKeys.length, 0);
      const need = added.reduce(
        (s, f) => s + f.analysis.unmatchedKeys.length,
        0,
      );
      const cands = added.reduce((s, f) => s + f.candidates.length, 0);
      const markers = added.reduce((s, f) => s + f.analysis.controls.length, 0);
      const review = added.reduce(
        (s, f) => s + f.candidates.filter((c) => c.suggested === null).length,
        0,
      );
      toast(
        `${added.length} file${added.length > 1 ? "s" : ""} loaded · ${keys} keys · ${cands} checkbox/textbox candidates · ${markers} markers${need ? ` · ${need} keys need mapping` : ""}${review ? ` · ${review} candidates need review` : ""}`,
        { icon: need || review ? "⚠️" : "✅", duration: 5000 },
      );
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files) return;
    const supported = Array.from(e.target.files).filter((f) => kindOf(f));
    if (!supported.length) {
      toast.error("Please upload .docx, .html or .htm files");
      return;
    }
    await processFiles(supported);
    e.target.value = "";
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const supported = Array.from(e.dataTransfer.files).filter((f) => kindOf(f));
    if (!supported.length) {
      toast.error("Please upload .docx, .html or .htm files");
      return;
    }
    processFiles(supported);
  };

  const removeFile = (index: number) => {
    const name = files[index].file.name;
    setFiles((prev) => prev.filter((_, i) => i !== index));
    // keep the same document selected (or the nearest one) after removal
    setSelectedIndex((sel) => {
      if (index < sel) return sel - 1;
      if (index === sel) return Math.max(0, Math.min(sel, files.length - 2));
      return sel;
    });
    toast(`Removed ${name}`, { icon: "🗑️", duration: 2000 });
  };

  const handleReset = () => {
    setFiles([]);
    setSelectedIndex(0);
    setMappings({});
    setExcluded([]);
    setCustomText("");
    setCustomValue("");
    setPreviewMode("original");
    toast("All data reset", { icon: "🔄", duration: 2000 });
  };

  /* ── candidate types ── */
  const setCandidateType = (fileIndex: number, candidateId: string, type: CandidateType) =>
    setFiles((prev) =>
      prev.map((f, i) =>
        i === fileIndex ? { ...f, types: { ...f.types, [candidateId]: type } } : f,
      ),
    );

  const setAllCandidateTypes = (
    fileIndex: number,
    mode: "suggested" | "ignore",
  ) =>
    setFiles((prev) =>
      prev.map((f, i) => {
        if (i !== fileIndex) return f;
        const types: Record<string, CandidateType> = {};
        f.candidates.forEach((c) => {
          types[c.id] = mode === "ignore" ? "ignore" : (c.suggested ?? "ignore");
        });
        return { ...f, types };
      }),
    );

  /* ── custom mappings ── */
  const handleAddCustom = () => {
    const key = customText.trim();
    const val = customValue.trim();
    if (!key || !val) {
      toast.error("Both text and placeholder value are required");
      return;
    }
    if (mappings[key]) {
      toast.error(`Mapping for "${key}" already exists`);
      return;
    }
    const formatted = /^(<<|«)/.test(val) ? val : `<<${val.toUpperCase()}>>`;
    setMappings((prev) => ({ ...prev, [key]: formatted }));
    setCustomText("");
    setCustomValue("");
    toast.success(`Mapping added: "${key}" → ${formatted}`);
  };

  const handleRemoveMapping = (key: string) => {
    setMappings((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  };

  /* ── Step 2: Word-saved .htm linked to a docx ── */
  const analyzeHtmlFile = async (file: File): Promise<LinkedHtml> => {
    const bytes = await file.arrayBuffer();
    const { html, charset } = decodeHtmlBytes(bytes);
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch("/api/esign/analyze", { method: "POST", body: fd });
    if (!res.ok) throw new Error(`analyze failed (${res.status})`);
    const analysis = (await res.json()) as EsignAnalysis;
    return { file, html, charset, analysis };
  };

  const handleWordHtmlUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file || !current || current.kind !== "docx") return;
    if (kindOf(file) !== "html") {
      toast.error("Please upload the .htm / .html that Word saved");
      return;
    }
    const index = selectedIndex;
    const toastId = toast.loading(`Reading ${file.name}…`);
    try {
      const linked = await analyzeHtmlFile(file);
      toast.dismiss(toastId);
      if (linked.analysis.controls.length === 0 && current.candidates.length > 0) {
        toast(
          `${file.name} has no <c1>/<t1> markers — did you save the ORIGINAL file? Download the mapped .docx first, then save that one as .htm.`,
          { icon: "⚠️", duration: 8000 },
        );
      }
      let update: Partial<EsignFile> = { linkedHtml: linked };
      // already built → rebuild the HTML half right away from the Word file
      if (current.mappedDocxBytes) {
        const built = await buildHtmlFile(linked);
        update = {
          ...update,
          replacedHtml: built.replacedHtml,
          replacedBytes: built.replacedBytes,
          replacedAnalysis: analyzeEsignHtml(built.replacedHtml),
          htmlSource: "word",
          stats: {
            keys: current.stats?.keys ?? built.stats?.keys ?? {},
            controls: built.stats?.controls ?? {},
            skipped: current.stats?.skipped ?? [],
          },
        };
        setPreviewMode("replaced");
        toast.success(`Word HTML linked and built: ${file.name}`);
      } else {
        toast.success(`Word HTML linked: ${file.name}. It will be used when you build.`);
      }
      setFiles((prev) => prev.map((f, i) => (i === index ? { ...f, ...update } : f)));
    } catch (err) {
      console.error(err);
      toast.dismiss(toastId);
      toast.error(`Failed to read ${file.name}`);
    }
  };

  const handleUnlinkWordHtml = () =>
    setFiles((prev) =>
      prev.map((f, i) =>
        i === selectedIndex
          ? { ...f, linkedHtml: undefined, htmlSource: f.htmlSource === "word" ? undefined : f.htmlSource, replacedHtml: f.htmlSource === "word" ? undefined : f.replacedHtml, replacedBytes: f.htmlSource === "word" ? undefined : f.replacedBytes }
          : f,
      ),
    );

  const handleDownloadMappedDocx = () => {
    if (!current?.mappedDocxBytes) return;
    downloadFile(
      new Blob([current.mappedDocxBytes as Uint8Array<ArrayBuffer>], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }),
      `${baseName(current.file.name)}.docx`,
    );
    toast.success(`Downloaded ${baseName(current.file.name)}.docx — open it in Word and Save As → Web Page, Filtered (.htm)`, { duration: 6000 });
  };

  /* ── build ── */
  const buildHtmlFile = async (item: { file: File }) => {
    const fd = new FormData();
    fd.append("file", item.file);
    fd.append("mappings", JSON.stringify(activeMappings));
    fd.append(
      "options",
      JSON.stringify({
        keyOutput: settings.keyOutput,
        replaceKeys: settings.replaceKeys,
        replaceControls: settings.replaceControls,
        templates: settings.templates,
        exclude: excluded,
      }),
    );
    const res = await fetch("/api/esign/replace", { method: "POST", body: fd });
    if (!res.ok) throw new Error(`Failed to process ${item.file.name}`);
    const replacedBytes = new Uint8Array(await res.arrayBuffer());
    const replacedHtml = decodeHtmlBytes(replacedBytes).html;
    return {
      replacedHtml,
      replacedBytes,
      stats: parseStatsHeader(res) as EsignReplaceStats | undefined,
    };
  };

  const buildDocxFile = async (item: EsignFile) => {
    // 1. server: candidates → <c1>/<t1> markers, keys → <<KEY>>
    setProcessingStep("mapping checkboxes, textboxes and keys in Word");
    const assignments = assignControlIds(item.candidates, item.types, item.analysis);
    const fd = new FormData();
    fd.append("file", item.file);
    fd.append("assignments", JSON.stringify(assignments));
    fd.append("mappings", JSON.stringify(activeMappings));
    fd.append(
      "options",
      JSON.stringify({ replaceKeys: settings.replaceKeys, exclude: excluded }),
    );
    const res = await fetch("/api/esign/docx/build", { method: "POST", body: fd });
    if (!res.ok) throw new Error(`Failed to build ${item.file.name}`);
    const mappedDocxBytes = new Uint8Array(await res.arrayBuffer());
    const docxStats = parseStatsHeader(res) as
      | { keys: Record<string, number>; skipped: string[] }
      | undefined;

    // 2. browser: mapped .docx → HTML (docx-preview, high fidelity)
    setProcessingStep("converting the mapped Word file to HTML");
    const html = await convertDocxToHtml(mappedDocxBytes, baseName(item.file.name));

    // 3. HTML engine: markers → form controls, key format
    setProcessingStep("turning markers into form controls");
    const htmlAnalysis = analyzeEsignHtml(html);
    const { html: replacedHtml, stats: htmlStats } = transformEsignHtml(
      html,
      htmlAnalysis,
      {
        mappings: activeMappings,
        exclude: excluded,
        keyOutput: settings.keyOutput,
        replaceKeys: settings.replaceKeys,
        replaceControls: settings.replaceControls,
        templates: settings.templates,
      },
    );
    const stats: EsignReplaceStats = {
      keys: docxStats?.keys ?? htmlStats.keys,
      controls: htmlStats.controls,
      skipped: docxStats?.skipped ?? htmlStats.skipped,
    };

    // Step 2: a Word-saved .htm of the mapped file is linked → prefer it
    if (item.linkedHtml) {
      setProcessingStep("building the Word-saved HTML");
      const word = await buildHtmlFile(item.linkedHtml);
      return {
        replacedHtml: word.replacedHtml,
        replacedBytes: word.replacedBytes,
        mappedDocxBytes,
        htmlSource: "word" as const,
        stats: {
          keys: stats.keys,
          controls: word.stats?.controls ?? stats.controls,
          skipped: stats.skipped,
        },
      };
    }

    return {
      replacedHtml,
      replacedBytes: new TextEncoder().encode(replacedHtml),
      mappedDocxBytes,
      htmlSource: "docx-preview" as const,
      stats,
    };
  };

  const handleBuild = async (): Promise<EsignFile[] | null> => {
    if (!files.length) return null;
    const missing = uniqueUnmatched.filter((k) => !mappings[k]);
    if (missing.length) {
      toast(
        `${missing.length} placeholder${missing.length > 1 ? "s" : ""} not in the key list — kept as they are`,
        { icon: "ℹ️", duration: 4000 },
      );
    }
    if (needsReview) {
      toast(
        `${needsReview} item${needsReview > 1 ? "s" : ""} without a decision — left as they are`,
        { icon: "ℹ️", duration: 4000 },
      );
    }

    setIsProcessing(true);
    const toastId = toast.loading("Building…");
    try {
      const updated = [...files];
      for (let i = 0; i < updated.length; i++) {
        setProcessingIndex(i);
        const item = updated[i];
        const result =
          item.kind === "docx" ? await buildDocxFile(item) : await buildHtmlFile(item);
        updated[i] = {
          ...item,
          ...result,
          replacedAnalysis: analyzeEsignHtml(result.replacedHtml),
        };
      }
      setFiles(updated);
      setPreviewMode("replaced");
      toast.dismiss(toastId);

      const keyHits = updated.reduce(
        (s, f) =>
          s + Object.values(f.stats?.keys ?? {}).reduce((a, b) => a + b, 0),
        0,
      );
      const ctrlHits = updated.reduce(
        (s, f) =>
          s +
          Object.values(f.stats?.controls ?? {}).reduce((a, b) => a + b, 0),
        0,
      );
      toast.success(
        `${updated.length} template${updated.length > 1 ? "s" : ""} built · ${keyHits} keys · ${ctrlHits} controls`,
        { duration: 4000 },
      );
      return updated;
    } catch (err) {
      console.error(err);
      toast.dismiss(toastId);
      toast.error("Error building templates. Please try again.");
      return null;
    } finally {
      setIsProcessing(false);
      setProcessingIndex(-1);
      setProcessingStep("");
    }
  };

  /* ── download ── */
  const downloadFiles = async (list: EsignFile[]) => {
    const processedFiles = list.filter((f) => f.replacedHtml !== undefined);
    if (!processedFiles.length) {
      toast.error("Please build the templates first");
      return;
    }
    try {
      const single = processedFiles.length === 1 ? processedFiles[0] : null;
      if (single && single.kind === "html") {
        downloadFile(
          new Blob([single.replacedBytes as Uint8Array<ArrayBuffer>], {
            type: "text/html",
          }),
          single.file.name,
        );
        toast.success(`Downloaded ${single.file.name}`);
        return;
      }

      const toastId = toast.loading("Zipping templates…");
      const JSZip = (await import("jszip")).default;
      const zip = new JSZip();
      let count = 0;
      processedFiles.forEach((f) => {
        const base = baseName(f.file.name);
        if (f.kind === "docx") {
          zip.file(`${base}.docx`, f.mappedDocxBytes as Uint8Array);
          zip.file(`${base}.htm`, f.replacedBytes as Uint8Array);
          count += 2;
        } else {
          zip.file(f.file.name, f.replacedBytes as Uint8Array);
          count += 1;
        }
      });
      const blob = await zip.generateAsync({ type: "blob" });
      const zipName = single ? `${baseName(single.file.name)}_esign.zip` : "esign_templates.zip";
      downloadFile(blob, zipName);
      toast.dismiss(toastId);
      toast.success(`Downloaded ${count} files as ${zipName}`);
    } catch (err) {
      console.error(err);
      toast.error("Error downloading. Please try again.");
    }
  };

  /** one click for ordinary users: build everything, then download */
  const handleBuildAndDownload = async () => {
    const built = await handleBuild();
    if (built) await downloadFiles(built);
  };

  const previewHtml =
    previewMode === "replaced" && current?.replacedHtml !== undefined
      ? current.replacedHtml
      : current?.html;

  /* ───────────────────────────── render ───────────────────────────── */
  return (
    <>
      <Toaster
        position="bottom-right"
        toastOptions={{ style: { fontSize: "13px" } }}
      />

      <div className="min-h-screen bg-background flex flex-col">
        {/* Top bar */}
        <div className="border-b bg-card p-4 flex-shrink-0">
          <div className="flex items-center justify-between gap-4 flex-wrap">
            <h1 className="text-2xl font-bold flex items-center gap-3 flex-wrap">
              <Link
                href="/"
                className="text-sm font-normal text-gray-500 hover:text-gray-800 flex items-center gap-1"
              >
                <ArrowLeft className="h-4 w-4" />
                DocX Key Replacer
              </Link>
              <span>
                Esign Templates{" "}
                <span className="text-sm text-gray-500 font-normal">
                  / Customized — Word file in, esign-ready files out
                </span>
              </span>
              {files.length > 0 && (
                <Badge variant="secondary" className="text-xs">
                  {files.length} file{files.length !== 1 ? "s" : ""}
                </Badge>
              )}
            </h1>

            <div className="flex items-center gap-2">
              <Button
                variant="outline"
                className="flex items-center gap-2"
                onClick={handleReset}
                disabled={!files.length && !Object.keys(mappings).length}
              >
                <RotateCcw className="h-4 w-4" />
                Reset
              </Button>
                            <Button
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-2"
              >
                <Upload className="h-4 w-4" />
                Upload .docx / .htm
              </Button>
            </div>
            <input
              ref={fileInputRef}
              type="file"
              accept=".docx,.html,.htm,text/html,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              multiple
              onChange={handleFileUpload}
              className="hidden"
            />
            <input
              ref={wordHtmlInputRef}
              type="file"
              accept=".htm,.html,text/html"
              onChange={handleWordHtmlUpload}
              className="hidden"
            />
          </div>
        </div>

        {/* File list */}
        {files.length > 0 && (
          <div className="border-b bg-muted/30 p-2 flex-shrink-0">
            <div className="flex items-center gap-2 overflow-x-auto">
              <span className="text-sm font-medium text-muted-foreground whitespace-nowrap">
                Files:
              </span>
              {files.map((f, index) => (
                <div
                  key={`${f.file.name}-${index}`}
                  title={f.file.name}
                  className={`flex items-center gap-2 px-3 py-1 rounded-md border cursor-pointer transition-colors ${selectedIndex === index ? "bg-primary text-primary-foreground" : "bg-background hover:bg-muted"}`}
                  onClick={() => setSelectedIndex(index)}
                >
                  {f.kind === "docx" ? (
                    <FileText className="h-3 w-3 flex-shrink-0" />
                  ) : (
                    <FileCode2 className="h-3 w-3 flex-shrink-0" />
                  )}
                  <span className="text-xs truncate max-w-40">{f.file.name}</span>
                  {isProcessing && processingIndex === index && (
                    <span className="text-xs animate-spin">⚙️</span>
                  )}
                  {f.replacedHtml !== undefined && !isProcessing && (
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

        {/* Main */}
        <div className="flex h-[calc(100vh-180px)]">
          {/* Preview */}
          <div className="flex-1 border-r min-w-0">
            <div className="h-full overflow-hidden p-4">
              <Card className="h-full flex flex-col">
                <CardHeader className="flex-shrink-0 pb-2">
                  <CardTitle className="flex items-center gap-3 flex-wrap text-base">
                    Template Preview
                    {current && (
                      <span className="text-sm font-normal text-muted-foreground truncate max-w-xs">
                        ({current.file.name}
                        {current.kind === "html" ? ` · ${current.charset}` : " · Word"}
                        {previewMode === "replaced" && current.htmlSource === "word"
                          ? " · Word HTML"
                          : previewMode === "replaced" && current.htmlSource === "docx-preview"
                            ? " · docx-preview HTML"
                            : ""}
                        )
                      </span>
                    )}
                    {current && (
                      <span className="flex items-center gap-1 ml-auto">
                        <Button
                          size="sm"
                          variant={previewMode === "original" ? "default" : "outline"}
                          className="h-7 text-xs"
                          onClick={() => setPreviewMode("original")}
                        >
                          Original
                        </Button>
                        <Button
                          size="sm"
                          variant={previewMode === "replaced" ? "default" : "outline"}
                          className="h-7 text-xs"
                          disabled={current.replacedHtml === undefined}
                          onClick={() => setPreviewMode("replaced")}
                        >
                          Built HTML
                        </Button>
                        <Button
                          size="sm"
                          variant={showHighlights ? "secondary" : "outline"}
                          className="h-7 text-xs flex items-center gap-1"
                          onClick={() => setShowHighlights((v) => !v)}
                          title="Toggle highlights"
                        >
                          <Eye className="h-3 w-3" />
                          {showHighlights ? "Highlights on" : "Highlights off"}
                        </Button>
                                              </span>
                    )}
                  </CardTitle>
                  {current && (
                    <div className="flex items-center gap-3 text-xs font-normal flex-wrap mt-1">
                      <LegendSwatch bg="#bbf7d0" border="#22c55e" label="Key" />
                      <LegendSwatch bg="#fef08a" border="#eab308" label="Needs a decision" />
                      <LegendSwatch bg="#e0f2fe" border="#0ea5e9" label="Checkbox" />
                      <LegendSwatch bg="#ede9fe" border="#8b5cf6" label="Textbox" />
                    </div>
                  )}
                </CardHeader>
                <CardContent className="flex-1 min-h-0 pt-0">
                  {isUploading || isProcessing ? (
                    <div className="flex h-full items-center justify-center">
                      <div className="text-center space-y-4">
                        <div className="w-16 h-16 mx-auto border-4 border-gray-200 border-t-blue-500 rounded-full animate-spin" />
                        <p className="text-sm text-gray-600">
                          {isUploading
                            ? "Reading file and detecting keys, checkboxes and textboxes…"
                            : `Building ${files[processingIndex]?.file.name ?? ""}${processingStep ? ` — ${processingStep}` : ""}…`}
                        </p>
                      </div>
                    </div>
                  ) : current && previewHtml !== undefined ? (
                    <div className="h-full border rounded-md overflow-hidden bg-white">
                      <EsignPreview
                        key={`${previewMode}-${selectedIndex}`}
                        html={previewHtml}
                        mappings={activeMappings}
                        exclude={excluded}
                        highlight={showHighlights}
                      />
                    </div>
                  ) : (
                    <div
                      className={`flex h-full items-center justify-center text-muted-foreground border-2 border-dashed rounded-lg transition-colors cursor-pointer ${isDragOver ? "border-primary bg-primary/5" : "border-muted-foreground/25"}`}
                      onDragOver={(e) => {
                        e.preventDefault();
                        setIsDragOver(true);
                      }}
                      onDragLeave={(e) => {
                        e.preventDefault();
                        setIsDragOver(false);
                      }}
                      onDrop={handleDrop}
                      onClick={() => fileInputRef.current?.click()}
                    >
                      <div className="text-center max-w-lg px-6 space-y-3">
                        <Upload
                          className={`mx-auto h-12 w-12 ${isDragOver ? "text-primary" : ""}`}
                        />
                        <p className="text-lg">
                          {isDragOver
                            ? "Drop the client's Word file here"
                            : "Upload the client's .docx to start the esign process"}
                        </p>
                        <p className="text-sm text-muted-foreground">
                          Checkboxes and textboxes in the Word file (content controls,
                          form fields, box symbols, <span className="font-mono">[ ]</span>,
                          underscores, pictures, shapes) are detected and mapped to{" "}
                          <span className="font-mono">&lt;c1&gt;</span> /{" "}
                          <span className="font-mono">&lt;c1c&gt;</span> /{" "}
                          <span className="font-mono">&lt;t1&gt;</span>, keys are replaced,
                          then the file is converted to HTML with real form controls. You
                          get both the mapped .docx and the .htm.
                        </p>
                        <p className="text-xs text-muted-foreground">
                          Already have an .htm with markers? Upload it directly to only do
                          the HTML step.
                        </p>
                      </div>
                    </div>
                  )}
                </CardContent>
              </Card>
            </div>
          </div>

          {/* Right panel — simple 3-step checklist */}
          <div className="w-[520px] flex-shrink-0">
            <div className="h-full overflow-hidden p-4">
              <Card className="h-full flex flex-col">
                <CardContent className="flex-1 overflow-y-auto space-y-4 pt-5">
                  {/* ── Step 1: upload ── */}
                  <StepCard
                    n={1}
                    title="Upload the client's Word file"
                    done={files.length > 0}
                    active={files.length === 0}
                  >
                    {files.length === 0 ? (
                      <p className="text-sm text-gray-600">
                        Drop the .docx on the left or click <strong>Upload</strong>. An .htm
                        that already contains <span className="font-mono">&lt;c1&gt;</span>{" "}
                        markers also works.
                      </p>
                    ) : (
                      <p className="text-sm text-gray-600">
                        <span className="font-medium">{current?.file.name}</span>
                        {current && (
                          <>
                            {" "}
                            · {current.candidates.filter((c) => (current.types[c.id] ?? c.suggested ?? "ignore") !== "ignore").length + current.analysis.controls.length}{" "}
                            checkboxes / textboxes · {current.analysis.matchedKeys.length} keys
                            recognised
                          </>
                        )}
                        {files.length > 1 && ` · ${files.length} files in total`}
                      </p>
                    )}
                  </StepCard>

                  {/* ── Step 2: review ── */}
                  <StepCard
                    n={2}
                    title="Review"
                    done={files.length > 0 && reviewCount === 0}
                    active={files.length > 0 && reviewCount > 0}
                    badge={reviewCount || undefined}
                  >
                    {!current && (
                      <p className="text-sm text-gray-500">Waiting for a file…</p>
                    )}

                    {current && reviewCount === 0 && (
                      <p className="text-sm text-green-700 flex items-center gap-2">
                        <CheckCircle2 className="h-4 w-4" />
                        Nothing to review — everything was recognised automatically.
                      </p>
                    )}

                    {current && currentUndecided.length > 0 && (
                      <div className="space-y-2">
                        <p className="text-sm text-gray-700">
                          What is this? ({currentUndecided.length} item
                          {currentUndecided.length > 1 ? "s" : ""} we could not tell)
                        </p>
                        {currentUndecided.map((c, i) => (
                          <DecisionCard
                            key={c.id}
                            candidate={c}
                            index={current.candidates.indexOf(c) + 1}
                            onChoose={(type) => setCandidateType(selectedIndex, c.id, type)}
                          />
                        ))}
                      </div>
                    )}

                    {unmappedKeys.length > 0 && (
                      <div className="space-y-2">
                        <p className="text-sm text-gray-700">
                          These placeholders are not in the key list — pick the right key,
                          or leave empty to keep the text as it is.
                        </p>
                        {unmappedKeys.map((key) => (
                          <div
                            key={key}
                            className="p-2.5 rounded-lg border bg-yellow-50 border-yellow-200 space-y-1.5"
                          >
                            <div className="text-sm font-mono break-all">{key}</div>
                            <PlaceholderAutocomplete
                              id={`map-${key}`}
                              value={mappings[key] || ""}
                              onChange={(val) =>
                                setMappings((prev) => ({ ...prev, [key]: val }))
                              }
                              placeholder="Type to search the key list…"
                              availableKeys={ESIGN_KNOWN_KEYS}
                            />
                          </div>
                        ))}
                      </div>
                    )}

                    {current && (
                      <details className="group border rounded-lg">
                        <summary className="cursor-pointer select-none px-3 py-2 text-sm text-gray-700 flex items-center gap-2">
                          <ChevronRight className="h-4 w-4 transition-transform group-open:rotate-90" />
                          Show everything detected
                          <span className="text-xs text-gray-500">
                            ({current.candidates.length + current.analysis.controls.length}{" "}
                            controls · {current.analysis.matchedKeys.length + uniqueUnmatched.length} keys)
                          </span>
                        </summary>
                        <div className="px-3 pb-3 space-y-5 border-t pt-3">
                          <ControlsPanel
                            file={current}
                            onTypeChange={(id, type) =>
                              setCandidateType(selectedIndex, id, type)
                            }
                            onBulk={(mode) => setAllCandidateTypes(selectedIndex, mode)}
                          />

                          {uniqueMatched.length > 0 && (
                            <div className="space-y-2">
                              <h4 className="font-medium text-sm">
                                Keys recognised ({uniqueMatched.length})
                              </h4>
                              <div className="flex flex-wrap gap-1.5">
                                {uniqueMatched.map((k) => (
                                  <span
                                    key={k}
                                    className="text-xs font-mono px-2 py-0.5 rounded border border-green-200 bg-green-50 text-green-800"
                                  >
                                    &lt;&lt;{k}&gt;&gt;
                                  </span>
                                ))}
                              </div>
                            </div>
                          )}

                          {plainTokens.length > 0 && (
                            <div className="space-y-2">
                              <div>
                                <h4 className="font-medium text-sm">
                                  Key names written as plain text ({plainTokens.length})
                                </h4>
                                <p className="text-xs text-gray-500">
                                  These will be replaced too. Untick any that is really a
                                  label and must stay as it is.
                                </p>
                              </div>
                              <div className="grid gap-1.5">
                                {plainTokens.map((t) => {
                                  const on = !excluded.includes(t.raw);
                                  return (
                                    <label
                                      key={t.raw}
                                      className={`flex items-center gap-2 p-2 rounded-lg border text-sm cursor-pointer ${on ? "bg-orange-50 border-orange-200" : "bg-gray-50 border-gray-200 text-gray-500"}`}
                                    >
                                      <input
                                        type="checkbox"
                                        checked={on}
                                        onChange={(e) =>
                                          setExcluded((prev) =>
                                            e.target.checked
                                              ? prev.filter((x) => x !== t.raw)
                                              : [...prev, t.raw],
                                          )
                                        }
                                        className="h-4 w-4"
                                      />
                                      <span className={`font-mono ${on ? "" : "line-through"}`}>
                                        {t.raw}
                                      </span>
                                      <span className="text-xs text-gray-500">×{t.count}</span>
                                      <span className="ml-auto font-mono text-xs text-green-800">
                                        → &lt;&lt;{t.knownKey}&gt;&gt;
                                      </span>
                                    </label>
                                  );
                                })}
                              </div>
                            </div>
                          )}

                          {/* custom mapping */}
                          <div className="border-t pt-3 space-y-2">
                            <h4 className="font-medium text-sm">Map some other text to a key</h4>
                            <p className="text-xs text-gray-500">
                              e.g. <span className="font-mono">student name</span> →{" "}
                              <span className="font-mono">&lt;&lt;STUDENTNAME&gt;&gt;</span>
                            </p>
                            <Input
                              value={customText}
                              onChange={(e) => setCustomText(e.target.value)}
                              onKeyDown={(e) => {
                                if (e.key === "Enter") handleAddCustom();
                              }}
                              placeholder="Text found in the document"
                              className="font-mono text-sm"
                            />
                            <PlaceholderAutocomplete
                              id="custom-value"
                              value={customValue}
                              onChange={setCustomValue}
                              placeholder="Key to use"
                              availableKeys={ESIGN_KNOWN_KEYS}
                            />
                            <Button
                              onClick={handleAddCustom}
                              variant="outline"
                              size="sm"
                              className="flex items-center gap-2"
                            >
                              <Plus className="h-4 w-4" />
                              Add
                            </Button>
                            {customEntries.length > 0 && (
                              <div className="grid gap-1.5 pt-1">
                                {customEntries.map(([key, value]) => (
                                  <div
                                    key={key}
                                    className="flex items-center gap-2 p-2 bg-gray-50 rounded-lg border border-gray-200 text-sm font-mono"
                                  >
                                    <span className="text-blue-600 truncate">&quot;{key}&quot;</span>
                                    <span className="text-gray-500">→</span>
                                    <span className="truncate">{value}</span>
                                    <Button
                                      variant="ghost"
                                      size="sm"
                                      onClick={() => handleRemoveMapping(key)}
                                      className="h-7 w-7 p-0 ml-auto hover:bg-red-100 hover:text-red-600"
                                    >
                                      <X className="h-4 w-4" />
                                    </Button>
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        </div>
                      </details>
                    )}
                  </StepCard>

                  {/* ── Step 3: build & download ── */}
                  <StepCard
                    n={3}
                    title="Build & download"
                    done={processedFiles.length > 0 && processedFiles.length === files.length}
                    active={files.length > 0 && reviewCount === 0 && processedFiles.length === 0}
                  >
                    <div className="flex items-center gap-2 flex-wrap">
                      <Button
                        onClick={handleBuildAndDownload}
                        disabled={!files.length || isProcessing || isUploading}
                        className="flex items-center gap-2"
                      >
                        <Download className="h-4 w-4" />
                        {isProcessing
                          ? `Building ${processingIndex + 1} of ${files.length}…`
                          : processedFiles.length
                            ? "Build again & download"
                            : `Build & download${files.length > 1 ? ` (${files.length} files)` : ""}`}
                      </Button>
                      {processedFiles.length > 0 && !isProcessing && (
                        <Button
                          variant="outline"
                          onClick={() => downloadFiles(processedFiles)}
                          className="flex items-center gap-2"
                        >
                          <Download className="h-4 w-4" />
                          Download again
                        </Button>
                      )}
                    </div>
                    {files.length > 0 && reviewCount > 0 && !processedFiles.length && (
                      <p className="text-xs text-gray-500">
                        You can build now; anything left unreviewed is simply kept as it is.
                      </p>
                    )}
                    {current?.stats && (
                      <p className="text-xs text-gray-600">
                        Built: {Object.values(current.stats.controls).reduce((a, b) => a + b, 0)}{" "}
                        controls, {Object.values(current.stats.keys).reduce((a, b) => a + b, 0)}{" "}
                        keys replaced
                        {current.stats.skipped.length > 0 &&
                          ` · left as is: ${current.stats.skipped.join(", ")}`}
                        {current.kind === "docx" &&
                          " · ZIP contains the mapped .docx and the .htm"}
                      </p>
                    )}
                    {current?.kind === "docx" && (
                      <details className="group border rounded-lg">
                        <summary className="cursor-pointer select-none px-3 py-2 text-sm text-gray-700 flex items-center gap-2">
                          <ChevronRight className="h-4 w-4 transition-transform group-open:rotate-90" />
                          Optional: HTML exactly as Word makes it
                          {current.linkedHtml && (
                            <span className="text-xs text-green-700">· linked</span>
                          )}
                        </summary>
                        <div className="px-3 pb-3 border-t pt-3">
                          <WordHtmlPanel
                            file={current}
                            onDownloadDocx={handleDownloadMappedDocx}
                            onPickHtml={() => wordHtmlInputRef.current?.click()}
                            onUnlink={handleUnlinkWordHtml}
                            onBuild={handleBuild}
                          />
                        </div>
                      </details>
                    )}
                  </StepCard>

                  {/* ── advanced ── */}
                  <details className="group border rounded-lg">
                    <summary className="cursor-pointer select-none px-3 py-2 text-sm text-gray-500 flex items-center gap-2">
                      <Settings2 className="h-4 w-4" />
                      Advanced settings
                      <ChevronRight className="h-4 w-4 ml-auto transition-transform group-open:rotate-90" />
                    </summary>
                    <div className="px-3 pb-3 border-t pt-3">
                      <SettingsPanel
                        settings={settings}
                        onChange={setSettings}
                        onResetTemplates={() =>
                          setSettings((s) => ({
                            ...s,
                            templates: { ...DEFAULT_TEMPLATES },
                          }))
                        }
                      />
                    </div>
                  </details>
                </CardContent>
              </Card>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="border-t bg-card py-3 flex-shrink-0 text-center text-sm text-gray-500">
          Developed by Abdul Basit
        </div>
      </div>
    </>
  );
}

/* ───────────────────────── checklist pieces ───────────────────────── */

function StepCard({
  n,
  title,
  done,
  active,
  badge,
  children,
}: {
  n: number;
  title: string;
  done: boolean;
  active?: boolean;
  badge?: number;
  children: React.ReactNode;
}) {
  return (
    <section
      className={`rounded-xl border p-4 space-y-3 ${done ? "border-green-200 bg-green-50/40" : active ? "border-gray-900 bg-white shadow-sm" : "border-gray-200 bg-white"}`}
    >
      <div className="flex items-center gap-3">
        <div
          className={`w-7 h-7 rounded-full flex items-center justify-center text-sm font-semibold flex-shrink-0 ${done ? "bg-green-600 text-white" : active ? "bg-gray-900 text-white" : "bg-gray-200 text-gray-600"}`}
        >
          {done ? "✓" : n}
        </div>
        <h3 className="font-semibold text-base">{title}</h3>
        {badge !== undefined && (
          <span className="ml-auto text-xs px-2 py-0.5 rounded-full bg-orange-100 text-orange-700">
            {badge} to review
          </span>
        )}
      </div>
      <div className="pl-10 space-y-3">{children}</div>
    </section>
  );
}

function DecisionCard({
  candidate,
  index,
  onChoose,
}: {
  candidate: DocxCandidate;
  index: number;
  onChoose: (type: CandidateType) => void;
}) {
  const c = candidate;
  return (
    <div className="p-3 rounded-lg border border-orange-200 bg-orange-50 space-y-2">
      <div className="flex gap-3">
        <div className="w-12 h-12 flex-shrink-0 rounded border bg-white flex items-center justify-center overflow-hidden">
          {c.thumbnail ? (
            // eslint-disable-next-line @next/next/no-img-element
            <img src={c.thumbnail} alt="" className="max-w-full max-h-full object-contain" />
          ) : (
            <KindIcon kind={c.kind} />
          )}
        </div>
        <div className="flex-1 min-w-0">
          <div className="text-sm font-medium truncate" title={c.label}>
            #{index} {c.label}
          </div>
          <div className="text-xs font-mono text-gray-600 truncate" title={c.context}>
            {c.context.split("▮").map((part, i, arr) => (
              <span key={i}>
                {part}
                {i < arr.length - 1 && (
                  <span className="inline-block px-1 rounded bg-sky-200 text-sky-900">▮</span>
                )}
              </span>
            ))}
          </div>
        </div>
      </div>
      <div className="flex flex-wrap gap-1.5">
        <Button size="sm" variant="outline" className="h-7 text-xs bg-white" onClick={() => onChoose("checkbox")}>
          ☐ Checkbox
        </Button>
        <Button size="sm" variant="outline" className="h-7 text-xs bg-white" onClick={() => onChoose("checkboxChecked")}>
          ☑ Checked box
        </Button>
        <Button size="sm" variant="outline" className="h-7 text-xs bg-white" onClick={() => onChoose("textbox")}>
          ▭ Textbox
        </Button>
        <Button size="sm" variant="ghost" className="h-7 text-xs text-gray-600" onClick={() => onChoose("ignore")}>
          Skip
        </Button>
      </div>
    </div>
  );
}

/* ───────────────────────── sub components ───────────────────────── */

function LegendSwatch({
  bg,
  border,
  label,
}: {
  bg: string;
  border: string;
  label: string;
}) {
  return (
    <span className="flex items-center gap-1">
      <span
        className="inline-block w-3 h-3 rounded"
        style={{ background: bg, border: `1px solid ${border}` }}
      />
      {label}
    </span>
  );
}

const TYPE_LABELS: Record<CandidateType, string> = {
  checkbox: "Checkbox",
  checkboxChecked: "Checked checkbox",
  textbox: "Textbox",
  ignore: "Ignore",
};

function KindIcon({ kind }: { kind: DocxCandidate["kind"] }) {
  const cls = "h-5 w-5";
  switch (kind) {
    case "image":
      return <ImageIcon className={`${cls} text-gray-500`} />;
    case "shape":
      return <Shapes className={`${cls} text-gray-500`} />;
    case "blank":
      return <TextCursorInput className={`${cls} text-violet-600`} />;
    default:
      return <SquareCheck className={`${cls} text-sky-600`} />;
  }
}

function ControlsPanel({
  file,
  onTypeChange,
  onBulk,
}: {
  file: EsignFile;
  onTypeChange: (id: string, type: CandidateType) => void;
  onBulk: (mode: "suggested" | "ignore") => void;
}) {
  const { analysis, candidates, types, stats } = file;
  const assignments = assignControlIds(candidates, types, analysis);
  const undecided = candidates.filter(
    (c) => c.suggested === null && types[c.id] === undefined,
  ).length;
  const assigned = candidates.filter(
    (c) => (types[c.id] ?? c.suggested ?? "ignore") !== "ignore",
  ).length;
  const { checkboxes, textboxes, duplicateIds, missingNumbers } = analysis;
  const hasWarnings =
    duplicateIds.length > 0 ||
    missingNumbers.checkbox.length > 0 ||
    missingNumbers.textbox.length > 0;

  return (
    <div className="space-y-5">
      <p className="text-xs text-gray-500 truncate" title={file.file.name}>
        <span className="font-mono">{file.file.name}</span>
      </p>

      {file.kind === "docx" && (
        <div className="space-y-3">
          <div className="flex items-center justify-between gap-2 flex-wrap">
            <h4 className="font-medium text-sm flex items-center gap-2">
              <Wand2 className="h-4 w-4 text-sky-600" />
              Detected in Word
              <Badge variant="outline" className="text-xs">
                {candidates.length}
              </Badge>
              {undecided > 0 && (
                <Badge className="text-xs bg-orange-100 text-orange-700 border-orange-200 hover:bg-orange-100">
                  {undecided} need your decision
                </Badge>
              )}
            </h4>
            <div className="flex items-center gap-1">
              <Button
                size="sm"
                variant="ghost"
                className="h-7 text-xs"
                onClick={() => onBulk("suggested")}
                disabled={!candidates.length}
              >
                Use suggestions
              </Button>
              <Button
                size="sm"
                variant="ghost"
                className="h-7 text-xs"
                onClick={() => onBulk("ignore")}
                disabled={!candidates.length}
              >
                Ignore all
              </Button>
            </div>
          </div>

          {candidates.length === 0 ? (
            <p className="text-xs text-muted-foreground">
              No checkbox / textbox candidates found. If the client used something
              unusual, type <span className="font-mono">&lt;c1&gt;</span> /{" "}
              <span className="font-mono">&lt;t1&gt;</span> in Word and re-upload.
            </p>
          ) : (
            <div className="grid gap-2">
              {candidates.map((c, idx) => {
                const type = types[c.id] ?? c.suggested ?? "ignore";
                const marker =
                  type === "ignore"
                    ? null
                    : `${assignments[c.id].id}${type === "checkboxChecked" ? "c" : ""}`;
                const needsDecision = c.suggested === null && types[c.id] === undefined;
                return (
                  <div
                    key={c.id}
                    className={`p-2.5 rounded-lg border flex gap-3 ${needsDecision ? "bg-orange-50 border-orange-200" : type === "ignore" ? "bg-gray-50 border-gray-200 opacity-80" : "bg-white border-gray-200"}`}
                  >
                    <div className="w-11 h-11 flex-shrink-0 rounded border bg-white flex items-center justify-center overflow-hidden">
                      {c.thumbnail ? (
                        // eslint-disable-next-line @next/next/no-img-element
                        <img
                          src={c.thumbnail}
                          alt=""
                          className="max-w-full max-h-full object-contain"
                        />
                      ) : (
                        <KindIcon kind={c.kind} />
                      )}
                    </div>
                    <div className="flex-1 min-w-0 space-y-1.5">
                      <div className="flex items-center gap-2 text-xs">
                        <span className="text-gray-400">#{idx + 1}</span>
                        <span className="font-medium truncate" title={c.label}>
                          {c.label}
                        </span>
                        {marker && (
                          <span
                            className={`ml-auto font-mono px-1.5 py-0.5 rounded border ${type === "textbox" ? "bg-violet-50 border-violet-200 text-violet-800" : "bg-sky-50 border-sky-200 text-sky-800"}`}
                          >
                            &lt;{marker}&gt;
                          </span>
                        )}
                        {type === "ignore" && (
                          <span className="ml-auto text-gray-400 flex items-center gap-1">
                            <Ban className="h-3 w-3" /> ignored
                          </span>
                        )}
                      </div>
                      <div
                        className="text-[11px] font-mono text-gray-600 truncate"
                        title={c.context}
                      >
                        {c.context.split("▮").map((part, i, arr) => (
                          <span key={i}>
                            {part}
                            {i < arr.length - 1 && (
                              <span className="inline-block px-1 rounded bg-sky-200 text-sky-900">
                                ▮
                              </span>
                            )}
                          </span>
                        ))}
                      </div>
                      <select
                        value={type}
                        onChange={(e) => onTypeChange(c.id, e.target.value as CandidateType)}
                        className={`text-xs border rounded-md px-2 py-1 bg-white w-full ${needsDecision ? "border-orange-300" : "border-gray-200"}`}
                      >
                        {(Object.keys(TYPE_LABELS) as CandidateType[]).map((t) => (
                          <option key={t} value={t}>
                            {TYPE_LABELS[t]}
                            {c.suggested === t ? " (suggested)" : ""}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
          <p className="text-xs text-gray-500">
            {assigned} of {candidates.length} will be mapped. Numbering follows document
            order and continues after any markers already typed in the file.
          </p>
        </div>
      )}

      {hasWarnings && (
        <div className="p-3 rounded-lg border border-orange-200 bg-orange-50 text-xs text-orange-800 space-y-1">
          {duplicateIds.length > 0 && (
            <p>
              <strong>Duplicate ids:</strong>{" "}
              <span className="font-mono">{duplicateIds.join(", ")}</span> — the same
              marker appears more than once, so the output will contain duplicate
              control ids.
            </p>
          )}
          {missingNumbers.checkbox.length > 0 && (
            <p>
              <strong>Checkbox numbering gaps:</strong>{" "}
              <span className="font-mono">
                {missingNumbers.checkbox.map((n) => `c${n}`).join(", ")}
              </span>{" "}
              not found.
            </p>
          )}
          {missingNumbers.textbox.length > 0 && (
            <p>
              <strong>Textbox numbering gaps:</strong>{" "}
              <span className="font-mono">
                {missingNumbers.textbox.map((n) => `t${n}`).join(", ")}
              </span>{" "}
              not found.
            </p>
          )}
        </div>
      )}

      {(checkboxes.length > 0 || textboxes.length > 0 || file.kind === "html") && (
        <>
          <ControlTable
            title={file.kind === "docx" ? "Markers already typed — checkboxes" : "Checkboxes"}
            icon={<SquareCheck className="h-4 w-4 text-sky-600" />}
            items={checkboxes}
            stats={stats?.controls}
            emptyText="No <c1> / <c1c> markers found."
          />
          <ControlTable
            title={file.kind === "docx" ? "Markers already typed — textboxes" : "Textboxes"}
            icon={<TextCursorInput className="h-4 w-4 text-violet-600" />}
            items={textboxes}
            stats={stats?.controls}
            emptyText="No <t1> markers found."
          />
        </>
      )}

      <div className="text-xs text-gray-500 border-t pt-3 space-y-1">
        <p>
          <span className="font-mono">&lt;c1&gt;</span> → unchecked checkbox ·{" "}
          <span className="font-mono">&lt;c1c&gt;</span> → checked checkbox ·{" "}
          <span className="font-mono">&lt;t1&gt;</span> → text input
        </p>
      </div>
    </div>
  );
}

function ControlTable({
  title,
  icon,
  items,
  stats,
  emptyText,
}: {
  title: string;
  icon: React.ReactNode;
  items: EsignAnalysis["checkboxes"];
  stats?: Record<string, number>;
  emptyText: string;
}) {
  return (
    <div className="space-y-2">
      <h4 className="font-medium text-sm flex items-center gap-2">
        {icon}
        {title}
        <Badge variant="outline" className="text-xs">
          {items.length}
        </Badge>
      </h4>
      {items.length === 0 ? (
        <p className="text-xs text-muted-foreground">{emptyText}</p>
      ) : (
        <div className="border rounded-lg overflow-hidden">
          <table className="w-full text-xs">
            <thead className="bg-muted/60 text-left">
              <tr>
                <th className="px-2 py-1.5 font-medium">Marker</th>
                <th className="px-2 py-1.5 font-medium">Id</th>
                <th className="px-2 py-1.5 font-medium">State</th>
                <th className="px-2 py-1.5 font-medium text-right">Found</th>
                {stats && (
                  <th className="px-2 py-1.5 font-medium text-right">Replaced</th>
                )}
              </tr>
            </thead>
            <tbody>
              {items.map((c) => (
                <tr key={c.marker} className="border-t">
                  <td className="px-2 py-1.5 font-mono">&lt;{c.marker}&gt;</td>
                  <td className="px-2 py-1.5 font-mono">{c.id}</td>
                  <td className="px-2 py-1.5">
                    {c.type === "textbox" ? (
                      <span className="text-violet-700">text input</span>
                    ) : c.checked ? (
                      <span className="text-green-700">checked ☑</span>
                    ) : (
                      <span className="text-gray-600">unchecked ☐</span>
                    )}
                  </td>
                  <td className="px-2 py-1.5 text-right">{c.count}</td>
                  {stats && (
                    <td className="px-2 py-1.5 text-right">
                      {stats[c.marker] ?? 0}
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function SettingsPanel({
  settings,
  onChange,
  onResetTemplates,
}: {
  settings: EsignSettings;
  onChange: React.Dispatch<React.SetStateAction<EsignSettings>>;
  onResetTemplates: () => void;
}) {
  const setTemplate = (name: keyof EsignTemplates, value: string) =>
    onChange((s) => ({ ...s, templates: { ...s.templates, [name]: value } }));

  return (
    <div className="space-y-5">
      <div>
        <h4 className="font-medium text-sm">Output settings</h4>
        <p className="text-xs text-gray-500 mt-1">
          Saved in this browser and applied to every template you build.
        </p>
      </div>

      <div className="space-y-2">
        <Label className="text-xs text-gray-500">Key output format in the HTML</Label>
        <div className="grid grid-cols-2 gap-2">
          <ChoiceCard
            active={settings.keyOutput === "encoded"}
            onClick={() => onChange((s) => ({ ...s, keyOutput: "encoded" }))}
            title="&lt;&lt;KEY&gt;&gt;"
            subtitle="HTML-encoded (recommended) — browsers display it as <<KEY>>"
          />
          <ChoiceCard
            active={settings.keyOutput === "literal"}
            onClick={() => onChange((s) => ({ ...s, keyOutput: "literal" }))}
            title="<<KEY>>"
            subtitle="Raw text — browsers treat <KEY> as a tag and show an empty <>"
          />
        </div>
        <p className="text-[11px] text-gray-500">
          The mapped .docx always gets literal{" "}
          <span className="font-mono">&lt;&lt;KEY&gt;&gt;</span> (Word has no encoding
          issue). Use raw only if your esign system searches the HTML source for{" "}
          <span className="font-mono">&lt;&lt;KEY&gt;&gt;</span> without decoding.
        </p>
      </div>

      <div className="space-y-2">
        <ToggleRow
          checked={settings.replaceKeys}
          onChange={(v) => onChange((s) => ({ ...s, replaceKeys: v }))}
          label="Normalise placeholder keys"
          hint="Rewrite every <<key>> / &lt;&lt;key&gt;&gt; / «key» / @key to its canonical placeholder"
        />
        <ToggleRow
          checked={settings.replaceControls}
          onChange={(v) => onChange((s) => ({ ...s, replaceControls: v }))}
          label="Convert markers to form controls in the HTML"
          hint="Replace <c1>, <c1c>, <t1> with the HTML below (the .docx keeps the markers)"
        />
      </div>

      <div className="space-y-3 border-t pt-4">
        <div className="flex items-center justify-between">
          <h4 className="font-medium text-sm">Control templates</h4>
          <Button size="sm" variant="ghost" className="h-7 text-xs" onClick={onResetTemplates}>
            Reset to defaults
          </Button>
        </div>
        <p className="text-xs text-gray-500">
          Placeholders: <span className="font-mono">{"{id}"}</span> (c1 / t1),{" "}
          <span className="font-mono">{"{n}"}</span> (1),{" "}
          <span className="font-mono">{"{marker}"}</span> (c1c),{" "}
          <span className="font-mono">{"{type}"}</span> (checkbox / textbox)
        </p>
        <TemplateField
          label="Checkbox  <c1>"
          value={settings.templates.checkbox}
          onChange={(v) => setTemplate("checkbox", v)}
        />
        <TemplateField
          label="Checked checkbox  <c1c>"
          value={settings.templates.checkboxChecked}
          onChange={(v) => setTemplate("checkboxChecked", v)}
        />
        <TemplateField
          label="Textbox  <t1>"
          value={settings.templates.textbox}
          onChange={(v) => setTemplate("textbox", v)}
        />
      </div>
    </div>
  );
}

function WordHtmlPanel({
  file,
  onDownloadDocx,
  onPickHtml,
  onUnlink,
  onBuild,
}: {
  file: EsignFile;
  onDownloadDocx: () => void;
  onPickHtml: () => void;
  onUnlink: () => void;
  onBuild: () => void;
}) {
  const built = !!file.mappedDocxBytes;
  const linked = file.linkedHtml;
  const Step = ({
    n,
    title,
    children,
    done,
  }: {
    n: number;
    title: string;
    children: React.ReactNode;
    done?: boolean;
  }) => (
    <div className={`flex gap-3 p-3 rounded-lg border ${done ? "bg-green-50 border-green-200" : "bg-white border-gray-200"}`}>
      <div
        className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-semibold flex-shrink-0 ${done ? "bg-green-600 text-white" : "bg-gray-900 text-white"}`}
      >
        {done ? "✓" : n}
      </div>
      <div className="flex-1 min-w-0 space-y-2">
        <div className="text-sm font-medium">{title}</div>
        {children}
      </div>
    </div>
  );

  return (
    <div className="space-y-4">
      <div>
        <h4 className="font-medium text-sm">Step 2 — Word-quality HTML (optional)</h4>
        <p className="text-xs text-gray-500 mt-1">
          The built HTML normally comes from the in-browser converter (docx-preview).
          For a file that looks exactly like Word&apos;s own export, let Word do the
          conversion and upload the result here. The markers and keys inside it are
          turned into form controls the same way, and the ZIP will contain this HTML
          instead.
        </p>
      </div>

      <Step n={1} title="Build the template" done={built}>
        <p className="text-xs text-gray-500">
          Detect and map the checkboxes / textboxes, then click Build. This produces
          the mapped .docx with <span className="font-mono">&lt;c1&gt;</span> /{" "}
          <span className="font-mono">&lt;t1&gt;</span> markers.
        </p>
        {!built && (
          <Button size="sm" className="h-7 text-xs" onClick={onBuild}>
            Build now
          </Button>
        )}
      </Step>

      <Step n={2} title="Download the mapped .docx" done={built && !!linked}>
        <Button
          size="sm"
          variant="outline"
          className="h-7 text-xs flex items-center gap-1"
          disabled={!built}
          onClick={onDownloadDocx}
        >
          <Download className="h-3 w-3" />
          Download {baseName(file.file.name)}.docx
        </Button>
      </Step>

      <Step n={3} title="Save it as HTML in Word" done={built && !!linked}>
        <ol className="text-xs text-gray-600 list-decimal ml-4 space-y-0.5">
          <li>Open the downloaded .docx in Microsoft Word.</li>
          <li>
            File → Save As → Save as type:{" "}
            <span className="font-medium">Web Page, Filtered (*.htm; *.html)</span>
          </li>
          <li>Save, then close Word (pictures land in a &quot;_files&quot; folder — that is normal).</li>
        </ol>
      </Step>

      <Step n={4} title="Upload the .htm Word saved" done={!!linked}>
        {linked ? (
          <div className="flex items-center gap-2 text-xs flex-wrap">
            <span className="font-mono px-2 py-1 rounded border bg-white truncate max-w-[260px]">
              {linked.file.name}
            </span>
            <span className="text-gray-500">
              {linked.analysis.controls.length} markers · {linked.analysis.matchedKeys.length}{" "}
              keys · {linked.charset}
            </span>
            <Button size="sm" variant="ghost" className="h-7 text-xs" onClick={onPickHtml}>
              Replace
            </Button>
            <Button
              size="sm"
              variant="ghost"
              className="h-7 text-xs text-red-600 hover:bg-red-50"
              onClick={onUnlink}
            >
              Unlink
            </Button>
          </div>
        ) : (
          <Button
            size="sm"
            className="h-7 text-xs flex items-center gap-1"
            disabled={!built}
            onClick={onPickHtml}
          >
            <Upload className="h-3 w-3" />
            Upload Word .htm for this file
          </Button>
        )}
        {linked && linked.analysis.controls.length === 0 && file.candidates.length > 0 && (
          <p className="text-xs text-orange-700">
            No markers found in this .htm — make sure you saved the <em>mapped</em> .docx
            from step 2, not the original.
          </p>
        )}
      </Step>

      <div className="text-xs text-gray-500 border-t pt-3">
        Currently using:{" "}
        <span className="font-medium">
          {file.htmlSource === "word"
            ? "Word HTML"
            : file.htmlSource === "docx-preview"
              ? "in-browser conversion (docx-preview)"
              : "— (not built yet)"}
        </span>
        {linked && file.htmlSource !== "word" && " · click Build to switch to the Word HTML"}
      </div>
    </div>
  );
}

function ChoiceCard({
  active,
  onClick,
  title,
  subtitle,
}: {
  active: boolean;
  onClick: () => void;
  title: string;
  subtitle: string;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`text-left p-3 rounded-lg border transition-colors ${active ? "border-primary bg-primary/5" : "border-gray-200 hover:bg-muted"}`}
    >
      <div className="font-mono text-sm">{title}</div>
      <div className="text-[11px] text-gray-500 mt-1">{subtitle}</div>
    </button>
  );
}

function ToggleRow({
  checked,
  onChange,
  label,
  hint,
}: {
  checked: boolean;
  onChange: (v: boolean) => void;
  label: string;
  hint: string;
}) {
  return (
    <label className="flex items-start gap-3 p-3 rounded-lg border border-gray-200 cursor-pointer hover:bg-muted/40">
      <input
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.target.checked)}
        className="mt-0.5 h-4 w-4"
      />
      <span className="space-y-0.5">
        <span className="block text-sm">{label}</span>
        <span className="block text-xs text-gray-500">{hint}</span>
      </span>
    </label>
  );
}

function TemplateField({
  label,
  value,
  onChange,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
}) {
  return (
    <div className="space-y-1">
      <Label className="text-xs text-gray-500 font-mono">{label}</Label>
      <textarea
        value={value}
        onChange={(e) => onChange(e.target.value)}
        rows={2}
        spellCheck={false}
        className="w-full border border-gray-200 rounded-md px-3 py-2 text-xs font-mono focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white resize-y"
      />
    </div>
  );
}
