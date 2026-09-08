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
  Map as MapIcon,
  Search,
  Hash,
} from "lucide-react";
import { PlaceholderAutocomplete } from "@/components/Placeholderautocomplete";
import { KeysList } from "@/components/keys-list";
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
  /** docx only: the mapped .docx rendered to HTML, markers still as <c1>/<t1> text */
  mappedHtml?: string;
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
  /** number every tag afresh in document order (ignore numbers typed in Word) */
  renumber: boolean;
  templates: EsignTemplates;
}

type PanelTab = "keys" | "controls" | "wordhtml" | "settings";

const DEFAULT_SETTINGS: EsignSettings = {
  // encoded is the only form a browser renders as <<KEY>>; a literal
  // <<KEY>> in HTML is parsed as a tag and shows up as "<>"
  keyOutput: "encoded",
  replaceKeys: true,
  replaceControls: true,
  renumber: false,
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
    // unclassified ones stay undecided so the user sees them flagged
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
  /** plain-text key matches ("Student Name", "studentname") the user ticked
   *  to replace — OFF by default: only exact matches (<<KEY>>, @KEY, KEY)
   *  are replaced automatically, like the DocX Key Replacer */
  const [includedPlain, setIncludedPlain] = useState<string[]>([]);
  const [customText, setCustomText] = useState("");
  const [customValue, setCustomValue] = useState("");
  const [panelTab, setPanelTab] = useState<PanelTab>("controls");
  const [previewMode, setPreviewMode] = useState<"original" | "mapped" | "replaced">(
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
  // "matched" = keys that WILL be replaced: exact forms (<<KEY>>, @KEY, KEY)
  // plus the plain-text suggestions the user ticked
  const uniqueMatched = [
    ...new Set(
      files.flatMap((f) =>
        f.analysis.tokens
          .filter(
            (t) =>
              t.knownKey && (t.form !== "plain" || includedPlain.includes(t.raw)),
          )
          .map((t) => t.knownKey as string),
      ),
    ),
  ];
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
  /** everything plain-text that is NOT ticked stays untouched */
  const excluded = plainTokens
    .filter((t) => !includedPlain.includes(t.raw))
    .map((t) => t.raw);
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
      setPanelTab(cands || markers ? "controls" : "keys");
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
    setIncludedPlain([]);
    setCustomText("");
    setCustomValue("");
    setPreviewMode("original");
    setPanelTab("controls");
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
    const assignments = assignControlIds(item.candidates, item.types, item.analysis, {
      renumber: settings.renumber,
    });
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
        mappedHtml: html,
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
      mappedHtml: html,
      htmlSource: "docx-preview" as const,
      stats,
    };
  };

  const handleBuild = async () => {
    if (!files.length) return;
    const missing = uniqueUnmatched.filter((k) => !mappings[k]);
    if (missing.length) {
      toast(
        `${missing.length} key${missing.length > 1 ? "s" : ""} have no mapping and will be left as-is`,
        { icon: "⚠️", duration: 4000 },
      );
    }
    if (needsReview) {
      toast(
        `${needsReview} candidate${needsReview > 1 ? "s" : ""} without a decision will be ignored`,
        { icon: "⚠️", duration: 4000 },
      );
    }

    setIsProcessing(true);
    const toastId = toast.loading("Building esign templates…");
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
      // Word files: show the mapped tags first (which checkbox/textbox got
      // which <c1>/<t1>), the built HTML is one click away
      setPreviewMode(updated[selectedIndex]?.mappedHtml ? "mapped" : "replaced");
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
    } catch (err) {
      console.error(err);
      toast.dismiss(toastId);
      toast.error("Error building templates. Please try again.");
    } finally {
      setIsProcessing(false);
      setProcessingIndex(-1);
      setProcessingStep("");
    }
  };

  /* ── download ── */
  const handleDownload = async () => {
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

  /** which HTML the preview shows for the current mode */
  const pickPreviewHtml = (): string | undefined => {
    if (!current) return undefined;
    if (previewMode === "replaced" && current.replacedHtml !== undefined) return current.replacedHtml;
    if (previewMode === "mapped" && current.mappedHtml !== undefined) return current.mappedHtml;
    return current.html;
  };

  const handleCopyHtml = async () => {
    const html = pickPreviewHtml();
    if (!html) return;
    try {
      await navigator.clipboard.writeText(html);
      toast("HTML copied to clipboard", { icon: "📋", duration: 1500 });
    } catch {
      toast.error("Clipboard not available");
    }
  };

  const previewHtml = pickPreviewHtml();

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
              <span className="inline-flex rounded-lg border bg-muted/40 p-0.5 text-sm font-medium">
                <Link
                  href="/"
                  className="px-3 py-1.5 rounded-md text-gray-600 hover:text-gray-900 hover:bg-white transition-colors"
                >
                  DocX Key Replacer
                </Link>
                <span className="px-3 py-1.5 rounded-md bg-primary text-primary-foreground shadow-sm">
                  Esign Templates
                </span>
              </span>
              <span className="text-sm text-gray-500 font-normal">
                Word file in → mapped .docx + .htm out
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
                variant={panelTab === "settings" ? "default" : "outline"}
                className="flex items-center gap-2"
                onClick={() =>
                  setPanelTab((t) => (t === "settings" ? "controls" : "settings"))
                }
              >
                <Settings2 className="h-4 w-4" />
                Settings
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
        <div className="flex h-[calc(100vh-220px)]">
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
                        {previewMode === "mapped"
                          ? " · mapped tags"
                          : previewMode === "replaced" && current.htmlSource === "word"
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
                        {current.kind === "docx" && (
                          <Button
                            size="sm"
                            variant={previewMode === "mapped" ? "default" : "outline"}
                            className="h-7 text-xs"
                            disabled={current.mappedHtml === undefined}
                            onClick={() => setPreviewMode("mapped")}
                            title="The mapped Word file: every checkbox/textbox shown as its <c1>/<t1> tag"
                          >
                            Mapped tags
                          </Button>
                        )}
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
                        <Button
                          size="sm"
                          variant="outline"
                          className="h-7 text-xs flex items-center gap-1"
                          onClick={handleCopyHtml}
                          title="Copy HTML source"
                        >
                          <ClipboardCopy className="h-3 w-3" />
                          Copy HTML
                        </Button>
                      </span>
                    )}
                  </CardTitle>
                  {current && (
                    <div className="flex items-center gap-3 text-xs font-normal flex-wrap mt-1">
                      <LegendSwatch bg="#bbf7d0" border="#22c55e" label="Key" />
                      <LegendSwatch bg="#dbeafe" border="#3b82f6" label="Mapped key" />
                      <LegendSwatch bg="#fef08a" border="#eab308" label="Not in key list" />
                      <LegendSwatch bg="#ffedd5" border="#f97316" label="Plain-text key" />
                      <LegendSwatch bg="#fff7ed" border="#fdba74" label="Suggested (not replaced)" />
                      <LegendSwatch bg="#e0f2fe" border="#0ea5e9" label="Checkbox" />
                      <LegendSwatch bg="#ede9fe" border="#8b5cf6" label="Textbox" />
                    </div>
                  )}
                  {/* build result — shown right after Build, on the built preview */}
                  {current?.stats && previewMode !== "original" && !isProcessing && (
                    <div className="mt-2 rounded-md border border-green-200 bg-green-50 px-3 py-2 text-xs text-green-900 flex items-center gap-3 flex-wrap">
                      <CheckCircle2 className="h-4 w-4 text-green-600 flex-shrink-0" />
                      <span className="font-medium">Built.</span>
                      <span>
                        {Object.values(current.stats.keys).reduce((a, b) => a + b, 0)} keys
                        replaced
                      </span>
                      <span>
                        {Object.entries(current.stats.controls).filter(([m]) => m.startsWith("c")).reduce((a, [, n]) => a + n, 0)}{" "}
                        checkboxes
                      </span>
                      <span>
                        {Object.entries(current.stats.controls).filter(([m]) => m.startsWith("t")).reduce((a, [, n]) => a + n, 0)}{" "}
                        textboxes
                      </span>
                      {current.stats.skipped.length > 0 && (
                        <span className="text-orange-700">
                          {current.stats.skipped.length} left as-is:{" "}
                          <span className="font-mono">{current.stats.skipped.join(", ")}</span>
                        </span>
                      )}
                      {current.kind === "docx" && (
                        <span className="text-gray-600">
                          · {current.htmlSource === "word" ? "Word HTML" : "converted in browser"} ·
                          download gives .docx + .htm
                        </span>
                      )}
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
                            : "Upload the client's Word file"}
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

          {/* Right panel */}
          <div className="w-[520px] max-w-[46vw] flex-shrink-0 min-w-0">
            <div className="h-full overflow-hidden p-4">
              <Card className="h-full flex flex-col min-w-0">
                <CardHeader className="flex-shrink-0 border-b pb-3">
                  <div className="flex flex-wrap items-center gap-1">
                    <TabButton
                      active={panelTab === "controls"}
                      onClick={() => setPanelTab("controls")}
                      icon={<SquareCheck className="h-4 w-4" />}
                      label="Controls"
                      badge={needsReview > 0 ? needsReview : totalControls || undefined}
                      badgeTone={needsReview > 0 ? "orange" : "blue"}
                    />
                    <TabButton
                      active={panelTab === "keys"}
                      onClick={() => setPanelTab("keys")}
                      icon={<KeyRound className="h-4 w-4" />}
                      label="Keys"
                      badge={
                        uniqueUnmatched.length - mappedCount > 0
                          ? uniqueUnmatched.length - mappedCount
                          : undefined
                      }
                      badgeTone="orange"
                    />
                    {current?.kind === "docx" && (
                      <TabButton
                        active={panelTab === "wordhtml"}
                        onClick={() => setPanelTab("wordhtml")}
                        icon={<FileText className="h-4 w-4" />}
                        label="Word HTML"
                        badge={current.linkedHtml ? 1 : undefined}
                        badgeTone="blue"
                      />
                    )}
                    <TabButton
                      active={panelTab === "settings"}
                      onClick={() => setPanelTab("settings")}
                      icon={<Settings2 className="h-4 w-4" />}
                      label="Settings"
                    />
                  </div>
                </CardHeader>

                <CardContent className="flex-1 overflow-y-auto overflow-x-hidden space-y-5 pt-4 min-w-0">
                  {/* ── KEYS ── */}
                  {panelTab === "keys" && (
                    <>
                      <div className="flex items-center gap-2 flex-wrap">
                        <Badge className="text-xs bg-black text-white hover:bg-black/90">
                          {ESIGN_KNOWN_KEYS.length} known keys
                        </Badge>
                        {uniqueMatched.length > 0 && (
                          <Badge
                            variant="secondary"
                            className="text-xs gap-1 bg-green-100 text-green-700 border-green-200"
                          >
                            <CheckCircle2 className="h-3 w-3" />
                            {uniqueMatched.length} matched
                          </Badge>
                        )}
                        {uniqueUnmatched.length > 0 && (
                          <Badge
                            variant="secondary"
                            className="text-xs gap-1 bg-orange-100 text-orange-700 border-orange-200"
                          >
                            <AlertCircle className="h-3 w-3" />
                            {mappedCount}/{uniqueUnmatched.length} mapped
                          </Badge>
                        )}
                      </div>

                      {!files.length && (
                        <p className="text-sm text-muted-foreground">
                          Upload a file to see the keys it uses.
                        </p>
                      )}

                      {files.length > 0 && (
                        <div className="rounded-lg border border-green-200 bg-green-50 p-3 text-xs text-green-800">
                          Keys that match the list are replaced automatically. Anything
                          else is handled below under <strong>Manage Mappings</strong>,
                          exactly like the DocX Key Replacer.
                        </div>
                      )}

                      {(uniqueUnmatched.length > 0 || files.length > 0) && (
                        <h4 className="font-semibold text-sm flex items-center gap-2 border-t pt-4">
                          <MapIcon className="h-4 w-4 text-gray-500" />
                          Manage Mappings
                          {uniqueUnmatched.length > 0 && (
                            <Badge variant="destructive" className="ml-1 text-xs">
                              {uniqueUnmatched.length - mappedCount} unmatched
                            </Badge>
                          )}
                        </h4>
                      )}

                      {uniqueUnmatched.length > 0 && (
                        <div className="space-y-3">
                          <div>
                            <h4 className="font-medium text-sm">
                              Not in the key list — choose the right key
                            </h4>
                            <p className="text-xs text-gray-500">
                              Leave empty to keep the text as it is.
                            </p>
                          </div>
                          <div className="grid gap-2">
                            {uniqueUnmatched.map((key) => {
                              const isMapped = !!mappings[key];
                              return (
                                <div
                                  key={key}
                                  className={`p-3 rounded-lg border space-y-2 ${isMapped ? "bg-green-50 border-green-200" : "bg-yellow-50 border-yellow-200"}`}
                                >
                                  <div className="text-sm font-mono bg-white p-2 rounded border border-gray-200 flex items-center justify-between gap-2 break-all">
                                    <span>{key}</span>
                                    {isMapped && (
                                      <CheckCircle2 className="h-4 w-4 text-green-500 flex-shrink-0" />
                                    )}
                                  </div>
                                  <PlaceholderAutocomplete
                                    id={`map-${key}`}
                                    value={mappings[key] || ""}
                                    onChange={(val) =>
                                      setMappings((prev) => ({ ...prev, [key]: val }))
                                    }
                                    placeholder={`e.g., <<${key.replace(/^<</, "").replace(/>>$/, "").replace(/^@/, "").replace(/\s+/g, "").toUpperCase()}>>`}
                                    availableKeys={ESIGN_KNOWN_KEYS}
                                  />
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}

                      {plainTokens.length > 0 && (
                        <div className="space-y-2">
                          <div>
                            <h4 className="font-medium text-sm">
                              Suggested — key names written as normal text ({plainTokens.length})
                            </h4>
                            <p className="text-xs text-gray-500">
                              Not replaced automatically (only exact matches are). Tick the
                              ones that are placeholders, e.g.{" "}
                              <span className="font-mono">studentname</span> or{" "}
                              <span className="font-mono">Student Name</span>; leave labels
                              unticked.
                            </p>
                          </div>
                          <div className="grid gap-1.5">
                            {plainTokens.map((t) => {
                              const on = includedPlain.includes(t.raw);
                              return (
                                <label
                                  key={t.raw}
                                  className={`flex items-center gap-2 p-2 rounded-lg border text-sm cursor-pointer ${on ? "bg-orange-50 border-orange-200" : "bg-white border-gray-200 text-gray-600"}`}
                                >
                                  <input
                                    type="checkbox"
                                    checked={on}
                                    onChange={(e) =>
                                      setIncludedPlain((prev) =>
                                        e.target.checked
                                          ? [...prev, t.raw]
                                          : prev.filter((x) => x !== t.raw),
                                      )
                                    }
                                    className="h-4 w-4"
                                  />
                                  <span className="font-mono">{t.raw}</span>
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

                      {uniqueMatched.length > 0 && (
                        <div className="space-y-2">
                          <h4 className="font-medium text-sm">
                            Matched keys ({uniqueMatched.length})
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

                      {/* custom mapping */}
                      <div className="border-t pt-4 space-y-3">
                        <div>
                          <h4 className="font-medium text-sm">Map other text to a key</h4>
                          <p className="text-xs text-gray-500 mt-1">
                            Map any text in the document (with or without brackets) to a
                            placeholder, e.g.{" "}
                            <span className="font-mono">student name</span> →{" "}
                            <span className="font-mono">&lt;&lt;STUDENTNAME&gt;&gt;</span>
                          </p>
                        </div>
                        <div className="space-y-2">
                          <Input
                            value={customText}
                            onChange={(e) => setCustomText(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") handleAddCustom();
                            }}
                            placeholder="Text found in document  OR  <<KEY>>  OR  @KEY"
                            className="font-mono text-sm"
                          />
                          <PlaceholderAutocomplete
                            id="custom-value"
                            value={customValue}
                            onChange={setCustomValue}
                            placeholder="e.g., <<FIRSTNAME>>"
                            availableKeys={ESIGN_KNOWN_KEYS}
                          />
                          <Button
                            onClick={handleAddCustom}
                            className="flex items-center gap-2 w-full"
                          >
                            <Plus className="h-4 w-4" />
                            Add Mapping
                          </Button>
                        </div>
                      </div>

                      {customEntries.length > 0 && (
                        <div className="space-y-2 border-t pt-4">
                          <h4 className="font-medium text-sm">
                            Custom mappings
                            <Badge variant="outline" className="ml-2 text-xs">
                              {customEntries.length}
                            </Badge>
                          </h4>
                          <div className="grid gap-2">
                            {customEntries.map(([key, value]) => (
                              <div
                                key={key}
                                className="flex items-center gap-2 p-2.5 bg-gray-50 rounded-lg border border-gray-200"
                              >
                                <div className="flex-1 min-w-0 text-sm font-mono break-all">
                                  <span className="text-blue-600 font-semibold">
                                    &quot;{key}&quot;
                                  </span>
                                  <span className="mx-2 text-gray-500">→</span>
                                  <span>{value}</span>
                                </div>
                                <Button
                                  variant="ghost"
                                  size="sm"
                                  onClick={() => handleRemoveMapping(key)}
                                  className="h-8 w-8 p-0 hover:bg-red-100 hover:text-red-600 flex-shrink-0"
                                >
                                  <X className="h-4 w-4" />
                                </Button>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                      {/* the same placeholder list the DocX module shows */}
                      {files.length > 0 && (
                        <div className="space-y-2 border-t pt-4">
                          <h4 className="font-semibold text-sm">Available Placeholders</h4>
                          <div className="h-[420px]">
                            <KeysList
                              matchedKeys={uniqueMatched}
                              unmatchedKeys={[]}
                              onKeyUpdate={(key, value) =>
                                setMappings((prev) => ({ ...prev, [key]: value }))
                              }
                            />
                          </div>
                        </div>
                      )}
                    </>
                  )}

                  {/* ── CONTROLS ── */}
                  {panelTab === "controls" && (
                    <>
                      {!current && (
                        <p className="text-sm text-muted-foreground">
                          Upload a Word file to see its checkboxes and textboxes, or an
                          .htm that already contains markers.
                        </p>
                      )}
                      {current && (
                        <ControlsPanel
                          file={current}
                          onTypeChange={(id, type) =>
                            setCandidateType(selectedIndex, id, type)
                          }
                          onBulk={(mode) => setAllCandidateTypes(selectedIndex, mode)}
                          renumber={settings.renumber}
                          onRenumberChange={(v) => setSettings((s) => ({ ...s, renumber: v }))}
                        />
                      )}
                    </>
                  )}

                  {/* ── STEP 2: WORD HTML ── */}
                  {panelTab === "wordhtml" && current?.kind === "docx" && (
                    <WordHtmlPanel
                      file={current}
                      onDownloadDocx={handleDownloadMappedDocx}
                      onPickHtml={() => wordHtmlInputRef.current?.click()}
                      onUnlink={handleUnlinkWordHtml}
                      onBuild={handleBuild}
                    />
                  )}

                  {/* ── OUTPUT SETTINGS ── */}
                  {panelTab === "settings" && (
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
                  )}
                </CardContent>
              </Card>
            </div>
          </div>
        </div>

        {/* Bottom bar */}
        <div className="border-t bg-card p-4 flex-shrink-0">
          <div className="flex justify-center gap-4 mb-2">
            <Button
              onClick={handleBuild}
              disabled={!files.length || isProcessing || isUploading}
              className="flex items-center gap-2"
            >
              {isProcessing
                ? `Building ${processingIndex + 1} of ${files.length}…`
                : `Build (${files.length} file${files.length !== 1 ? "s" : ""})`}
            </Button>
            <Button
              onClick={handleDownload}
              disabled={!processedFiles.length || isProcessing}
              className="flex items-center gap-2"
              variant="outline"
            >
              <Download className="h-4 w-4" />
              {processedFiles.length === 1 && processedFiles[0].kind === "html"
                ? "Download File"
                : "Download ZIP (.docx + .htm)"}
            </Button>
          </div>
          <div className="text-center text-sm mt-[20px]">Developed by Abdul Basit</div>
        </div>
      </div>
    </>
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

function TabButton({
  active,
  onClick,
  icon,
  label,
  badge,
  badgeTone = "blue",
}: {
  active: boolean;
  onClick: () => void;
  icon: React.ReactNode;
  label: string;
  badge?: number;
  badgeTone?: "orange" | "blue";
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-md text-sm whitespace-nowrap transition-colors ${active ? "bg-primary text-primary-foreground" : "hover:bg-muted text-gray-700"}`}
    >
      {icon}
      {label}
      {badge !== undefined && (
        <span
          className={`text-[11px] px-1.5 rounded-full ${active ? "bg-white/20" : badgeTone === "orange" ? "bg-orange-100 text-orange-700" : "bg-blue-100 text-blue-700"}`}
        >
          {badge}
        </span>
      )}
    </button>
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
    case "marker":
      return <Hash className={`${cls} text-sky-700`} />;
    default:
      return <SquareCheck className={`${cls} text-sky-600`} />;
  }
}

function ControlsPanel({
  file,
  onTypeChange,
  onBulk,
  renumber,
  onRenumberChange,
}: {
  file: EsignFile;
  onTypeChange: (id: string, type: CandidateType) => void;
  onBulk: (mode: "suggested" | "ignore") => void;
  renumber: boolean;
  onRenumberChange: (v: boolean) => void;
}) {
  const { analysis, candidates, types, stats } = file;
  const assignments = assignControlIds(candidates, types, analysis, { renumber });
  const typedCount = candidates.filter((c) => c.kind === "marker").length;
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

  // search across label, context text, kind, chosen type and marker id
  const [search, setSearch] = useState("");
  const q = search.trim().toLowerCase();
  const visible = candidates.filter((c) => {
    if (!q) return true;
    const type = types[c.id] ?? c.suggested ?? "ignore";
    const marker = type === "ignore" ? "" : `${assignments[c.id].id}${type === "checkboxChecked" ? "c" : ""}`;
    return [c.label, c.context, c.kind, TYPE_LABELS[type], marker, c.text ?? ""]
      .join(" ")
      .toLowerCase()
      .includes(q);
  });

  return (
    <div className="space-y-5">
      <p className="text-xs text-gray-500 truncate" title={file.file.name}>
        <span className="font-mono">{file.file.name}</span>
      </p>

      {file.kind === "docx" && candidates.length > 0 && (
        <div className="relative">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
          <Input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search controls… (e.g. lunch, c3, textbox, picture)"
            className="pl-9 h-9 text-sm"
          />
          {q && (
            <span className="absolute right-3 top-1/2 -translate-y-1/2 text-xs text-gray-500">
              {visible.length}/{candidates.length}
            </span>
          )}
        </div>
      )}

      {file.kind === "docx" && (
        <div className="space-y-3">
          <div className="flex items-center justify-between gap-2 flex-wrap">
            <h4 className="font-medium text-sm flex items-center gap-2">
              <Wand2 className="h-4 w-4 text-sky-600" />
              Found in the Word file
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
                Accept all suggestions
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

          <label className="flex items-start gap-2 p-2.5 rounded-lg border border-sky-200 bg-sky-50 text-xs cursor-pointer">
            <input
              type="checkbox"
              checked={renumber}
              onChange={(e) => onRenumberChange(e.target.checked)}
              className="mt-0.5 h-4 w-4"
            />
            <span>
              <span className="font-medium">Renumber all tags in document order</span>
              <span className="block text-gray-600 mt-0.5">
                Off: tags typed in Word ({typedCount}) keep their numbers, new ones continue
                after them. On: everything is numbered afresh top to bottom (c1, c2 … / t1,
                t2 …), so a tag added in the middle shifts the ones after it. Type{" "}
                <span className="font-mono">&lt;c&gt;</span>,{" "}
                <span className="font-mono">&lt;cc&gt;</span> (checked) or{" "}
                <span className="font-mono">&lt;t&gt;</span> in Word to add one without
                choosing a number.
              </span>
            </span>
          </label>

          {candidates.length === 0 ? (
            <p className="text-xs text-muted-foreground">
              No checkbox / textbox candidates found. If the client used something
              unusual, type <span className="font-mono">&lt;c1&gt;</span> /{" "}
              <span className="font-mono">&lt;t1&gt;</span> in Word and re-upload.
            </p>
          ) : (
            <div className="grid gap-2">
              {visible.length === 0 && (
                <p className="text-xs text-muted-foreground">No control matches “{search}”.</p>
              )}
              {visible.map((c) => {
                const idx = candidates.indexOf(c);
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
            {assigned} of {candidates.length} will become controls, numbered in the order they appear.
          </p>
        </div>
      )}

      {hasWarnings && !(file.kind === "docx" && renumber) && (
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

      {/* for Word files typed tags are listed above as candidates; the tables
          only make sense for an .htm that already carries markers */}
      {file.kind === "html" && (
        <>
          <ControlTable
            title="Checkboxes"
            icon={<SquareCheck className="h-4 w-4 text-sky-600" />}
            items={checkboxes}
            stats={stats?.controls}
            emptyText="No <c1> / <c1c> markers found."
          />
          <ControlTable
            title="Textboxes"
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
          label="Convert tags to form controls in the HTML"
          hint="Replace <c1>, <c1c>, <t1> with the HTML below (the .docx keeps the markers)"
        />
        <ToggleRow
          checked={settings.renumber}
          onChange={(v) => onChange((s) => ({ ...s, renumber: v }))}
          label="Renumber all tags in document order"
          hint="Ignore the numbers typed in Word and number every checkbox / textbox afresh from top to bottom"
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
        <h4 className="font-medium text-sm">Word-quality HTML (optional)</h4>
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
