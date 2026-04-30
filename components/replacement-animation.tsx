"use client";

import { motion, AnimatePresence } from "framer-motion";
import { FileText, ArrowRight, Sparkles, CheckCircle2 } from "lucide-react";
import { useEffect, useState, useRef, useMemo } from "react";

interface ReplacementAnimationProps {
  isOpen: boolean;
  processingIndex: number;
  totalFiles: number;
  currentFileName?: string;
  matchedKeys: string[];
  unmatchedKeys: string[];
  customPlaceholders: Record<string, string>;
}

export function ReplacementAnimation({
  isOpen,
  processingIndex,
  totalFiles,
  currentFileName,
  matchedKeys,
  unmatchedKeys,
  customPlaceholders,
}: ReplacementAnimationProps) {
  const [show, setShow] = useState(false);
  const startRef = useRef(0);
  const minDuration = 3500;

  const transformations = useMemo(() => {
    const items: {
      old: string;
      new: string;
      type: "matched" | "mapped" | "skipped";
    }[] = [];
    matchedKeys.forEach((k) => items.push({ old: k, new: k, type: "matched" }));
    unmatchedKeys.forEach((k) => {
      const mapped = customPlaceholders[k];
      items.push(
        mapped
          ? { old: k, new: mapped, type: "mapped" }
          : { old: k, new: "—", type: "skipped" },
      );
    });
    return items;
  }, [matchedKeys, unmatchedKeys, customPlaceholders]);

  useEffect(() => {
    if (isOpen) {
      startRef.current = Date.now();
      setShow(true);
    } else if (show) {
      const elapsed = Date.now() - startRef.current;
      const remain = Math.max(0, minDuration - elapsed);
      const timer = setTimeout(() => setShow(false), remain);
      return () => clearTimeout(timer);
    }
  }, [isOpen, show]);

  const progress =
    totalFiles > 0 ? ((processingIndex + 1) / totalFiles) * 100 : 0;

  return (
    <AnimatePresence>
      {show && (
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 0.4 }}
          className="fixed inset-0 z-[100] flex items-center justify-center"
          style={{
            backgroundColor: "rgba(255, 255, 255, 0.05)",
            backdropFilter: "blur(20px) saturate(180%)",
            WebkitBackdropFilter: "blur(20px) saturate(180%)",
          }}
        >
          {/* Main Glass Card */}
          <motion.div
            initial={{ scale: 0.85, opacity: 0, y: 40 }}
            animate={{ scale: 1, opacity: 1, y: 0 }}
            exit={{ scale: 0.9, opacity: 0, y: 20 }}
            transition={{ duration: 0.5, ease: [0.23, 1, 0.32, 1] }}
            className="relative w-full max-w-lg mx-6 rounded-3xl overflow-hidden"
            style={{
              background:
                "linear-gradient(135deg, rgba(255,255,255,0.15) 0%, rgba(255,255,255,0.05) 100%)",
              border: "1px solid rgba(255,255,255,0.2)",
              boxShadow:
                "0 25px 50px -12px rgba(0,0,0,0.15), inset 0 1px 0 rgba(255,255,255,0.2)",
            }}
          >
            {/* Shimmer border effect */}
            <motion.div
              className="absolute inset-0 rounded-3xl"
              style={{
                background:
                  "linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent)",
                backgroundSize: "200% 100%",
              }}
              animate={{ backgroundPosition: ["200% 0", "-200% 0"] }}
              transition={{ repeat: Infinity, duration: 3, ease: "linear" }}
            />

            {/* Content */}
            <div className="relative p-8">
              {/* Floating Document Icon */}
              <div className="flex justify-center mb-6">
                <motion.div
                  className="relative"
                  animate={{ y: [0, -8, 0] }}
                  transition={{
                    repeat: Infinity,
                    duration: 3,
                    ease: "easeInOut",
                  }}
                >
                  {/* Glow */}
                  <div className="absolute inset-0 bg-cyan-400/30 blur-2xl rounded-full scale-150" />

                  <div className="relative bg-white/20 backdrop-blur-xl rounded-2xl p-4 border border-white/30">
                    <FileText className="w-10 h-10 text-cyan-600" />
                  </div>
                </motion.div>
              </div>

              {/* Title with typewriter feel */}
              <motion.h2
                className="text-2xl font-bold text-center mb-1"
                style={{ color: "#1e293b" }}
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                transition={{ delay: 0.2 }}
              >
                Replacing Keys
              </motion.h2>
              <motion.p
                className="text-center text-sm mb-6"
                style={{ color: "#64748b" }}
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                transition={{ delay: 0.3 }}
              >
                {currentFileName}
              </motion.p>

              {/* Progress Bar - Liquid fill */}
              <div className="relative h-3 bg-slate-200/50 rounded-full overflow-hidden mb-6">
                <motion.div
                  className="absolute inset-y-0 left-0 rounded-full"
                  style={{
                    background: "linear-gradient(90deg, #06b6d4, #3b82f6)",
                  }}
                  initial={{ width: 0 }}
                  animate={{ width: `${progress}%` }}
                  transition={{ duration: 0.6, ease: "easeOut" }}
                />
                {/* Shine on progress */}
                <motion.div
                  className="absolute inset-y-0 w-20 rounded-full"
                  style={{
                    background:
                      "linear-gradient(90deg, transparent, rgba(255,255,255,0.4), transparent)",
                  }}
                  animate={{ x: ["-100%", "400%"] }}
                  transition={{ repeat: Infinity, duration: 2, ease: "linear" }}
                />
              </div>

              {/* Transformations - Horizontal scrolling cards */}
              <div className="relative h-16 mb-4 overflow-hidden">
                <motion.div
                  className="flex gap-3 absolute"
                  animate={{ x: [0, -300] }}
                  transition={{
                    repeat: Infinity,
                    duration: 15,
                    ease: "linear",
                  }}
                >
                  {[...transformations, ...transformations].map((t, i) => (
                    <div
                      key={i}
                      className="flex-shrink-0 flex items-center gap-2 px-4 py-2 rounded-xl border"
                      style={{
                        background:
                          t.type === "skipped"
                            ? "rgba(148, 163, 184, 0.1)"
                            : "rgba(34, 197, 94, 0.08)",
                        borderColor:
                          t.type === "skipped"
                            ? "rgba(148, 163, 184, 0.2)"
                            : "rgba(34, 197, 94, 0.2)",
                      }}
                    >
                      <span
                        className="text-xs font-mono"
                        style={{ color: "#64748b" }}
                      >
                        {t.old.length > 12 ? t.old.slice(0, 12) + "..." : t.old}
                      </span>
                      <ArrowRight className="w-3 h-3 text-cyan-500" />
                      <span
                        className={`text-xs font-mono ${
                          t.type === "skipped"
                            ? "text-slate-400 line-through"
                            : "text-green-600"
                        }`}
                      >
                        {t.new.length > 12 ? t.new.slice(0, 12) + "..." : t.new}
                      </span>
                      {t.type === "mapped" && (
                        <Sparkles className="w-3 h-3 text-green-500" />
                      )}
                    </div>
                  ))}
                </motion.div>
              </div>

              {/* Stats Row */}
              <div className="flex justify-center gap-6">
                <div className="text-center">
                  <div className="text-lg font-bold text-green-600">
                    {matchedKeys.length}
                  </div>
                  <div className="text-[10px] uppercase tracking-wider text-slate-400">
                    Matched
                  </div>
                </div>
                <div className="w-px bg-slate-300/50" />
                <div className="text-center">
                  <div className="text-lg font-bold text-blue-600">
                    {unmatchedKeys.filter((k) => customPlaceholders[k]).length}
                  </div>
                  <div className="text-[10px] uppercase tracking-wider text-slate-400">
                    Mapped
                  </div>
                </div>
                <div className="w-px bg-slate-300/50" />
                <div className="text-center">
                  <div className="text-lg font-bold text-slate-500">
                    {unmatchedKeys.filter((k) => !customPlaceholders[k]).length}
                  </div>
                  <div className="text-[10px] uppercase tracking-wider text-slate-400">
                    Skipped
                  </div>
                </div>
              </div>

              {/* File indicator */}
              <div className="mt-4 text-center text-xs text-slate-400">
                File {Math.min(processingIndex + 1, totalFiles)} of {totalFiles}
              </div>
            </div>
          </motion.div>
        </motion.div>
      )}
    </AnimatePresence>
  );
}
