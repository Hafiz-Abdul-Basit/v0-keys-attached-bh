"use client";
import React, { useState, useRef, useEffect, useCallback } from "react";
import { Input } from "@/components/ui/input";
import { Search, ChevronDown, X } from "lucide-react";

interface PlaceholderAutocompleteProps {
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  availableKeys: string[]; 
  id?: string;
  className?: string;
}

export function PlaceholderAutocomplete({
  value,
  onChange,
  placeholder = "e.g., <<PLACEHOLDER>>",
  availableKeys,
  id,
  className,
}: PlaceholderAutocompleteProps) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState(value || "");
  const [activeIndex, setActiveIndex] = useState(-1);

  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);
  const itemRefs = useRef<(HTMLButtonElement | null)[]>([]);

  // Sync external value changes
  useEffect(() => {
    setSearch(value || "");
  }, [value]);

  // Compute filtered list
  const filtered = (() => {
    const cleanSearch = search
      .replace(/^<</, "")
      .replace(/>>$/, "")
      .toLowerCase()
      .trim();
    if (!cleanSearch) return availableKeys;
    return availableKeys.filter((key) =>
      key.toLowerCase().includes(cleanSearch),
    );
  })();

  // Keep itemRefs array sized correctly
  useEffect(() => {
    itemRefs.current = itemRefs.current.slice(0, filtered.length);
  }, [filtered.length]);

  // Close on outside click
  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (
        containerRef.current &&
        !containerRef.current.contains(e.target as Node)
      ) {
        setOpen(false);
        setActiveIndex(-1);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  // Scroll active item into view
  useEffect(() => {
    if (activeIndex >= 0 && itemRefs.current[activeIndex]) {
      itemRefs.current[activeIndex]?.scrollIntoView({
        block: "nearest",
        behavior: "smooth",
      });
    }
  }, [activeIndex]);

  const handleSelect = useCallback(
    (key: string) => {
      const formatted = key.startsWith("<<") ? key : `<<${key}>>`;
      setSearch(formatted);
      onChange(formatted);
      setOpen(false);
      setActiveIndex(-1);
      inputRef.current?.focus();
    },
    [onChange],
  );

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setSearch(e.target.value);
    onChange(e.target.value);
    setActiveIndex(-1);
    setOpen(true);
  };

  const handleClear = () => {
    setSearch("");
    onChange("");
    setOpen(false);
    setActiveIndex(-1);
    inputRef.current?.focus();
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (!open) {
      // Open dropdown on ArrowDown / ArrowUp if closed
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        setOpen(true);
        setActiveIndex(e.key === "ArrowDown" ? 0 : filtered.length - 1);
      }
      return;
    }

    switch (e.key) {
      case "ArrowDown":
        e.preventDefault();
        setActiveIndex((prev) => (prev < filtered.length - 1 ? prev + 1 : 0));
        break;

      case "ArrowUp":
        e.preventDefault();
        setActiveIndex((prev) => (prev > 0 ? prev - 1 : filtered.length - 1));
        break;

      case "Enter":
        e.preventDefault();
        if (activeIndex >= 0 && filtered[activeIndex]) {
          handleSelect(filtered[activeIndex]);
        }
        break;

      case "Tab":
        // Tab selects the highlighted item (if any), then moves focus naturally
        if (activeIndex >= 0 && filtered[activeIndex]) {
          e.preventDefault();
          handleSelect(filtered[activeIndex]);
        } else {
          setOpen(false);
          setActiveIndex(-1);
        }
        break;

      case "Escape":
        e.preventDefault();
        setOpen(false);
        setActiveIndex(-1);
        break;

      case "Home":
        e.preventDefault();
        setActiveIndex(0);
        break;

      case "End":
        e.preventDefault();
        setActiveIndex(filtered.length - 1);
        break;
    }
  };

  const listboxId = `autocomplete-list-${id ?? "default"}`;

  return (
    <div ref={containerRef} className={`relative ${className || ""}`}>
      <div className="relative flex items-center">
        <Input
          ref={inputRef}
          id={id}
          role="combobox"
          aria-autocomplete="list"
          aria-expanded={open}
          aria-controls={listboxId}
          aria-activedescendant={
            activeIndex >= 0 ? `${listboxId}-option-${activeIndex}` : undefined
          }
          value={search}
          onChange={handleInputChange}
          onFocus={() => setOpen(true)}
          onKeyDown={handleKeyDown}
          placeholder={placeholder}
          className="font-mono text-sm w-full pr-16"
          autoComplete="off"
        />
        <div className="absolute right-2 flex items-center gap-1">
          {search && (
            <button
              type="button"
              tabIndex={-1}
              onClick={handleClear}
              className="text-gray-400 hover:text-gray-600 transition-colors p-0.5 rounded"
              aria-label="Clear"
            >
              <X className="h-3 w-3" />
            </button>
          )}
          <button
            type="button"
            tabIndex={-1}
            onClick={() => {
              setOpen((prev) => !prev);
              setActiveIndex(-1);
              inputRef.current?.focus();
            }}
            className="text-gray-400 hover:text-gray-600 transition-colors p-0.5 rounded"
            aria-label="Toggle dropdown"
          >
            <ChevronDown
              className={`h-3.5 w-3.5 transition-transform duration-200 ${
                open ? "rotate-180" : ""
              }`}
            />
          </button>
        </div>
      </div>

      {open && (
        <div
          id={listboxId}
          role="listbox"
          ref={listRef}
          className="absolute z-[9999] mt-1 w-full bg-white border border-gray-200 rounded-lg shadow-lg overflow-hidden"
          style={{ maxHeight: "220px" }}
        >
          {/* Hint bar */}
          <div className="flex items-center gap-2 px-3 py-2 border-b border-gray-100 bg-gray-50 flex-wrap">
            <Search className="h-3.5 w-3.5 text-gray-400 flex-shrink-0" />
            <span className="text-xs text-gray-400 flex items-center gap-1 flex-wrap">
              {filtered.length}/{availableKeys.length} keys &nbsp;·&nbsp;
              <kbd className="font-sans bg-gray-200 text-gray-500 rounded px-1 py-0.5 text-[10px]">
                ↑
              </kbd>
              <kbd className="font-sans bg-gray-200 text-gray-500 rounded px-1 py-0.5 text-[10px]">
                ↓
              </kbd>
              navigate &nbsp;·&nbsp;
              <kbd className="font-sans bg-gray-200 text-gray-500 rounded px-1 py-0.5 text-[10px]">
                ↵
              </kbd>
              select &nbsp;·&nbsp;
              <kbd className="font-sans bg-gray-200 text-gray-500 rounded px-1 py-0.5 text-[10px]">
                Esc
              </kbd>
              close
            </span>
          </div>

          <div className="overflow-y-auto" style={{ maxHeight: "170px" }}>
            {filtered.length === 0 ? (
              <div className="px-3 py-4 text-center text-sm text-gray-400">
                No matching keys found
              </div>
            ) : (
              filtered.map((key, index) => {
                const formatted = `<<${key}>>`;
                const isSelected = search === formatted || search === key;
                const isActive = index === activeIndex;

                return (
                  <button
                    key={key}
                    id={`${listboxId}-option-${index}`}
                    role="option"
                    aria-selected={isSelected}
                    ref={(el) => {
                      itemRefs.current[index] = el;
                    }}
                    type="button"
                    onMouseDown={(e) => {
                      e.preventDefault(); // prevent blur race
                      handleSelect(key);
                    }}
                    onMouseEnter={() => setActiveIndex(index)}
                    className={`w-full text-left px-3 py-2 text-sm font-mono transition-colors flex items-center justify-between
                      ${isActive ? "bg-blue-100 text-blue-800" : ""}
                      ${isSelected && !isActive ? "bg-blue-50 text-blue-700 font-semibold" : ""}
                      ${!isActive && !isSelected ? "text-gray-700 hover:bg-gray-50" : ""}
                    `}
                  >
                    <span className="truncate">{formatted}</span>
                    <span className="flex items-center gap-1 ml-2 flex-shrink-0">
                      {isActive && (
                        <kbd className="font-sans bg-blue-200 text-blue-700 rounded px-1 py-0.5 text-[10px] leading-tight">
                          ↵
                        </kbd>
                      )}
                      {isSelected && (
                        <span className="text-xs text-blue-500">✓</span>
                      )}
                    </span>
                  </button>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
}
