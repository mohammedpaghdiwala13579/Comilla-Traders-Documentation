import React, { useState, useEffect, useRef, useCallback } from "react";
import { AnimatePresence, motion } from "motion/react";
import {
  Bold,
  Italic,
  Underline,
  Highlighter,
  Baseline,
  ChevronDown,
  RemoveFormatting,
  X,
  Check,
} from "lucide-react";
import { applyInlineFormatting, InlineFormatType } from "../utils/textFormatter";

interface FloatingTextToolbarProps {
  onFormatted?: () => void;
}

const TEXT_COLORS = [
  { name: "Default (Black)", value: "#000000" },
  { name: "Slate Dark", value: "#334155" },
  { name: "Crimson Red", value: "#DC2626" },
  { name: "Dark Red", value: "#991B1B" },
  { name: "Emerald Green", value: "#16A34A" },
  { name: "Forest Green", value: "#166534" },
  { name: "Royal Blue", value: "#2563EB" },
  { name: "Navy Blue", value: "#1E40AF" },
  { name: "Amber Orange", value: "#D97706" },
  { name: "Purple", value: "#9333EA" },
];

const HIGHLIGHT_COLORS = [
  { name: "No Color", value: "transparent" },
  { name: "Classic Yellow", value: "#FEF08A" },
  { name: "Mint Green", value: "#BBF7D0" },
  { name: "Sky Blue", value: "#BAE6FD" },
  { name: "Soft Pink", value: "#FBCFE8" },
  { name: "Warm Orange", value: "#FED7AA" },
  { name: "Lavender Purple", value: "#E9D5FF" },
  { name: "Bright Yellow", value: "#FDE047" },
  { name: "Bright Green", value: "#4ADE80" },
  { name: "Bright Cyan", value: "#38BDF8" },
];

export const FloatingTextToolbar: React.FC<FloatingTextToolbarProps> = ({ onFormatted }) => {
  const [visible, setVisible] = useState<boolean>(false);
  const [position, setPosition] = useState<{ top: number; left: number }>({ top: 0, left: 0 });
  const [placement, setPlacement] = useState<"top" | "bottom">("top");

  // Active format indicators
  const [isBold, setIsBold] = useState<boolean>(false);
  const [isItalic, setIsItalic] = useState<boolean>(false);
  const [isUnderline, setIsUnderline] = useState<boolean>(false);
  const [currentFontSize, setCurrentFontSize] = useState<number>(11);
  const [currentTextColor, setCurrentTextColor] = useState<string>("#000000");
  const [currentHighlightColor, setCurrentHighlightColor] = useState<string>("#FEF08A");

  // Dropdown popovers
  const [openDropdown, setOpenDropdown] = useState<"textColor" | "highlight" | "fontSize" | null>(null);

  // References
  const toolbarRef = useRef<HTMLDivElement>(null);
  const savedRangeRef = useRef<Range | null>(null);
  const targetEditableRef = useRef<HTMLElement | null>(null);
  const isInteractingRef = useRef<boolean>(false);

  // Detect active formatting on current selection
  const detectActiveFormatting = useCallback((range: Range) => {
    try {
      setIsBold(document.queryCommandState("bold"));
      setIsItalic(document.queryCommandState("italic"));
      setIsUnderline(document.queryCommandState("underline"));

      const container = range.commonAncestorContainer;
      const node = container.nodeType === Node.ELEMENT_NODE ? (container as HTMLElement) : container.parentElement;
      if (node) {
        const computed = window.getComputedStyle(node);
        if (computed.color) {
          setCurrentTextColor(computed.color);
        }
        if (computed.backgroundColor && computed.backgroundColor !== "rgba(0, 0, 0, 0)" && computed.backgroundColor !== "transparent") {
          setCurrentHighlightColor(computed.backgroundColor);
        }
        if (computed.fontSize) {
          const px = parseFloat(computed.fontSize);
          if (!isNaN(px)) {
            const pt = Math.round(px * 0.75);
            setCurrentFontSize(pt);
          }
        }
      }
    } catch (e) {
      // Ignore quirks
    }
  }, []);

  // Update floating position based on selection bounds
  const updatePosition = useCallback(() => {
    if (typeof window === "undefined") return;

    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) {
      // Hide toolbar if no active selection and not interacting with dropdown
      if (!isInteractingRef.current) {
        setVisible(false);
        setOpenDropdown(null);
      }
      return;
    }

    const text = sel.toString().trim();
    if (!text) {
      if (!isInteractingRef.current) {
        setVisible(false);
        setOpenDropdown(null);
      }
      return;
    }

    const range = sel.getRangeAt(0);
    const commonAncestor = range.commonAncestorContainer;
    const editableEl = (
      commonAncestor instanceof HTMLElement ? commonAncestor : commonAncestor.parentElement
    )?.closest<HTMLElement>("[contenteditable='true']");

    if (!editableEl) {
      if (!isInteractingRef.current) {
        setVisible(false);
        setOpenDropdown(null);
      }
      return;
    }

    const rect = range.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) {
      return;
    }

    // Cache the range and target element
    savedRangeRef.current = range.cloneRange();
    targetEditableRef.current = editableEl;
    detectActiveFormatting(range);

    const toolbarWidth = 320;
    const toolbarHeight = 44;
    const viewportWidth = window.innerWidth;

    let top = rect.top - toolbarHeight - 10;
    let computedPlacement: "top" | "bottom" = "top";

    // If too close to top of viewport, flip to bottom
    if (top < 12) {
      top = rect.bottom + 10;
      computedPlacement = "bottom";
    }

    let left = rect.left + rect.width / 2 - toolbarWidth / 2;
    // Keep within horizontal screen bounds
    left = Math.max(12, Math.min(viewportWidth - toolbarWidth - 12, left));

    setPosition({ top, left });
    setPlacement(computedPlacement);
    setVisible(true);
  }, [detectActiveFormatting]);

  useEffect(() => {
    let timeoutId: NodeJS.Timeout | null = null;

    const handleSelectionChange = () => {
      // Small debounce to avoid flickering while dragging mouse
      if (timeoutId) clearTimeout(timeoutId);
      timeoutId = setTimeout(() => {
        updatePosition();
      }, 40);
    };

    const handleMouseUp = () => {
      // Instant check on mouse up after selection
      setTimeout(() => {
        updatePosition();
      }, 10);
    };

    const handleKeyUp = (e: KeyboardEvent) => {
      // Shift keys or navigation selection
      if (e.shiftKey || ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End"].includes(e.key)) {
        setTimeout(() => {
          updatePosition();
        }, 20);
      }
    };

    const handleScrollOrResize = () => {
      if (visible) {
        updatePosition();
      }
    };

    const handleDocumentMouseDown = (e: MouseEvent) => {
      if (toolbarRef.current && toolbarRef.current.contains(e.target as Node)) {
        isInteractingRef.current = true;
        return;
      }
      isInteractingRef.current = false;
      setOpenDropdown(null);
    };

    document.addEventListener("selectionchange", handleSelectionChange);
    window.addEventListener("mouseup", handleMouseUp);
    window.addEventListener("keyup", handleKeyUp);
    window.addEventListener("scroll", handleScrollOrResize, true);
    window.addEventListener("resize", handleScrollOrResize);
    document.addEventListener("mousedown", handleDocumentMouseDown);

    return () => {
      if (timeoutId) clearTimeout(timeoutId);
      document.removeEventListener("selectionchange", handleSelectionChange);
      window.removeEventListener("mouseup", handleMouseUp);
      window.removeEventListener("keyup", handleKeyUp);
      window.removeEventListener("scroll", handleScrollOrResize, true);
      window.removeEventListener("resize", handleScrollOrResize);
      document.removeEventListener("mousedown", handleDocumentMouseDown);
    };
  }, [updatePosition, visible]);

  // Execute formatting action on the target selection range
  const handleApplyFormat = (type: InlineFormatType, value?: string | number) => {
    const range = savedRangeRef.current;
    if (!range) return;

    // Restore selection range
    const sel = window.getSelection();
    if (sel) {
      try {
        sel.removeAllRanges();
        sel.addRange(range);
      } catch (e) {}
    }

    // Apply inline format specifically targeting selection range
    applyInlineFormatting(type, value, range);

    // If target editable element is available, ensure input event fires and notify callback
    if (targetEditableRef.current) {
      targetEditableRef.current.dispatchEvent(new Event("input", { bubbles: true, cancelable: true }));
    }

    if (onFormatted) {
      onFormatted();
    }

    // Re-query range and detect formatting
    if (sel && sel.rangeCount > 0) {
      const newRange = sel.getRangeAt(0);
      savedRangeRef.current = newRange.cloneRange();
      detectActiveFormatting(newRange);
    }

    if (openDropdown) {
      setOpenDropdown(null);
    }
  };

  return (
    <AnimatePresence>
      {visible && (
        <motion.div
          ref={toolbarRef}
          id="floating-text-formatting-toolbar"
          initial={{ opacity: 0, scale: 0.92, y: placement === "top" ? 6 : -6 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.92, transition: { duration: 0.12 } }}
          transition={{ duration: 0.15, ease: "easeOut" }}
          style={{
            position: "fixed",
            top: `${position.top}px`,
            left: `${position.left}px`,
            zIndex: 9999,
          }}
          onMouseDown={(e) => {
            // Prevent toolbar clicks from stealing focus or collapsing selection
            e.preventDefault();
            e.stopPropagation();
            isInteractingRef.current = true;
          }}
          className="no-print print:hidden flex items-center bg-slate-900/95 backdrop-blur-md text-white rounded-xl shadow-2xl border border-slate-700/80 px-1.5 py-1 text-xs select-none"
        >
          {/* Bold Button */}
          <button
            type="button"
            title="Bold (Ctrl+B)"
            onClick={() => handleApplyFormat("bold")}
            onMouseDown={(e) => e.preventDefault()}
            className={`p-1.5 rounded-lg transition-colors flex items-center justify-center ${
              isBold
                ? "bg-blue-600 text-white font-bold shadow-sm"
                : "text-slate-200 hover:text-white hover:bg-slate-800"
            }`}
          >
            <Bold size={14} strokeWidth={2.6} />
          </button>

          {/* Italic Button */}
          <button
            type="button"
            title="Italic (Ctrl+I)"
            onClick={() => handleApplyFormat("italic")}
            onMouseDown={(e) => e.preventDefault()}
            className={`p-1.5 rounded-lg transition-colors flex items-center justify-center ${
              isItalic
                ? "bg-blue-600 text-white font-bold shadow-sm"
                : "text-slate-200 hover:text-white hover:bg-slate-800"
            }`}
          >
            <Italic size={14} strokeWidth={2.4} />
          </button>

          {/* Underline Button */}
          <button
            type="button"
            title="Underline (Ctrl+U)"
            onClick={() => handleApplyFormat("underline")}
            onMouseDown={(e) => e.preventDefault()}
            className={`p-1.5 rounded-lg transition-colors flex items-center justify-center ${
              isUnderline
                ? "bg-blue-600 text-white font-bold shadow-sm"
                : "text-slate-200 hover:text-white hover:bg-slate-800"
            }`}
          >
            <Underline size={14} strokeWidth={2.4} />
          </button>

          {/* Font Size Selector with Dropdown */}
          <div className="relative">
            <button
              type="button"
              title="Font Size (Selection Only)"
              onClick={() => setOpenDropdown(openDropdown === "fontSize" ? null : "fontSize")}
              onMouseDown={(e) => e.preventDefault()}
              className={`px-1.5 py-1 rounded-lg transition-colors flex items-center gap-1 ${
                openDropdown === "fontSize" ? "bg-slate-800 text-white" : "text-slate-200 hover:text-white hover:bg-slate-800"
              }`}
            >
              <span className="font-mono text-[11px] font-bold min-w-[20px] text-center">
                {currentFontSize ? `${currentFontSize}` : "11"}
              </span>
              <ChevronDown size={10} className="text-slate-400" />
            </button>

            {/* Font Size Popover */}
            {openDropdown === "fontSize" && (
              <div
                className="absolute top-full mt-1.5 left-1/2 -translate-x-1/2 bg-slate-900 border border-slate-700 rounded-xl shadow-2xl p-2 z-50 w-36"
                onMouseDown={(e) => e.preventDefault()}
              >
                <div className="flex items-center justify-between pb-1 mb-1.5 border-b border-slate-800 text-[11px] font-medium text-slate-300">
                  <span>Font Size</span>
                  <button
                    type="button"
                    onClick={() => setOpenDropdown(null)}
                    className="text-slate-400 hover:text-slate-200 p-0.5 rounded"
                  >
                    <X size={12} />
                  </button>
                </div>
                <div className="grid grid-cols-3 gap-1">
                  {[8, 9, 10, 11, 12, 14, 16, 18, 24].map((size) => (
                    <button
                      key={size}
                      type="button"
                      onClick={() => {
                        setCurrentFontSize(size);
                        handleApplyFormat("fontSize", size);
                      }}
                      onMouseDown={(e) => e.preventDefault()}
                      className={`py-1 text-center font-mono rounded text-[11px] transition-colors ${
                        currentFontSize === size
                          ? "bg-blue-600 text-white font-bold"
                          : "text-slate-200 hover:bg-slate-800"
                      }`}
                    >
                      {size}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Divider */}
          <div className="w-px h-4 bg-slate-700 mx-1" />

          {/* Text Color Button with Dropdown */}
          <div className="relative">
            <button
              type="button"
              title="Text Color"
              onClick={() => setOpenDropdown(openDropdown === "textColor" ? null : "textColor")}
              onMouseDown={(e) => e.preventDefault()}
              className={`px-1.5 py-1 rounded-lg transition-colors flex items-center gap-1 ${
                openDropdown === "textColor" ? "bg-slate-800 text-white" : "text-slate-200 hover:text-white hover:bg-slate-800"
              }`}
            >
              <div className="flex flex-col items-center justify-center">
                <Baseline size={13} strokeWidth={2.2} />
                <div
                  className="w-3.5 h-1 rounded-full mt-0.5"
                  style={{ backgroundColor: currentTextColor || "#000000" }}
                />
              </div>
              <ChevronDown size={10} className="text-slate-400" />
            </button>

            {/* Text Color Popover */}
            {openDropdown === "textColor" && (
              <div
                className="absolute top-full mt-1.5 left-1/2 -translate-x-1/2 bg-slate-900 border border-slate-700 rounded-xl shadow-2xl p-2.5 z-50 w-52"
                onMouseDown={(e) => e.preventDefault()}
              >
                <div className="flex items-center justify-between pb-1.5 mb-1.5 border-b border-slate-800 text-[11px] font-medium text-slate-300">
                  <span>Text Color</span>
                  <button
                    type="button"
                    onClick={() => setOpenDropdown(null)}
                    className="text-slate-400 hover:text-slate-200 p-0.5 rounded"
                  >
                    <X size={12} />
                  </button>
                </div>

                <div className="grid grid-cols-5 gap-1.5">
                  {TEXT_COLORS.map((c) => (
                    <button
                      key={c.value}
                      type="button"
                      title={c.name}
                      onClick={() => {
                        setCurrentTextColor(c.value);
                        handleApplyFormat("fontColor", c.value);
                      }}
                      onMouseDown={(e) => e.preventDefault()}
                      className="w-7 h-7 rounded-lg border border-slate-700 hover:scale-110 active:scale-95 transition-transform flex items-center justify-center relative shadow-sm"
                      style={{ backgroundColor: c.value }}
                    >
                      {currentTextColor === c.value && (
                        <Check size={12} className={c.value === "#000000" || c.value === "#1E40AF" || c.value === "#991B1B" ? "text-white" : "text-black"} />
                      )}
                    </button>
                  ))}
                </div>

                {/* Custom Color Input */}
                <div className="mt-2 pt-2 border-t border-slate-800 flex items-center justify-between text-[11px]">
                  <span className="text-slate-400">Custom Color:</span>
                  <label className="flex items-center gap-1.5 cursor-pointer bg-slate-800 hover:bg-slate-700 px-2 py-0.5 rounded-md border border-slate-700">
                    <div
                      className="w-3.5 h-3.5 rounded border border-white/40 shadow-inner"
                      style={{ backgroundColor: currentTextColor }}
                    />
                    <span className="text-[10px] text-slate-300 font-mono">Pick</span>
                    <input
                      type="color"
                      value={currentTextColor.startsWith("#") ? currentTextColor : "#000000"}
                      onChange={(e) => {
                        const val = e.target.value;
                        setCurrentTextColor(val);
                        handleApplyFormat("fontColor", val);
                      }}
                      className="sr-only"
                    />
                  </label>
                </div>
              </div>
            )}
          </div>

          {/* Highlighter Button with Dropdown */}
          <div className="relative">
            <button
              type="button"
              title="Text Highlighter"
              onClick={() => setOpenDropdown(openDropdown === "highlight" ? null : "highlight")}
              onMouseDown={(e) => e.preventDefault()}
              className={`px-1.5 py-1 rounded-lg transition-colors flex items-center gap-1 ${
                openDropdown === "highlight" ? "bg-slate-800 text-white" : "text-slate-200 hover:text-white hover:bg-slate-800"
              }`}
            >
              <div className="flex flex-col items-center justify-center">
                <Highlighter size={13} strokeWidth={2.2} />
                <div
                  className="w-3.5 h-1 rounded-full mt-0.5"
                  style={{
                    backgroundColor: currentHighlightColor === "transparent" ? "#64748B" : currentHighlightColor || "#FEF08A",
                  }}
                />
              </div>
              <ChevronDown size={10} className="text-slate-400" />
            </button>

            {/* Highlight Color Popover */}
            {openDropdown === "highlight" && (
              <div
                className="absolute top-full mt-1.5 left-1/2 -translate-x-1/2 bg-slate-900 border border-slate-700 rounded-xl shadow-2xl p-2.5 z-50 w-52"
                onMouseDown={(e) => e.preventDefault()}
              >
                <div className="flex items-center justify-between pb-1.5 mb-1.5 border-b border-slate-800 text-[11px] font-medium text-slate-300">
                  <span>Highlighter</span>
                  <button
                    type="button"
                    onClick={() => setOpenDropdown(null)}
                    className="text-slate-400 hover:text-slate-200 p-0.5 rounded"
                  >
                    <X size={12} />
                  </button>
                </div>

                <div className="grid grid-cols-5 gap-1.5">
                  {HIGHLIGHT_COLORS.map((h) => {
                    const isNone = h.value === "transparent";
                    return (
                      <button
                        key={h.value}
                        type="button"
                        title={h.name}
                        onClick={() => {
                          setCurrentHighlightColor(h.value);
                          handleApplyFormat("highlight", h.value);
                        }}
                        onMouseDown={(e) => e.preventDefault()}
                        className={`w-7 h-7 rounded-lg border border-slate-700 hover:scale-110 active:scale-95 transition-transform flex items-center justify-center relative shadow-sm ${
                          isNone ? "bg-slate-800" : ""
                        }`}
                        style={{ backgroundColor: isNone ? undefined : h.value }}
                      >
                        {isNone ? (
                          <div className="w-full h-full flex items-center justify-center text-slate-400">
                            <span className="text-[10px] font-bold">∅</span>
                          </div>
                        ) : (
                          currentHighlightColor === h.value && (
                            <Check size={12} className="text-slate-900" />
                          )
                        )}
                      </button>
                    );
                  })}
                </div>

                {/* Custom Highlight Picker */}
                <div className="mt-2 pt-2 border-t border-slate-800 flex items-center justify-between text-[11px]">
                  <span className="text-slate-400">Custom Fill:</span>
                  <label className="flex items-center gap-1.5 cursor-pointer bg-slate-800 hover:bg-slate-700 px-2 py-0.5 rounded-md border border-slate-700">
                    <div
                      className="w-3.5 h-3.5 rounded border border-white/40 shadow-inner"
                      style={{ backgroundColor: currentHighlightColor === "transparent" ? "#FEF08A" : currentHighlightColor }}
                    />
                    <span className="text-[10px] text-slate-300 font-mono">Pick</span>
                    <input
                      type="color"
                      value={currentHighlightColor.startsWith("#") ? currentHighlightColor : "#FEF08A"}
                      onChange={(e) => {
                        const val = e.target.value;
                        setCurrentHighlightColor(val);
                        handleApplyFormat("highlight", val);
                      }}
                      className="sr-only"
                    />
                  </label>
                </div>
              </div>
            )}
          </div>

          {/* Divider */}
          <div className="w-px h-4 bg-slate-700 mx-1" />

          {/* Clear / Reset Formatting on Selection */}
          <button
            type="button"
            title="Clear Formatting on Selection"
            onClick={() => handleApplyFormat("clearFormat")}
            onMouseDown={(e) => e.preventDefault()}
            className="p-1.5 rounded-lg text-slate-300 hover:text-red-400 hover:bg-slate-800 transition-colors flex items-center justify-center"
          >
            <RemoveFormatting size={14} strokeWidth={2.2} />
          </button>
        </motion.div>
      )}
    </AnimatePresence>
  );
};

export default FloatingTextToolbar;
