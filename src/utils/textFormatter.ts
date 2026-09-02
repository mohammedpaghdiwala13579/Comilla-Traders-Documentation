/**
 * Robust utility for applying inline text formatting (Text Color, Highlight / Fill Color, Bold, Italic, Underline, Font Size, Font Family, Reset)
 * to specific words, phrases, or text selections in contenteditable cells and document fields.
 */

// Memory store for the last active selection across the document
let savedSelection: {
  element: HTMLElement;
  range: Range;
  text: string;
} | null = null;

/**
 * Saves the current text selection if it belongs to a contenteditable element
 */
export function saveCurrentSelection(): boolean {
  if (typeof window === "undefined") return false;
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return false;

  const range = sel.getRangeAt(0);
  const commonAncestor = range.commonAncestorContainer;
  const editableEl = commonAncestor instanceof HTMLElement
    ? commonAncestor.closest<HTMLElement>("[contenteditable='true']")
    : commonAncestor.parentElement?.closest<HTMLElement>("[contenteditable='true']");

  if (editableEl) {
    savedSelection = {
      element: editableEl,
      range: range.cloneRange(),
      text: sel.toString(),
    };
    return true;
  }
  return false;
}

/**
 * Returns the currently saved selection or active selection
 */
export function getSavedSelection() {
  return savedSelection;
}

/**
 * Restores the saved selection and focuses the editable element
 */
export function restoreSavedSelection(): boolean {
  if (typeof window === "undefined" || !savedSelection) return false;
  try {
    const sel = window.getSelection();
    if (!sel) return false;

    savedSelection.element.focus();
    sel.removeAllRanges();
    sel.addRange(savedSelection.range);
    return true;
  } catch (e) {
    return false;
  }
}

// Automatically save selection on selectionchange, mouseup, and keyup across the page
if (typeof document !== "undefined") {
  document.addEventListener("selectionchange", () => {
    saveCurrentSelection();
  });
}

/**
 * Finds the word boundaries around the cursor in a text node
 */
function expandRangeToWordAtCursor(sel: Selection, range: Range): Range | null {
  if (!range.collapsed) return range;

  let node = range.startContainer;
  let offset = range.startOffset;

  // If container is element, try to get child text node
  if (node.nodeType === Node.ELEMENT_NODE && node.childNodes.length > 0) {
    const child = node.childNodes[Math.min(offset, node.childNodes.length - 1)];
    if (child && child.nodeType === Node.TEXT_NODE) {
      node = child;
      offset = 0;
    } else {
      return null;
    }
  }

  if (node.nodeType !== Node.TEXT_NODE) return null;

  const text = node.textContent || "";
  if (!text || text.trim().length === 0) return null;

  // Scan backwards to find start of word
  let start = offset;
  while (start > 0 && !/\s|[.,;!?'"()\[\]{}]/.test(text.charAt(start - 1))) {
    start--;
  }

  // Scan forwards to find end of word
  let end = offset;
  while (end < text.length && !/\s|[.,;!?'"()\[\]{}]/.test(text.charAt(end))) {
    end++;
  }

  if (start >= end) return null;

  try {
    const newRange = document.createRange();
    newRange.setStart(node, start);
    newRange.setEnd(node, end);
    sel.removeAllRanges();
    sel.addRange(newRange);
    return newRange;
  } catch (e) {
    return null;
  }
}

export type InlineFormatType =
  | "highlight"
  | "hiliteColor"
  | "fontColor"
  | "foreColor"
  | "bold"
  | "italic"
  | "underline"
  | "fontSize"
  | "fontFamily"
  | "clearFormat"
  | "resetWord";

/**
 * Applies inline formatting specifically to the selected text or word under cursor
 */
export function applyInlineFormatting(
  type: InlineFormatType,
  value?: string | number
): boolean {
  if (typeof window === "undefined") return false;

  let sel = window.getSelection();

  // If selection is absent or outside an editable container, attempt to restore saved selection
  const isCurrentlyInEditable = () => {
    if (!sel || sel.rangeCount === 0) return false;
    const anc = sel.getRangeAt(0).commonAncestorContainer;
    const el = anc instanceof HTMLElement ? anc : anc.parentElement;
    return !!el?.closest("[contenteditable='true']");
  };

  if (!isCurrentlyInEditable()) {
    restoreSavedSelection();
    sel = window.getSelection();
  }

  if (!sel || sel.rangeCount === 0) {
    return false;
  }

  let range = sel.getRangeAt(0);
  const commonAncestor = range.commonAncestorContainer;
  const editableElement = commonAncestor instanceof HTMLElement
    ? commonAncestor.closest<HTMLElement>("[contenteditable='true']")
    : commonAncestor.parentElement?.closest<HTMLElement>("[contenteditable='true']");

  if (!editableElement) {
    return false;
  }

  // If the selection is collapsed (cursor at position), auto-expand to the specific word under cursor!
  if (range.collapsed) {
    const expanded = expandRangeToWordAtCursor(sel, range);
    if (expanded) {
      range = expanded;
    }
  }

  // If still collapsed or empty, nothing to format inline
  if (range.collapsed && type !== "highlight" && type !== "fontColor") {
    return false;
  }

  // Ensure styleWithCSS is true so modern inline CSS styles are generated
  try {
    document.execCommand("styleWithCSS", false, "true");
  } catch (e) {
    // Ignore legacy quirks
  }

  const normalizedType = type === "hiliteColor" ? "highlight" : type === "foreColor" ? "fontColor" : type;

  // 1. HIGHLIGHT / BACKGROUND COLOR
  if (normalizedType === "highlight") {
    const isTransparent = !value || value === "transparent" || value === "none" || value === "inherit" || value === "";
    
    if (isTransparent) {
      // Remove highlight from the selected word / range
      removeStyleFromRange(range, editableElement, "backgroundColor");
    } else {
      const colorVal = String(value);
      // Modern browsers: backColor produces <span style="background-color: ...">
      const ok = document.execCommand("backColor", false, colorVal);
      if (!ok) {
        wrapRangeWithStyle(range, { backgroundColor: colorVal });
      }
    }
  }

  // 2. FONT / TEXT COLOR
  else if (normalizedType === "fontColor") {
    const isDefault = !value || value === "inherit" || value === "automatic" || value === "auto" || value === "#000000";
    if (isDefault) {
      removeStyleFromRange(range, editableElement, "color");
    } else {
      const colorVal = String(value);
      const ok = document.execCommand("foreColor", false, colorVal);
      if (!ok) {
        wrapRangeWithStyle(range, { color: colorVal });
      }
    }
  }

  // 3. BOLD
  else if (normalizedType === "bold") {
    document.execCommand("bold", false, undefined);
  }

  // 4. ITALIC
  else if (normalizedType === "italic") {
    document.execCommand("italic", false, undefined);
  }

  // 5. UNDERLINE
  else if (normalizedType === "underline") {
    if (value === "none") {
      removeStyleFromRange(range, editableElement, "textDecoration");
    } else if (value === "double") {
      wrapRangeWithStyle(range, { textDecoration: "underline", textDecorationStyle: "double" });
    } else {
      document.execCommand("underline", false, undefined);
    }
  }

  // 6. FONT SIZE
  else if (normalizedType === "fontSize") {
    const sizeStr = typeof value === "number" ? `${value}pt` : String(value || "11pt");
    wrapRangeWithStyle(range, { fontSize: sizeStr.endsWith("pt") || sizeStr.endsWith("px") ? sizeStr : `${sizeStr}pt` });
  }

  // 7. FONT FAMILY
  else if (normalizedType === "fontFamily") {
    const familyStr = String(value || "Arial, sans-serif");
    wrapRangeWithStyle(range, { fontFamily: familyStr });
  }

  // 8. CLEAR FORMATTING / RESET SPECIFIC WORD
  else if (normalizedType === "clearFormat" || normalizedType === "resetWord") {
    resetFormattingInRange(range, editableElement);
  }

  // Update saved selection after formatting
  saveCurrentSelection();

  // Dispatch an input event so React immediately captures the innerHTML state update
  editableElement.dispatchEvent(new Event("input", { bubbles: true, cancelable: true }));

  return true;
}

/**
 * Helper to wrap a DOM range with a span containing specific inline styles
 */
function wrapRangeWithStyle(range: Range, styles: Record<string, string>) {
  if (range.collapsed) return;

  try {
    const selectedContent = range.extractContents();
    const span = document.createElement("span");
    
    Object.entries(styles).forEach(([prop, val]) => {
      // @ts-ignore
      span.style[prop] = val;
    });

    span.appendChild(selectedContent);
    range.insertNode(span);

    // Re-select the wrapped span so the user keeps their selection
    const sel = window.getSelection();
    if (sel) {
      const newRange = document.createRange();
      newRange.selectNodeContents(span);
      sel.removeAllRanges();
      sel.addRange(newRange);
    }
  } catch (e) {
    console.warn("Could not wrap range with style:", e);
  }
}

/**
 * Removes a specific CSS style property (like backgroundColor or color) from within a range
 */
function removeStyleFromRange(range: Range, container: HTMLElement, styleProp: "backgroundColor" | "color" | "textDecoration") {
  try {
    // Find any styled spans intersecting or containing the range
    const walker = document.createTreeWalker(container, NodeFilter.SHOW_ELEMENT, null);
    const nodesToClean: HTMLElement[] = [];

    let curr = walker.nextNode();
    while (curr) {
      if (curr instanceof HTMLElement && range.intersectsNode(curr)) {
        if (curr.style[styleProp]) {
          nodesToClean.push(curr);
        }
      }
      curr = walker.nextNode();
    }

    nodesToClean.forEach((el) => {
      el.style[styleProp] = "";
      // If element has no styles and is a span, unwrap it
      if (!el.getAttribute("style") && el.tagName.toLowerCase() === "span") {
        el.replaceWith(...Array.from(el.childNodes));
      }
    });

    // Also run backColor transparent command as browser fallback
    if (styleProp === "backgroundColor") {
      try {
        document.execCommand("backColor", false, "transparent");
      } catch (e) {}
    }
  } catch (e) {
    console.warn("Could not remove style from range:", e);
  }
}

/**
 * Strips all inline formatting tags (spans, b, i, u, font, mark) within the range, resetting to plain text
 */
function resetFormattingInRange(range: Range, container: HTMLElement) {
  try {
    const fragment = range.extractContents();
    const temp = document.createElement("div");
    temp.appendChild(fragment);

    // Extract text content only
    const plainText = temp.textContent || "";
    const textNode = document.createTextNode(plainText);
    range.insertNode(textNode);

    const sel = window.getSelection();
    if (sel) {
      const newRange = document.createRange();
      newRange.selectNode(textNode);
      sel.removeAllRanges();
      sel.addRange(newRange);
    }
  } catch (e) {
    // Fallback: document.execCommand removeFormat
    try {
      document.execCommand("removeFormat", false, undefined);
    } catch (err) {}
  }
}

/**
 * Strips HTML tags to return plain text (useful for Excel export and calculations)
 */
export function stripHtml(html: string): string {
  if (!html) return "";
  if (!html.includes("<") && !html.includes("&")) return html;
  
  if (typeof DOMParser !== "undefined") {
    const doc = new DOMParser().parseFromString(html, "text/html");
    return doc.body.textContent || "";
  }
  return html.replace(/<[^>]*>/g, "");
}

/**
 * Parses numeric inputs safely supporting negative numbers (-500, −500, (500)),
 * commas (1,234.50), rich-text HTML values, and currency symbols.
 */
export function parseNumericInput(val: any): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const rawStr = String(val).trim();
  if (!rawStr) return 0;

  const clean = stripHtml(rawStr).trim();
  if (!clean) return 0;

  // Handle accounting negative format: (500) or ( 500.50 )
  const accountingMatch = clean.match(/^\s*\(\s*([\d,]+(?:\.\d+)?)\s*\)\s*$/);
  if (accountingMatch) {
    const inner = accountingMatch[1].replace(/,/g, "");
    const parsed = parseFloat(inner);
    return isNaN(parsed) ? 0 : -Math.abs(parsed);
  }

  // Handle leading negative sign (- or unicode −)
  const isNegative = clean.startsWith("-") || clean.startsWith("−");
  const sanitized = clean
    .replace(/^[−-]/, "")
    .replace(/,/g, "")
    .replace(/\s+/g, "");

  const parsed = parseFloat(sanitized);
  if (isNaN(parsed)) return 0;
  return isNegative ? -Math.abs(parsed) : parsed;
}
