/**
 * Robust utility for applying inline text formatting (Text Color, Highlight / Fill Color, Bold, Italic, Underline)
 * to specific text selections in contenteditable cells and document fields.
 */

export function applyInlineFormatting(
  type: "foreColor" | "hiliteColor" | "bold" | "italic" | "underline",
  value?: string
): boolean {
  if (typeof window === "undefined") return false;

  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) {
    return false;
  }

  const selectedText = selection.toString();
  if (!selectedText || selectedText.length === 0) {
    return false;
  }

  // Ensure styleWithCSS is true so modern spans with CSS styles are generated
  try {
    document.execCommand("styleWithCSS", false, "true");
  } catch (e) {
    // Ignore legacy browser quirks
  }

  if (type === "foreColor") {
    document.execCommand("foreColor", false, value || "#000000");
  } else if (type === "hiliteColor") {
    if (!value || value === "transparent" || value === "none") {
      document.execCommand("removeFormat", false, undefined);
    } else {
      const ok = document.execCommand("hiliteColor", false, value);
      if (!ok) {
        document.execCommand("backColor", false, value);
      }
    }
  } else if (type === "bold" || type === "italic" || type === "underline") {
    document.execCommand(type, false, undefined);
  }

  // Dispatch an input event to the closest contentEditable parent so React captures the state update
  const anchorNode = selection.anchorNode;
  const editableElement = anchorNode instanceof HTMLElement
    ? anchorNode.closest("[contenteditable='true']")
    : anchorNode?.parentElement?.closest("[contenteditable='true']");

  if (editableElement) {
    const inputEvent = new Event("input", { bubbles: true, cancelable: true });
    editableElement.dispatchEvent(inputEvent);
  }

  return true;
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
