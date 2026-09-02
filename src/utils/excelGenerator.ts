import ExcelJS from "exceljs";
import { QuotationRow, MergedRegion, CellFormatMap, CellFormat, CellBorders } from "../types";
import { numberToWords } from "./numberToWords";
import { cleanCellText } from "./tsvParser";
import { parseNumericInput, stripHtml } from "./textFormatter";

/**
 * Converts CSS color (hex, rgb, rgba, named colors) to 8-character ARGB for ExcelJS
 */
export function colorToArgb(color: string | null | undefined): string | null {
  if (!color) return null;
  const c = color.trim().toLowerCase();
  if (c === "transparent" || c === "none" || c === "initial" || c === "inherit" || c === "") return null;

  // Hex format #RGB, #RRGGBB, #RRGGBBAA
  if (c.startsWith("#")) {
    const hex = c.slice(1);
    if (hex.length === 3) {
      return (
        "FF" +
        hex[0] + hex[0] +
        hex[1] + hex[1] +
        hex[2] + hex[2]
      ).toUpperCase();
    }
    if (hex.length === 6) {
      return ("FF" + hex).toUpperCase();
    }
    if (hex.length === 8) {
      return (hex.slice(6, 8) + hex.slice(0, 6)).toUpperCase();
    }
  }

  // rgb(r, g, b) or rgba(r, g, b, a)
  const rgbMatch = c.match(/^rgba?\(\s*([0-9]+)\s*,\s*([0-9]+)\s*,\s*([0-9]+)(?:\s*,\s*([0-9.]+))?\s*\)$/);
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1], 10);
    const g = parseInt(rgbMatch[2], 10);
    const b = parseInt(rgbMatch[3], 10);
    const a = rgbMatch[4] !== undefined ? parseFloat(rgbMatch[4]) : 1;
    if (a <= 0) return null;
    const aHex = Math.round(a * 255).toString(16).padStart(2, "0");
    const rHex = Math.max(0, Math.min(255, r)).toString(16).padStart(2, "0");
    const gHex = Math.max(0, Math.min(255, g)).toString(16).padStart(2, "0");
    const bHex = Math.max(0, Math.min(255, b)).toString(16).padStart(2, "0");
    return (aHex + rHex + gHex + bHex).toUpperCase();
  }

  const NAMED_COLORS: Record<string, string> = {
    black: "FF000000",
    white: "FFFFFFFF",
    red: "FFFF0000",
    green: "FF008000",
    blue: "FF0000FF",
    yellow: "FFFFFF00",
    amber: "FFF59E0B",
    gray: "FF808080",
    grey: "FF808080",
    slate: "FF64748B",
    orange: "FFF97316",
    purple: "FFA855F7",
    teal: "FF14B8A6",
    cyan: "FF06B6D4",
  };
  if (NAMED_COLORS[c]) {
    return NAMED_COLORS[c];
  }

  if (/^[0-9a-f]{6}$/i.test(c)) {
    return ("FF" + c).toUpperCase();
  }
  if (/^[0-9a-f]{8}$/i.test(c)) {
    return c.toUpperCase();
  }

  return null;
}

export type ExcelBorderDef = Partial<ExcelJS.Border>;

/**
 * Parses CSS border string (e.g. "1px solid black", "3px double black", "none") into ExcelJS Border
 */
export function parseBorderSide(
  borderStr?: string,
  defaultStyle?: ExcelJS.BorderStyle
): ExcelBorderDef | undefined {
  if (borderStr === undefined) {
    if (!defaultStyle) return undefined;
    return { style: defaultStyle, color: { argb: "FF000000" } };
  }
  const s = borderStr.trim().toLowerCase();
  if (s === "none" || s === "0" || s === "0px" || s === "hidden" || s === "") {
    return undefined;
  }

  let style: ExcelJS.BorderStyle = "thin";
  if (s.includes("double")) {
    style = "double";
  } else if (s.includes("3px") || s.includes("thick")) {
    style = "thick";
  } else if (s.includes("2px") || s.includes("medium")) {
    style = "medium";
  } else if (s.includes("dotted")) {
    style = "dotted";
  } else if (s.includes("dashed")) {
    style = "dashed";
  } else {
    style = "thin";
  }

  let argb = "FF000000";
  const hexMatch = s.match(/#([0-9a-f]{3,8})/i);
  if (hexMatch) {
    const parsed = colorToArgb(hexMatch[0]);
    if (parsed) argb = parsed;
  } else {
    const rgbMatch = s.match(/rgba?\([^)]+\)/i);
    if (rgbMatch) {
      const parsed = colorToArgb(rgbMatch[0]);
      if (parsed) argb = parsed;
    }
  }

  return { style, color: { argb } };
}

/**
 * Converts HTML into plain text while preserving intentional line breaks (<br>, </p>, </div>)
 */
export const htmlToPlainText = (html: string): string => {
  if (!html) return "";
  if (!html.includes("<") && !html.includes("&")) return html;
  const replaced = html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/&nbsp;/gi, " ");

  if (typeof DOMParser !== "undefined") {
    try {
      const doc = new DOMParser().parseFromString(replaced, "text/html");
      return (doc.body.textContent || "").replace(/\u00A0/g, " ");
    } catch (e) {}
  }
  return replaced.replace(/<[^>]*>/g, "").replace(/\u00A0/g, " ");
};

interface TextRunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | "single" | "double";
  strike?: boolean;
  color?: string; // ARGB
  bgColor?: string; // ARGB
  size?: number; // pt
  fontFamily?: string;
}

export interface ParsedHtmlResult {
  richText: ExcelJS.RichText[];
  plainText: string;
  hasFormatting: boolean;
  highlightColor?: string; // ARGB
}

/**
 * Parses an HTML string (from contenteditable RichTextCell) into ExcelJS richText runs
 * preserving specific word highlight, color, size, bold, italic, underline, and fonts.
 */
export function parseHtmlToExcelRuns(
  htmlOrText: string,
  baseFont: Partial<ExcelJS.Font> = {}
): ParsedHtmlResult {
  if (!htmlOrText) {
    return { richText: [], plainText: "", hasFormatting: false };
  }

  // If no HTML tags at all, return standard plain text
  if (!htmlOrText.includes("<")) {
    return {
      richText: [
        {
          text: htmlOrText,
          font: { ...baseFont },
        },
      ],
      plainText: htmlOrText,
      hasFormatting: false,
    };
  }

  if (typeof DOMParser === "undefined") {
    const clean = stripHtml(htmlOrText);
    return {
      richText: [{ text: clean, font: { ...baseFont } }],
      plainText: clean,
      hasFormatting: false,
    };
  }

  const normalizedHtml = htmlOrText
    .replace(/<br\s*\/?>/gi, "<br>")
    .replace(/&nbsp;/gi, " ");

  const doc = new DOMParser().parseFromString(`<div>${normalizedHtml}</div>`, "text/html");
  const root = doc.body.firstElementChild || doc.body;

  const rawRuns: { text: string; style: TextRunStyle }[] = [];
  let foundHighlight: string | undefined = undefined;
  let hasSpecialFormatting = false;

  const traverse = (node: Node, currentStyle: TextRunStyle) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent || "";
      if (text.length > 0) {
        rawRuns.push({ text, style: { ...currentStyle } });
      }
      return;
    }

    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as HTMLElement;
      const tag = el.tagName.toUpperCase();

      if (tag === "BR") {
        rawRuns.push({ text: "\n", style: { ...currentStyle } });
        return;
      }

      const nextStyle: TextRunStyle = { ...currentStyle };

      // HTML Tags
      if (tag === "B" || tag === "STRONG") {
        nextStyle.bold = true;
        hasSpecialFormatting = true;
      }
      if (tag === "I" || tag === "EM") {
        nextStyle.italic = true;
        hasSpecialFormatting = true;
      }
      if (tag === "U") {
        nextStyle.underline = true;
        hasSpecialFormatting = true;
      }
      if (tag === "S" || tag === "STRIKE" || tag === "DEL") {
        nextStyle.strike = true;
        hasSpecialFormatting = true;
      }
      if (tag === "MARK") {
        const bg = colorToArgb(el.style.backgroundColor || "#fef08a");
        if (bg) {
          nextStyle.bgColor = bg;
          foundHighlight = bg;
          hasSpecialFormatting = true;
        }
      }
      if (tag === "FONT") {
        const fontColor = el.getAttribute("color");
        if (fontColor) {
          const argb = colorToArgb(fontColor);
          if (argb) {
            nextStyle.color = argb;
            hasSpecialFormatting = true;
          }
        }
        const face = el.getAttribute("face");
        if (face) {
          nextStyle.fontFamily = face;
          hasSpecialFormatting = true;
        }
        const sizeAttr = el.getAttribute("size");
        if (sizeAttr) {
          const num = parseInt(sizeAttr, 10);
          if (!isNaN(num)) {
            const sizeMap = [8, 9, 10, 11, 14, 18, 24, 36];
            nextStyle.size = sizeMap[Math.min(num, sizeMap.length - 1)];
            hasSpecialFormatting = true;
          }
        }
      }

      // Inline Styles
      if (el.style) {
        if (el.style.fontWeight) {
          const fw = el.style.fontWeight.toLowerCase();
          if (fw === "bold" || fw === "700" || fw === "800" || fw === "900" || fw === "bolder") {
            nextStyle.bold = true;
            hasSpecialFormatting = true;
          } else if (fw === "normal" || fw === "400") {
            nextStyle.bold = false;
          }
        }

        if (el.style.fontStyle) {
          const fs = el.style.fontStyle.toLowerCase();
          if (fs === "italic" || fs === "oblique") {
            nextStyle.italic = true;
            hasSpecialFormatting = true;
          } else if (fs === "normal") {
            nextStyle.italic = false;
          }
        }

        if (el.style.textDecoration || el.style.textDecorationLine) {
          const td = (el.style.textDecoration || el.style.textDecorationLine).toLowerCase();
          if (td.includes("underline")) {
            nextStyle.underline = td.includes("double") || el.style.textDecorationStyle === "double" ? "double" : true;
            hasSpecialFormatting = true;
          } else if (td.includes("line-through")) {
            nextStyle.strike = true;
            hasSpecialFormatting = true;
          } else if (td === "none") {
            nextStyle.underline = false;
            nextStyle.strike = false;
          }
        }

        if (el.style.color) {
          const argb = colorToArgb(el.style.color);
          if (argb) {
            nextStyle.color = argb;
            hasSpecialFormatting = true;
          }
        }

        if (el.style.backgroundColor || el.style.background) {
          const bg = colorToArgb(el.style.backgroundColor || el.style.background);
          if (bg) {
            nextStyle.bgColor = bg;
            foundHighlight = bg;
            hasSpecialFormatting = true;
          }
        }

        if (el.style.fontSize) {
          const fsStr = el.style.fontSize.trim().toLowerCase();
          if (fsStr.endsWith("pt")) {
            nextStyle.size = parseFloat(fsStr);
            hasSpecialFormatting = true;
          } else if (fsStr.endsWith("px")) {
            nextStyle.size = Math.round(parseFloat(fsStr) * 0.75);
            hasSpecialFormatting = true;
          } else {
            const num = parseFloat(fsStr);
            if (!isNaN(num) && num > 0) {
              nextStyle.size = num;
              hasSpecialFormatting = true;
            }
          }
        }

        if (el.style.fontFamily) {
          const cleaned = el.style.fontFamily.replace(/['"]/g, "").split(",")[0].trim();
          if (cleaned) {
            nextStyle.fontFamily = cleaned;
            hasSpecialFormatting = true;
          }
        }
      }

      const isBlock = tag === "DIV" || tag === "P" || tag === "TR";
      if (isBlock && rawRuns.length > 0 && !rawRuns[rawRuns.length - 1].text.endsWith("\n")) {
        rawRuns.push({ text: "\n", style: { ...currentStyle } });
      }

      for (let i = 0; i < node.childNodes.length; i++) {
        traverse(node.childNodes[i], nextStyle);
      }

      if (isBlock && rawRuns.length > 0 && !rawRuns[rawRuns.length - 1].text.endsWith("\n")) {
        rawRuns.push({ text: "\n", style: { ...currentStyle } });
      }
    }
  };

  const baseArgb = baseFont.color?.argb ? String(baseFont.color.argb) : undefined;
  traverse(root, {
    bold: baseFont.bold,
    italic: baseFont.italic,
    underline: baseFont.underline === "double" ? "double" : !!baseFont.underline,
    color: baseArgb,
    size: baseFont.size,
    fontFamily: baseFont.name,
  });

  // Consolidate adjacent runs with identical styling
  const consolidatedRuns: { text: string; style: TextRunStyle }[] = [];
  for (const run of rawRuns) {
    if (!run.text) continue;
    if (consolidatedRuns.length === 0) {
      consolidatedRuns.push({ ...run });
    } else {
      const prev = consolidatedRuns[consolidatedRuns.length - 1];
      const sameStyle =
        prev.style.bold === run.style.bold &&
        prev.style.italic === run.style.italic &&
        prev.style.underline === run.style.underline &&
        prev.style.strike === run.style.strike &&
        prev.style.color === run.style.color &&
        prev.style.bgColor === run.style.bgColor &&
        prev.style.size === run.style.size &&
        prev.style.fontFamily === run.style.fontFamily;
      if (sameStyle) {
        prev.text += run.text;
      } else {
        consolidatedRuns.push({ ...run });
      }
    }
  }

  // Convert to ExcelJS.RichText
  const richText: ExcelJS.RichText[] = consolidatedRuns.map((r) => {
    const font: Partial<ExcelJS.Font> = {
      name: r.style.fontFamily || baseFont.name || "Arial",
      size: r.style.size !== undefined ? r.style.size : (baseFont.size || 8),
      bold: r.style.bold !== undefined ? r.style.bold : baseFont.bold,
      italic: r.style.italic !== undefined ? r.style.italic : baseFont.italic,
      underline: r.style.underline !== undefined ? (r.style.underline === "double" ? "double" : !!r.style.underline) : baseFont.underline,
      strike: r.style.strike !== undefined ? r.style.strike : baseFont.strike,
      color: r.style.color ? { argb: r.style.color } : (baseFont.color || { argb: "FF000000" }),
    };

    return {
      text: r.text,
      font,
    };
  });

  const plainText = consolidatedRuns.map((r) => r.text).join("");

  return {
    richText,
    plainText,
    hasFormatting: hasSpecialFormatting,
    highlightColor: foundHighlight,
  };
}

const getBase64Image = async (url: string): Promise<{ base64: string; ext: string } | null> => {
  try {
    const res = await fetch(url, { referrerPolicy: "no-referrer" });
    if (!res.ok) return null;
    const blob = await res.blob();
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64 = reader.result as string;
        const matches = base64.match(/^data:image\/([a-zA-Z+]+);base64,(.+)$/);
        if (matches && matches.length === 3) {
          resolve({ ext: matches[1], base64: matches[2] });
        } else {
          resolve(null);
        }
      };
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch (e) {
    console.warn("Could not load image for Excel:", url, e);
    return null;
  }
};

/**
 * Calculates visual text lines in an item's description based on column width & line breaks.
 */
export const calculateItemVisualLines = (desc: string, isChallan: boolean): number => {
  if (!desc) return 1;
  const plain = htmlToPlainText(desc).trim();
  if (!plain) return 1;
  const maxCharsPerLine = isChallan ? 54 : 46;
  const parts = plain.split(/\r?\n/);
  let lines = 0;
  for (const part of parts) {
    const trimmed = part.trim();
    if (trimmed.length === 0) {
      lines += 1;
    } else {
      lines += Math.max(1, Math.ceil(trimmed.length / maxCharsPerLine));
    }
  }
  return Math.max(1, lines);
};

/**
 * Computes tight, compact Excel row height (in points) strictly proportional to text content.
 * Prevents overdisplaying and ensures text is fully visible with no clipping or unnecessary white space.
 */
export const getItemRowHeight = (visualLines: number): number => {
  if (visualLines <= 1) return 16.0;
  if (visualLines === 2) return 25.5;
  if (visualLines === 3) return 35.0;
  return 16.0 + (visualLines - 1) * 9.5;
};

export interface ExcelPageChunk {
  rows: QuotationRow[];
  startSlIndex: number;
  isLastPage: boolean;
}

/**
 * Dynamically paginates items to maximize space utilization on each A4 page.
 * Uses exact vertical height budgets based on item text.
 * When the page reaches its maximum vertical capacity, remaining items are transferred to the next page.
 */
export const paginateRowsForExcel = (
  allRows: QuotationRow[],
  docType: "quotation" | "challan" | "invoice"
): ExcelPageChunk[] => {
  if (allRows.length === 0) {
    return [{ rows: [], startSlIndex: 1, isLastPage: true }];
  }

  const isChallan = docType === "challan";
  const isInvoice = docType === "invoice";
  
  // Maximum usable vertical height budget for items on A4 page (in points)
  // Header: Row 1 (22pt) + 10-row gap (140pt) + Metadata (67.5pt) + Table Header (18pt) = 247.5pt.
  // Last page: Available accounts for Totals and Signature/Stamp blocks.
  const REGULAR_PAGE_BUDGET = 515; 
  const LAST_PAGE_BUDGET = isChallan ? 515 : isInvoice ? 405 : 445;

  const chunks: ExcelPageChunk[] = [];
  let currentChunk: QuotationRow[] = [];
  let currentHeight = 0;
  let currentSl = 1;

  for (let i = 0; i < allRows.length; i++) {
    const row = allRows[i];
    const lines = calculateItemVisualLines(row.desc, isChallan);
    const rowHeight = getItemRowHeight(lines);

    const isPotentialLastItem = i === allRows.length - 1;
    const budgetForCurrentPage = isPotentialLastItem ? LAST_PAGE_BUDGET : REGULAR_PAGE_BUDGET;

    // Check if adding this item exceeds the maximum A4 vertical budget
    if (currentChunk.length > 0 && currentHeight + rowHeight > budgetForCurrentPage) {
      chunks.push({
        rows: currentChunk,
        startSlIndex: currentSl,
        isLastPage: false,
      });
      currentSl += currentChunk.length;
      currentChunk = [];
      currentHeight = 0;
    }

    currentChunk.push(row);
    currentHeight += rowHeight;
  }

  if (currentChunk.length > 0) {
    // If the last page exceeds the LAST_PAGE_BUDGET (which includes totals & signatures),
    // cleanly split the overflow items to the next page
    if (!isChallan && currentHeight > LAST_PAGE_BUDGET && currentChunk.length > 1) {
      let splitIdx = currentChunk.length - 1;
      let remHeight = currentHeight;
      while (splitIdx > 0 && remHeight > LAST_PAGE_BUDGET) {
        const lines = calculateItemVisualLines(currentChunk[splitIdx].desc, isChallan);
        remHeight -= getItemRowHeight(lines);
        splitIdx--;
      }

      const page1Rows = currentChunk.slice(0, splitIdx + 1);
      const page2Rows = currentChunk.slice(splitIdx + 1);

      if (page1Rows.length > 0) {
        chunks.push({
          rows: page1Rows,
          startSlIndex: currentSl,
          isLastPage: false,
        });
        currentSl += page1Rows.length;
      }
      if (page2Rows.length > 0) {
        chunks.push({
          rows: page2Rows,
          startSlIndex: currentSl,
          isLastPage: true,
        });
      }
    } else {
      chunks.push({
        rows: currentChunk,
        startSlIndex: currentSl,
        isLastPage: true,
      });
    }
  }

  chunks.forEach((chunk, idx) => {
    chunk.isLastPage = idx === chunks.length - 1;
  });

  return chunks;
};

/**
 * Builds a compact, space-maximized worksheet matching the exact visual document.
 * Items use space as per text only, no overdisplaying of empty rows, centered stamp, and clean unclipped documents.
 */
const buildDocumentWorksheet = (
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
  docType: "quotation" | "challan" | "invoice",
  messers: string,
  address: string,
  challanNo: string,
  dateVal: string,
  requisitionNo: string,
  pageRows: QuotationRow[],
  mergedRegions: MergedRegion[],
  invoiceNo?: string,
  poNumber?: string,
  vatPercent: number = 0,
  transportationFee: number = 0,
  logoId: number | null = null,
  stampId: number | null = null,
  startSlIndex: number = 1,
  pageIndex: number = 1,
  totalPages: number = 1,
  isLastPage: boolean = true,
  allDocumentRows: QuotationRow[] = [],
  cellFormats?: CellFormatMap
) => {
  // Page Setup: Fit to 1 Page Wide and 1 Page Tall on standard A4 portrait
  worksheet.pageSetup = {
    paperSize: 9, // A4
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.25,
      bottom: 0.25,
      header: 0.1,
      footer: 0.1,
    },
    horizontalCentered: true,
    showGridLines: false, // Print with clean sharp borders
  };

  (worksheet.properties as any).pageSetUpPr = { fitToPage: true };

  const isChallan = docType === "challan";
  const isInvoice = docType === "invoice";
  const totalCols = isChallan ? 4 : 6;
  const lastColLetter = isChallan ? "D" : "F";

  // Set tight, well-proportioned column widths
  if (isChallan) {
    worksheet.columns = [
      { key: "A", width: 7.0 },  // SL / Logo
      { key: "B", width: 56.0 }, // Description / Company info
      { key: "C", width: 16.0 }, // Qty / Right Box Label
      { key: "D", width: 22.0 }, // Unit / Right Box Value
    ];
  } else {
    worksheet.columns = [
      { key: "A", width: 6.5 },  // SL / Logo
      { key: "B", width: 49.0 }, // Description / Company info
      { key: "C", width: 9.0 },  // Qty / Right Box Label
      { key: "D", width: 10.5 }, // Unit / Right Box Label
      { key: "E", width: 12.0 }, // Price / Right Box Value
      { key: "F", width: 14.5 }, // Amount / Right Box Value
    ];
  }

  // Standard Excel grid lines enabled for easy editing
  worksheet.views = [{ showGridLines: true }];

  // 1. Format Title in Row 1 (Center of page)
  worksheet.getRow(1).height = 22;
  worksheet.mergeCells(`A1:${lastColLetter}1`);
  const titleCell = worksheet.getCell("A1");
  const baseTitle = isChallan
    ? "DELIVERY CHALLAN"
    : isInvoice
    ? "INVOICE"
    : "QUOTATION";

  titleCell.value = baseTitle;
  titleCell.font = { name: "Arial", size: 12, bold: true, color: { argb: "000000" } };
  titleCell.alignment = { vertical: "middle", horizontal: "center" };

  // 2. 10 Blank Rows Gap (Rows 2 to 11) after format name
  for (let r = 2; r <= 11; r++) {
    worksheet.getRow(r).height = 14;
  }

  // 3. Metadata Boxes (Rows 12 to 16 - Compact, tightly placed)
  // Left Box: Messers & Address
  worksheet.mergeCells("A12:B12");
  worksheet.getCell("A12").value = "MESSERS:";
  worksheet.getCell("A12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };

  worksheet.mergeCells("A13:B13");
  const messersCell = worksheet.getCell("A13");
  const messersParsed = parseHtmlToExcelRuns(messers || "", {
    name: "Arial",
    size: 8.5,
    bold: true,
    color: { argb: "FF000000" },
  });
  if (messersParsed.hasFormatting && messersParsed.richText.length > 0) {
    messersCell.value = { richText: messersParsed.richText };
  } else {
    messersCell.value = messersParsed.plainText;
    messersCell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "FF000000" } };
  }
  if (messersParsed.highlightColor) {
    messersCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: messersParsed.highlightColor },
    };
  }
  messersCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  worksheet.mergeCells("A14:B14");
  worksheet.getCell("A14").value = "ADDRESS:";
  worksheet.getCell("A14").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };

  worksheet.mergeCells("A15:B16");
  const addrCell = worksheet.getCell("A15");
  const addrParsed = parseHtmlToExcelRuns(address || "", {
    name: "Arial",
    size: 8,
    color: { argb: "FF000000" },
  });
  if (addrParsed.hasFormatting && addrParsed.richText.length > 0) {
    addrCell.value = { richText: addrParsed.richText };
  } else {
    addrCell.value = addrParsed.plainText;
    addrCell.font = { name: "Arial", size: 8, color: { argb: "FF000000" } };
  }
  if (addrParsed.highlightColor) {
    addrCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: addrParsed.highlightColor },
    };
  }
  addrCell.alignment = { vertical: "top", wrapText: true };
  addrCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  // Left Box outer borders (Rows 12 to 16)
  for (let r = 12; r <= 16; r++) {
    worksheet.getRow(r).height = 13.5;
    for (let c = 1; c <= 2; c++) {
      const cell = worksheet.getCell(r, c);
      const cellBorders: any = { ...cell.border };
      if (r === 12) cellBorders.top = { style: "thin", color: { argb: "000000" } };
      if (r === 16) cellBorders.bottom = { style: "thin", color: { argb: "000000" } };
      if (c === 1) cellBorders.left = { style: "thin", color: { argb: "000000" } };
      if (c === 2) cellBorders.right = { style: "thin", color: { argb: "000000" } };
      cell.border = cellBorders;
    }
  }

  // Right Box (Rows 12 to 16) - DATE is strictly placed at the TOP in every format
  if (isInvoice) {
    // Row 12 (TOP): DATE
    worksheet.mergeCells("C12:D12");
    worksheet.getCell("C12").value = "DATE:";
    worksheet.getCell("C12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E12:F12");
    const dVal = worksheet.getCell("E12");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 13: INVOICE NO
    worksheet.mergeCells("C13:D13");
    worksheet.getCell("C13").value = "INVOICE NO.:";
    worksheet.getCell("C13").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E13:F13");
    const iVal = worksheet.getCell("E13");
    iVal.value = invoiceNo || "";
    iVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    iVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 14: CHALLAN NO
    worksheet.mergeCells("C14:D14");
    worksheet.getCell("C14").value = "CHALLAN NO.:";
    worksheet.getCell("C14").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E14:F14");
    const cVal = worksheet.getCell("E14");
    cVal.value = challanNo || "";
    cVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    cVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 15: REQUISITION NO
    worksheet.mergeCells("C15:D15");
    worksheet.getCell("C15").value = "REQUISITION NO.:";
    worksheet.getCell("C15").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E15:F15");
    const rVal = worksheet.getCell("E15");
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 16: P.O. NUMBER
    worksheet.mergeCells("C16:D16");
    worksheet.getCell("C16").value = "P.O. NUMBER:";
    worksheet.getCell("C16").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E16:F16");
    const pVal = worksheet.getCell("E16");
    pVal.value = poNumber || "";
    pVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    pVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
  } else if (isChallan) {
    // Row 12 (TOP): DATE
    worksheet.getCell("C12").value = "DATE:";
    worksheet.getCell("C12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    const dVal = worksheet.getCell("D12");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 13: CHALLAN NO
    worksheet.getCell("C13").value = "CHALLAN NO.:";
    worksheet.getCell("C13").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    const cVal = worksheet.getCell("D13");
    cVal.value = challanNo || "";
    cVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    cVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 14: REQUISITION NO
    worksheet.getCell("C14").value = "REQUISITION NO.:";
    worksheet.getCell("C14").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    const rVal = worksheet.getCell("D14");
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Rows 15, 16: blank dotted lines
    for (let r = 15; r <= 16; r++) {
      worksheet.getCell(`D${r}`).border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    }
  } else {
    // Quotation - Row 12 (TOP): DATE
    worksheet.mergeCells("C12:D12");
    worksheet.getCell("C12").value = "DATE:";
    worksheet.getCell("C12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E12:F12");
    const dVal = worksheet.getCell("E12");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 13: REQUISITION NO
    worksheet.mergeCells("C13:D13");
    worksheet.getCell("C13").value = "REQUISITION NO.:";
    worksheet.getCell("C13").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E13:F13");
    const rVal = worksheet.getCell("E13");
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    for (let r = 14; r <= 16; r++) {
      worksheet.mergeCells(`E${r}:F${r}`);
      worksheet.getCell(`E${r}`).border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    }
  }

  // Right Box outer borders (Rows 12 to 16)
  const rStartCol = 3;
  for (let r = 12; r <= 16; r++) {
    for (let c = rStartCol; c <= totalCols; c++) {
      const cell = worksheet.getCell(r, c);
      const cellBorders: any = { ...cell.border };
      if (r === 12) cellBorders.top = { style: "thin", color: { argb: "000000" } };
      if (r === 16) cellBorders.bottom = { style: "thin", color: { argb: "000000" } };
      if (c === rStartCol) cellBorders.left = { style: "thin", color: { argb: "000000" } };
      if (c === totalCols) cellBorders.right = { style: "thin", color: { argb: "000000" } };
      cell.border = cellBorders;
    }
  }

  // 4. Main Table Header (Row 17)
  const headerRow = worksheet.getRow(17);
  headerRow.height = 18;

  const colHeaders = isChallan
    ? ["SL", "Description of Marine Items / Spare Parts", "Qty", "Unit"]
    : ["SL", "Description of Marine Items / Spare Parts", "Qty", "Unit", "Price", "Amount"];

  colHeaders.forEach((text, i) => {
    const colIdx = i + 1;
    const cell = headerRow.getCell(colIdx);
    cell.value = text;
    cell.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    cell.alignment = {
      vertical: "middle",
      horizontal: colIdx === 2 ? "left" : "center",
      wrapText: true,
    };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "F1F5F9" },
    };
    cell.border = {
      top: { style: "medium", color: { argb: "000000" } },
      bottom: { style: "medium", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };
  });

  // Helper to apply custom CellFormat, inline RichText, highlights, colors, adjustments & borders to ExcelJS cell
  const applyCellFormatToExcel = (
    cell: ExcelJS.Cell,
    rIndex: number,
    cIndex: number,
    defaultAlign: "left" | "center" | "right",
    rawValue?: string,
    options?: {
      isLastRow?: boolean;
      isNumeric?: boolean;
      numericVal?: number;
      formula?: string;
    }
  ) => {
    const key = `${rIndex}_${cIndex}`;
    const fmt = cellFormats ? cellFormats[key] : undefined;

    // 1. Build Base Font
    const baseFont: Partial<ExcelJS.Font> = {
      name: fmt?.fontFamily || cell.font?.name || "Arial",
      size: fmt?.fontSize !== undefined ? fmt.fontSize : (cell.font?.size || 8),
      bold: fmt?.bold !== undefined ? fmt.bold : cell.font?.bold,
      italic: fmt?.italic !== undefined ? fmt.italic : cell.font?.italic,
      underline: fmt?.underline ? (fmt.underline === "double" ? "double" : fmt.underline === "single" ? true : false) : cell.font?.underline,
      color: fmt?.color ? { argb: colorToArgb(fmt.color) || "FF000000" } : (cell.font?.color || { argb: "FF000000" }),
    };

    // 2. Parse Value & Inline RichText (word-level formatting)
    let inlineHighlight: string | undefined = undefined;
    if (options?.formula) {
      cell.value = { formula: options.formula } as any;
      cell.font = baseFont;
    } else if (rawValue && rawValue.includes("<")) {
      const parsed = parseHtmlToExcelRuns(rawValue, baseFont);
      if (parsed.hasFormatting && parsed.richText.length > 0) {
        cell.value = { richText: parsed.richText };
        inlineHighlight = parsed.highlightColor;
      } else {
        if (options?.isNumeric && options.numericVal !== undefined && !isNaN(options.numericVal)) {
          cell.value = options.numericVal;
        } else {
          cell.value = parsed.plainText;
        }
        cell.font = baseFont;
      }
    } else if (options?.isNumeric && options.numericVal !== undefined && !isNaN(options.numericVal)) {
      cell.value = options.numericVal;
      cell.font = baseFont;
    } else if (rawValue !== undefined) {
      cell.value = rawValue;
      cell.font = baseFont;
    } else {
      cell.font = baseFont;
    }

    // 3. Highlight / Background Fill
    const fillArgb = fmt?.bgColor && fmt.bgColor !== "transparent"
      ? colorToArgb(fmt.bgColor)
      : inlineHighlight;

    if (fillArgb) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: fillArgb },
      };
    } else if (fmt?.bgColor === "transparent") {
      delete (cell as any).fill;
    }

    // 4. Adjustments (Alignment, Indent, Orientation, WrapText)
    const alignObj: Partial<ExcelJS.Alignment> = {
      vertical: fmt?.valign === "top" ? "top" : fmt?.valign === "middle" ? "middle" : fmt?.valign === "bottom" ? "bottom" : (cell.alignment?.vertical || "middle"),
      horizontal: fmt?.align || defaultAlign,
      wrapText: true,
    };

    if (fmt?.indent) {
      alignObj.indent = fmt.indent;
    }

    if (fmt?.orientation) {
      if (fmt.orientation === "angle-up") alignObj.textRotation = 45;
      else if (fmt.orientation === "angle-down") alignObj.textRotation = -45;
      else if (fmt.orientation === "vertical") alignObj.textRotation = 255;
      else if (fmt.orientation === "rotate-up") alignObj.textRotation = 90;
      else if (fmt.orientation === "rotate-down") alignObj.textRotation = -90;
    }
    cell.alignment = alignObj;

    // 5. Borders (Respects custom toolbar borders, presets, and standard table grid)
    const isLastRow = !!options?.isLastRow;
    const defaultBottomStyle: ExcelJS.BorderStyle = isLastRow ? "medium" : "thin";

    const topBorder = parseBorderSide(fmt?.borders?.top, "thin");
    const bottomBorder = parseBorderSide(fmt?.borders?.bottom, defaultBottomStyle);
    const leftBorder = parseBorderSide(fmt?.borders?.left, "thin");
    const rightBorder = parseBorderSide(fmt?.borders?.right, "thin");

    cell.border = {
      top: topBorder,
      bottom: bottomBorder,
      left: leftBorder,
      right: rightBorder,
    };
  };

  // 5. Table rows filling - Uses space strictly as per text only
  let currentRowNum = 18;
  const numItemsOnThisPage = pageRows.length;
  // Display only the actual items on this page (or 1 minimal row if totally empty)
  const displayRowCount = Math.max(numItemsOnThisPage, 1);

  for (let idx = 0; idx < displayRowCount; idx++) {
    const r = worksheet.getRow(currentRowNum);
    const rowData = pageRows[idx];

    const slVal = startSlIndex + idx;
    const rawDesc = rowData ? rowData.desc : "";
    const rawQty = rowData ? rowData.qty : "";
    const rawUnit = rowData ? rowData.unit : "";
    const rawPrice = rowData ? rowData.price : "";

    const cleanQtyStr = rawQty ? htmlToPlainText(rawQty).trim() : "";
    const qtyVal = cleanQtyStr ? parseNumericInput(cleanQtyStr) : "";
    const isNumericQty = typeof qtyVal === "number" && !isNaN(qtyVal) && qtyVal !== 0;

    const cleanPriceStr = rawPrice ? htmlToPlainText(rawPrice).trim() : "";
    const priceVal = cleanPriceStr ? parseNumericInput(cleanPriceStr) : "";
    const isNumericPrice = typeof priceVal === "number" && !isNaN(priceVal) && priceVal !== 0;

    // Dynamic compact row height based strictly on actual visual lines
    const plainDesc = htmlToPlainText(rawDesc);
    const visualLines = calculateItemVisualLines(plainDesc, isChallan);
    let rowH = getItemRowHeight(visualLines);

    // If description or any cell has large custom font size, adjust row height accordingly
    const descFmt = cellFormats ? cellFormats[`${startSlIndex - 1 + idx}_0`] : undefined;
    if (descFmt?.fontSize && descFmt.fontSize > 11) {
      rowH = Math.max(rowH, visualLines * descFmt.fontSize * 1.35);
    }
    r.height = rowH;

    const isLastItemRow = idx === displayRowCount - 1;

    // Set SL (Col 1)
    const cellSL = r.getCell(1);
    applyCellFormatToExcel(cellSL, startSlIndex - 1 + idx, -1, "center", rowData ? String(slVal) : "", {
      isLastRow: isLastItemRow,
    });

    // Set Description (Col 2)
    const cellDesc = r.getCell(2);
    applyCellFormatToExcel(cellDesc, startSlIndex - 1 + idx, 0, "left", rawDesc, {
      isLastRow: isLastItemRow,
    });

    // Set Qty (Col 3)
    const cellQty = r.getCell(3);
    applyCellFormatToExcel(cellQty, startSlIndex - 1 + idx, 1, "center", rawQty, {
      isLastRow: isLastItemRow,
      isNumeric: isNumericQty,
      numericVal: typeof qtyVal === "number" ? qtyVal : undefined,
    });
    if (isNumericQty) {
      cellQty.numFmt = "#,##0.00";
    }

    // Set Unit (Col 4)
    const cellUnit = r.getCell(4);
    applyCellFormatToExcel(cellUnit, startSlIndex - 1 + idx, 2, "center", rawUnit, {
      isLastRow: isLastItemRow,
    });

    // Set Price & Amount for Quotation / Invoice (Cols 5 & 6)
    if (!isChallan) {
      const cellPrice = r.getCell(5);
      applyCellFormatToExcel(cellPrice, startSlIndex - 1 + idx, 3, "right", rawPrice, {
        isLastRow: isLastItemRow,
        isNumeric: isNumericPrice,
        numericVal: typeof priceVal === "number" ? priceVal : undefined,
      });
      if (isNumericPrice) {
        cellPrice.numFmt = "#,##0.00";
      }

      const cellAmount = r.getCell(6);
      const hasContent = rowData && (plainDesc.trim() || cleanQtyStr || cleanPriceStr);
      const amountFormula = hasContent
        ? `=IF(OR(C${currentRowNum}="", E${currentRowNum}=""), 0, C${currentRowNum}*E${currentRowNum})`
        : undefined;

      applyCellFormatToExcel(cellAmount, startSlIndex - 1 + idx, 4, "right", undefined, {
        isLastRow: isLastItemRow,
        formula: amountFormula,
      });
      cellAmount.numFmt = "#,##0.00";
    }

    currentRowNum++;
  }

  // Apply Cell Merging to Excel Worksheet for this page
  const pageStartRowIdx = startSlIndex - 1;
  const pageEndRowIdx = pageStartRowIdx + numItemsOnThisPage - 1;

  mergedRegions.forEach((region) => {
    if (region.startRow >= pageStartRowIdx && region.endRow <= pageEndRowIdx) {
      const excelStartRow = (region.startRow - pageStartRowIdx) + 18;
      const excelEndRow = (region.endRow - pageStartRowIdx) + 18;
      const excelStartCol = region.startCol + 2;
      const excelEndCol = region.endCol + 2;

      if (
        excelStartRow >= 18 &&
        excelStartCol >= 1 &&
        excelEndRow < 18 + displayRowCount &&
        excelEndCol <= totalCols
      ) {
        try {
          worksheet.mergeCells(excelStartRow, excelStartCol, excelEndRow, excelEndCol);
        } catch (err) {
          console.warn("Could not merge cells in Excel workbook:", region, err);
        }
      }
    }
  });

  // 6. Totals & Words Block (Rendered strictly on the LAST page)
  if (!isChallan && isLastPage) {
    const totalRow = currentRowNum;
    const numTotalRows = isInvoice ? 4 : 1;

    for (let rOffset = 0; rOffset < numTotalRows; rOffset++) {
      worksheet.getRow(totalRow + rOffset).height = 17;
    }

    if (isInvoice) {
      worksheet.mergeCells(`A${totalRow}:D${totalRow + 3}`);
    } else {
      worksheet.mergeCells(`A${totalRow}:D${totalRow}`);
    }

    const wordCell = worksheet.getCell(`A${totalRow}`);

    const effectiveRowsForTotal = allDocumentRows.length > 0 ? allDocumentRows : pageRows;
    const subtotalValue = effectiveRowsForTotal.reduce((sum, r) => sum + r.amount, 0);
    const calculatedVat = isInvoice ? (subtotalValue * (vatPercent || 0)) / 100 : 0;
    const finalGrandTotal = isInvoice ? subtotalValue + calculatedVat + (transportationFee || 0) : subtotalValue;

    const words = numberToWords(Math.round(finalGrandTotal));
    const wordsStr = words ? words.toUpperCase() : "ZERO TAKA ONLY";
    wordCell.value = `AMOUNT IN WORDS: ${wordsStr}`;

    wordCell.font = { name: "Arial", size: 7.5, bold: true, italic: true, color: { argb: "000000" } };
    wordCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };

    for (let rOffset = 0; rOffset < numTotalRows; rOffset++) {
      const rNum = totalRow + rOffset;
      for (let c = 1; c <= 4; c++) {
        const cell = worksheet.getCell(rNum, c);
        cell.border = {
          top: rOffset === 0 ? { style: "medium", color: { argb: "000000" } } : undefined,
          bottom: rOffset === numTotalRows - 1 ? { style: "medium", color: { argb: "000000" } } : undefined,
          left: c === 1 ? { style: "medium", color: { argb: "000000" } } : undefined,
          right: c === 4 ? { style: "medium", color: { argb: "000000" } } : undefined,
        };
      }
    }

    const sumRange = `F18:F${totalRow - 1}`;

    if (isInvoice) {
      // Row 1: SUBTOTAL
      worksheet.getCell(`E${totalRow}`).value = totalPages > 1 ? "GRAND SUBTOTAL" : "SUBTOTAL";
      worksheet.getCell(`E${totalRow}`).font = { name: "Arial", size: 8, bold: true };
      worksheet.getCell(`E${totalRow}`).alignment = { vertical: "middle", horizontal: "right" };

      const subtotalValCell = worksheet.getCell(`F${totalRow}`);
      if (totalPages === 1) {
        subtotalValCell.value = { formula: `=SUM(${sumRange})` } as any;
      } else {
        subtotalValCell.value = subtotalValue;
      }
      subtotalValCell.font = { name: "Arial", size: 8, bold: true };
      subtotalValCell.alignment = { vertical: "middle", horizontal: "right" };
      subtotalValCell.numFmt = "#,##0.00";

      // Row 2: VAT
      worksheet.getCell(`E${totalRow + 1}`).value = `VAT (${vatPercent || 0}%)`;
      worksheet.getCell(`E${totalRow + 1}`).font = { name: "Arial", size: 8, bold: true };
      worksheet.getCell(`E${totalRow + 1}`).alignment = { vertical: "middle", horizontal: "right" };

      const vatValCell = worksheet.getCell(`F${totalRow + 1}`);
      vatValCell.value = calculatedVat;
      vatValCell.font = { name: "Arial", size: 8, bold: true };
      vatValCell.alignment = { vertical: "middle", horizontal: "right" };
      vatValCell.numFmt = "#,##0.00";

      // Row 3: TRANSPORTATION
      worksheet.getCell(`E${totalRow + 2}`).value = "TRANSPORTATION FEE";
      worksheet.getCell(`E${totalRow + 2}`).font = { name: "Arial", size: 8, bold: true };
      worksheet.getCell(`E${totalRow + 2}`).alignment = { vertical: "middle", horizontal: "right" };

      const transValCell = worksheet.getCell(`F${totalRow + 2}`);
      transValCell.value = transportationFee || 0;
      transValCell.font = { name: "Arial", size: 8, bold: true };
      transValCell.alignment = { vertical: "middle", horizontal: "right" };
      transValCell.numFmt = "#,##0.00";

      // Row 4: GRAND TOTAL
      worksheet.getCell(`E${totalRow + 3}`).value = "GRAND TOTAL";
      worksheet.getCell(`E${totalRow + 3}`).font = { name: "Arial", size: 8.5, bold: true };
      worksheet.getCell(`E${totalRow + 3}`).alignment = { vertical: "middle", horizontal: "right" };

      const grandValCell = worksheet.getCell(`F${totalRow + 3}`);
      grandValCell.value = finalGrandTotal;
      grandValCell.font = { name: "Arial", size: 8.5, bold: true };
      grandValCell.alignment = { vertical: "middle", horizontal: "right" };
      grandValCell.numFmt = "#,##0.00";

      for (let rOffset = 0; rOffset < numTotalRows; rOffset++) {
        const rNum = totalRow + rOffset;
        for (let c = 5; c <= 6; c++) {
          const cell = worksheet.getCell(rNum, c);
          cell.border = {
            top: rOffset === 0 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
            bottom: rOffset === numTotalRows - 1 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
            left: c === 5 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
            right: c === 6 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
          };
        }
      }
    } else {
      // Quotation TOTAL
      worksheet.getCell(`E${totalRow}`).value = totalPages > 1 ? "GRAND TOTAL" : "TOTAL";
      worksheet.getCell(`E${totalRow}`).font = { name: "Arial", size: 8, bold: true };
      worksheet.getCell(`E${totalRow}`).alignment = { vertical: "middle", horizontal: "right" };

      const valCell = worksheet.getCell(`F${totalRow}`);
      if (totalPages === 1) {
        valCell.value = { formula: `=SUM(${sumRange})` } as any;
      } else {
        valCell.value = subtotalValue;
      }
      valCell.font = { name: "Arial", size: 8, bold: true };
      valCell.alignment = { vertical: "middle", horizontal: "right" };
      valCell.numFmt = "#,##0.00";

      for (let c = 5; c <= 6; c++) {
        const cell = worksheet.getCell(totalRow, c);
        cell.border = {
          top: { style: "medium", color: { argb: "000000" } },
          bottom: { style: "medium", color: { argb: "000000" } },
          left: c === 5 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
          right: c === 6 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "000000" } },
        };
      }
    }

    currentRowNum = totalRow + numTotalRows;
  }

  // 6. Signatures Area - Spacious room & Enlarged Stamp
  worksheet.getRow(currentRowNum).height = 10;
  currentRowNum++;

  // "For Comilla Traders" row
  const authHeaderRow = currentRowNum;
  if (!isChallan) {
    worksheet.getRow(authHeaderRow).height = 15;
    worksheet.mergeCells(`E${authHeaderRow}:F${authHeaderRow}`);
    const authTitle = worksheet.getCell(`E${authHeaderRow}`);
    authTitle.value = "For Comilla Traders";
    authTitle.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
    authTitle.alignment = { vertical: "middle", horizontal: "center" };
    currentRowNum++;
  }

  // Generous room for signatures and enlarged stamp
  worksheet.getRow(currentRowNum).height = 42;
  const sigRow = currentRowNum + 1;
  worksheet.getRow(sigRow).height = 18;

  // Receiver's Signature
  worksheet.mergeCells(`A${sigRow}:B${sigRow}`);
  const recSig = worksheet.getCell(`A${sigRow}`);
  recSig.value = "Receiver's Signature";
  recSig.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
  recSig.alignment = { vertical: "middle", horizontal: "center" };
  recSig.border = { top: { style: "thin", color: { argb: "000000" } } };

  // Authorized Signature
  const authColStart = isChallan ? "C" : "E";
  const authColEnd = isChallan ? "D" : "F";
  worksheet.mergeCells(`${authColStart}${sigRow}:${authColEnd}${sigRow}`);
  const authSig = worksheet.getCell(`${authColStart}${sigRow}`);
  authSig.value = "Authorized Signature";
  authSig.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
  authSig.alignment = { vertical: "middle", horizontal: "center" };
  authSig.border = { top: { style: "thin", color: { argb: "000000" } } };

  // Centered Official Stamp (Enlarged and positioned over Authorized Signature)
  if (stampId !== null) {
    const stampCol = isChallan ? 2.65 : 4.64;
    worksheet.addImage(stampId, {
      tl: { col: stampCol, row: sigRow - 2.25 },
      ext: { width: 88, height: 88 },
    });
  }

  // 10. Footer Disclaimer Notice
  const noticeRow = sigRow + 2;
  worksheet.getRow(sigRow + 1).height = 4;
  worksheet.getRow(noticeRow).height = 13;
  worksheet.mergeCells(`A${noticeRow}:${lastColLetter}${noticeRow}`);
  const noticeCell = worksheet.getCell(`A${noticeRow}`);
  noticeCell.value = "ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE.";
  noticeCell.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "000000" } };
  noticeCell.alignment = { vertical: "middle", horizontal: "center" };
};

/**
 * Generates complete multi-sheet Excel Workbook for the currently selected document type.
 * Maximizes A4 space utilization, uses space per text only, compact layout, centered stamp, and automatic page overflow transfer.
 */
export const generateExcelWorkbook = async (
  docType: "quotation" | "challan" | "invoice",
  messers: string,
  address: string,
  challanNo: string,
  dateVal: string,
  requisitionNo: string,
  rows: QuotationRow[],
  mergedRegions: MergedRegion[],
  invoiceNo?: string,
  poNumber?: string,
  vatPercent?: number,
  transportationFee?: number,
  cellFormats?: CellFormatMap
): Promise<ExcelJS.Workbook> => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Comilla Traders";
  workbook.lastModifiedBy = "Comilla Traders";
  workbook.created = new Date();
  workbook.modified = new Date();

  // Fetch images in parallel
  const [logoData, stampData] = await Promise.all([
    getBase64Image("https://i.ibb.co.com/gFBkpt8B/Chat-GPT-Image-Apr-23-2026-01-10-13-PM.png"),
    getBase64Image("https://i.ibb.co.com/jZswrtn6/image-4-removebg-preview.png"),
  ]);

  let logoId: number | null = null;
  if (logoData) {
    logoId = workbook.addImage({
      base64: logoData.base64,
      extension: (logoData.ext as any) || "png",
    });
  }

  let stampId: number | null = null;
  if (stampData) {
    stampId = workbook.addImage({
      base64: stampData.base64,
      extension: (stampData.ext as any) || "png",
    });
  }

  let effectiveRows = rows;
  let lastNonEmptyIndex = -1;
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (r.desc.trim() !== "" || r.qty.trim() !== "" || r.price.trim() !== "") {
      lastNonEmptyIndex = i;
      break;
    }
  }
  if (lastNonEmptyIndex >= 0) {
    effectiveRows = rows.slice(0, lastNonEmptyIndex + 1);
  }

  // Dynamic capacity-aware pagination maximizing A4 page utilization
  const pageChunks = paginateRowsForExcel(effectiveRows, docType);
  const totalPages = pageChunks.length;

  const baseSheetNames: Record<string, string> = {
    quotation: "Quotation",
    challan: "Challan",
    invoice: "Invoice",
  };

  pageChunks.forEach((chunk, pageIdx) => {
    const pageNumber = pageIdx + 1;
    const sheetName =
      totalPages === 1
        ? baseSheetNames[docType]
        : `${baseSheetNames[docType]} - Pg ${pageNumber}`;

    const ws = workbook.addWorksheet(sheetName);
    buildDocumentWorksheet(
      workbook,
      ws,
      docType,
      messers,
      address,
      challanNo,
      dateVal,
      requisitionNo,
      chunk.rows,
      mergedRegions,
      invoiceNo,
      poNumber,
      vatPercent || 0,
      transportationFee || 0,
      logoId,
      stampId,
      chunk.startSlIndex,
      pageNumber,
      totalPages,
      chunk.isLastPage,
      effectiveRows,
      cellFormats
    );
  });

  return workbook;
};
