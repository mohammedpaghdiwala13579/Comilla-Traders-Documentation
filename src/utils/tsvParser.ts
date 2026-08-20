/**
 * Robust Clipboard & Excel / Google Sheets parser
 * Supports:
 * - HTML table extraction from Excel/Google Sheets rich clipboard
 * - TSV (Tab Separated Values) with multi-line quote escapes
 * - CSV (Comma / Semicolon separated values)
 * - Intelligent header detection & column alignment
 */

export interface ParsedClipboardResult {
  grid: string[][];
  hasHeader: boolean;
  hasSerialColumn: boolean;
  columnMapping?: ("sl" | "desc" | "qty" | "unit" | "price" | "amount" | "ignore")[];
  detectedFormat: "html_table" | "tsv" | "csv" | "plain_lines";
}

/**
 * Extracts a 2D string array from an HTML table clipboard payload
 */
export function parseHTMLTable(html: string): string[][] | null {
  if (!html || !html.includes("<table") || typeof DOMParser === "undefined") {
    return null;
  }

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");
    const table = doc.querySelector("table");
    if (!table) return null;

    const rows = Array.from(table.querySelectorAll("tr"));
    if (rows.length === 0) return null;

    const grid: string[][] = [];

    rows.forEach((tr) => {
      const cells = Array.from(tr.querySelectorAll("th, td"));
      if (cells.length === 0) return;

      const rowValues = cells.map((cell) => {
        // Clean line breaks inside cell: replace <br> with newline
        const clones = cell.cloneNode(true) as HTMLElement;
        const brs = clones.querySelectorAll("br");
        brs.forEach((br) => br.replaceWith("\n"));

        let text = clones.textContent || "";
        // Clean up non-breaking spaces
        text = text.replace(/\u00A0/g, " ").trim();
        return text;
      });

      // Avoid adding completely empty rows
      const hasContent = rowValues.some((val) => val.length > 0);
      if (hasContent) {
        grid.push(rowValues);
      }
    });

    return grid.length > 0 ? grid : null;
  } catch (e) {
    console.warn("Failed to parse HTML table from clipboard:", e);
    return null;
  }
}

/**
 * State-machine parser for TSV (Tab-Separated Values)
 */
export function parseTSV(text: string): string[][] {
  const result: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          cell += '"';
          i++; // Skip next quote
        } else {
          inQuotes = false;
        }
      } else {
        cell += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === "\t") {
        row.push(cell.trim());
        cell = "";
      } else if (char === "\r") {
        if (nextChar === "\n") {
          row.push(cell.trim());
          result.push(row);
          row = [];
          cell = "";
          i++; // Skip \n
        } else {
          row.push(cell.trim());
          result.push(row);
          row = [];
          cell = "";
        }
      } else if (char === "\n") {
        row.push(cell.trim());
        result.push(row);
        row = [];
        cell = "";
      } else {
        cell += char;
      }
    }
  }

  if (cell !== "" || row.length > 0) {
    row.push(cell.trim());
    result.push(row);
  }

  // Remove empty trailing row
  while (
    result.length > 0 &&
    result[result.length - 1].every((c) => c === "")
  ) {
    result.pop();
  }

  return result;
}

/**
 * Parses CSV text with delimiter (comma or semicolon)
 */
export function parseCSV(text: string, delimiter: "," | ";" = ","): string[][] {
  const result: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          cell += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cell += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === delimiter) {
        row.push(cell.trim());
        cell = "";
      } else if (char === "\r") {
        if (nextChar === "\n") {
          row.push(cell.trim());
          result.push(row);
          row = [];
          cell = "";
          i++;
        } else {
          row.push(cell.trim());
          result.push(row);
          row = [];
          cell = "";
        }
      } else if (char === "\n") {
        row.push(cell.trim());
        result.push(row);
        row = [];
        cell = "";
      } else {
        cell += char;
      }
    }
  }

  if (cell !== "" || row.length > 0) {
    row.push(cell.trim());
    result.push(row);
  }

  while (
    result.length > 0 &&
    result[result.length - 1].every((c) => c === "")
  ) {
    result.pop();
  }

  return result;
}

/**
 * Universal clipboard parser that extracts 2D grid
 */
export function parseClipboardData(input: {
  text: string;
  html?: string;
}): ParsedClipboardResult {
  const { text, html } = input;

  // 1. Try HTML Table parsing first (from Excel / Google Sheets)
  if (html && html.includes("<table")) {
    const htmlGrid = parseHTMLTable(html);
    if (htmlGrid && htmlGrid.length > 0 && htmlGrid.some((r) => r.length > 1 || htmlGrid.length > 1)) {
      return analyzeGrid(htmlGrid, "html_table");
    }
  }

  const rawText = (text || "").trim();
  if (!rawText) {
    return {
      grid: [],
      hasHeader: false,
      hasSerialColumn: false,
      detectedFormat: "plain_lines",
    };
  }

  // 2. Try TSV (Tab separated)
  if (rawText.includes("\t")) {
    const tsvGrid = parseTSV(rawText);
    if (tsvGrid.length > 0) {
      return analyzeGrid(tsvGrid, "tsv");
    }
  }

  // 3. Try CSV (Semicolon or comma separated if multi-column on lines)
  const lines = rawText.split(/\r?\n/).filter((l) => l.trim().length > 0);
  if (lines.length > 0) {
    const commaCount = (lines[0].match(/,/g) || []).length;
    const semiCount = (lines[0].match(/;/g) || []).length;

    if (commaCount >= 1 && (commaCount >= semiCount)) {
      const csvGrid = parseCSV(rawText, ",");
      if (csvGrid.length > 0 && csvGrid[0].length > 1) {
        return analyzeGrid(csvGrid, "csv");
      }
    } else if (semiCount >= 1) {
      const semiGrid = parseCSV(rawText, ";");
      if (semiGrid.length > 0 && semiGrid[0].length > 1) {
        return analyzeGrid(semiGrid, "csv");
      }
    }
  }

  // 4. Fallback: single column with multiple lines
  const lineGrid = lines.map((l) => [l.trim()]);
  return analyzeGrid(lineGrid, "plain_lines");
}

/**
 * Analyzes grid structure, detects headers and whether Column 0 is a Serial number
 */
function analyzeGrid(
  rawGrid: string[][],
  detectedFormat: "html_table" | "tsv" | "csv" | "plain_lines"
): ParsedClipboardResult {
  if (rawGrid.length === 0) {
    return {
      grid: [],
      hasHeader: false,
      hasSerialColumn: false,
      detectedFormat,
    };
  }

  // Check if first row is a header
  const firstRow = rawGrid[0].map((c) => c.toLowerCase().trim());
  const headerKeywords = [
    "sl", "s/n", "s.no", "no", "item", "item #", "item no",
    "description", "particulars", "items", "details", "desc", "material", "name",
    "qty", "quantity", "qnty",
    "unit", "uom", "pkg", "unit of measure",
    "price", "rate", "unit price", "unit rate",
    "amount", "total", "total amount", "subtotal"
  ];

  const headerMatches = firstRow.filter((cell) =>
    headerKeywords.some((kw) => cell === kw || cell.startsWith(kw + " ") || cell.includes(kw))
  );

  const hasHeader = headerMatches.length >= 2 || (firstRow.length <= 2 && headerMatches.length >= 1);

  // Check if column 0 consists mainly of numbers (1, 2, 3, 4...) -> Serial Number column
  const dataRows = hasHeader ? rawGrid.slice(1) : rawGrid;
  let numericSerialCount = 0;
  
  dataRows.forEach((r, idx) => {
    if (r.length > 0) {
      const val = r[0].replace(/[.\-\s]/g, "").trim();
      const num = parseInt(val, 10);
      if (!isNaN(num) && (num === idx + 1 || num === idx || num > 0)) {
        numericSerialCount++;
      }
    }
  });

  const hasSerialColumn = dataRows.length > 0 && numericSerialCount >= Math.min(3, dataRows.length);

  return {
    grid: rawGrid,
    hasHeader,
    hasSerialColumn,
    detectedFormat,
  };
}
