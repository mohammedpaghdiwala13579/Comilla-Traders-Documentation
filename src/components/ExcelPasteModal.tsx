import React, { useState, useEffect } from "react";
import { X, FileSpreadsheet, Check, ArrowRight, AlertCircle, Upload, Plus, Layers, HelpCircle } from "lucide-react";
import * as XLSX from "xlsx";
import { parseClipboardData, parseHTMLTable, parseTSV, parseCSV } from "../utils/tsvParser";
import { QuotationRow } from "../types";

interface ExcelPasteModalProps {
  isOpen: boolean;
  onClose: () => void;
  onImportRows: (newRows: QuotationRow[], mode: "replace" | "append" | "insert_at", insertIndex?: number) => void;
  selectedRowIndex: number;
  totalCurrentRows: number;
  docType: "quotation" | "challan" | "invoice";
}

type ColumnTarget = "ignore" | "sl" | "desc" | "qty" | "unit" | "price" | "amount";

export default function ExcelPasteModal({
  isOpen,
  onClose,
  onImportRows,
  selectedRowIndex,
  totalCurrentRows,
  docType,
}: ExcelPasteModalProps) {
  const [rawText, setRawText] = useState("");
  const [parsedGrid, setParsedGrid] = useState<string[][]>([]);
  const [columnMappings, setColumnMappings] = useState<ColumnTarget[]>([]);
  const [hasHeaderRow, setHasHeaderRow] = useState(false);
  const [importMode, setImportMode] = useState<"replace" | "append" | "insert_at">("append");
  const [extraBlankRows, setExtraBlankRows] = useState<number>(0);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);

  useEffect(() => {
    if (!isOpen) {
      setRawText("");
      setParsedGrid([]);
      setColumnMappings([]);
      setStatusMessage(null);
    }
  }, [isOpen]);

  if (!isOpen) return null;

  // Process raw text or HTML data
  const processRawInput = (text: string, html?: string) => {
    const result = parseClipboardData({ text, html });
    if (result.grid.length === 0) {
      setParsedGrid([]);
      setColumnMappings([]);
      return;
    }

    setParsedGrid(result.grid);
    setHasHeaderRow(result.hasHeader);

    // Auto-determine column mappings
    const detectedCols = autoMapColumns(result.grid, result.hasHeader, docType);
    setColumnMappings(detectedCols);
    setStatusMessage(`Successfully detected ${result.grid.length} row(s) and ${result.grid[0]?.length || 0} column(s).`);
  };

  // Heuristic auto-mapping for columns
  const autoMapColumns = (grid: string[][], hasHeader: boolean, currentDocType: string): ColumnTarget[] => {
    if (grid.length === 0) return [];
    const colCount = Math.max(...grid.map((r) => r.length));
    const mappings: ColumnTarget[] = Array(colCount).fill("ignore");

    const headerRow = hasHeader ? grid[0].map((c) => c.toLowerCase().trim()) : [];
    const sampleRows = (hasHeader ? grid.slice(1, 10) : grid.slice(0, 10));

    // First pass: try matching header names
    if (hasHeader && headerRow.length > 0) {
      headerRow.forEach((h, idx) => {
        if (idx >= colCount) return;
        if (/^(sl|s\/n|s\.no|no|#|item\s*#|item\s*no)$/.test(h)) {
          mappings[idx] = "sl";
        } else if (/^(desc|description|item|items|particulars|particular|details|name|specification)$/.test(h)) {
          mappings[idx] = "desc";
        } else if (/^(qty|quantity|qnty|count|pieces|pcs)$/.test(h)) {
          mappings[idx] = "qty";
        } else if (/^(unit|uom|unit\s*of\s*measure|pkg|type)$/.test(h)) {
          mappings[idx] = "unit";
        } else if (/^(price|rate|unit\s*price|unit\s*rate|cost|amount\/unit)$/.test(h)) {
          mappings[idx] = currentDocType === "challan" ? "ignore" : "price";
        } else if (/^(amount|total|total\s*price|subtotal|line\s*total)$/.test(h)) {
          mappings[idx] = "amount";
        }
      });
    }

    // Second pass: fill unmapped columns based on standard conventions & sample values
    const unmappedIndices = mappings.map((m, idx) => (m === "ignore" ? idx : -1)).filter((idx) => idx !== -1);
    
    // If standard 4-column copy without headers [Desc, Qty, Unit, Price]
    if (colCount === 4 && unmappedIndices.length === 4) {
      return currentDocType === "challan" 
        ? ["desc", "qty", "unit", "ignore"] 
        : ["desc", "qty", "unit", "price"];
    }

    // If 5-column copy [SL, Desc, Qty, Unit, Price]
    if (colCount === 5 && unmappedIndices.length === 5) {
      return currentDocType === "challan" 
        ? ["sl", "desc", "qty", "unit", "ignore"] 
        : ["sl", "desc", "qty", "unit", "price"];
    }

    // If 6-column copy [SL, Desc, Qty, Unit, Price, Amount]
    if (colCount >= 6 && unmappedIndices.length >= 6) {
      const res: ColumnTarget[] = ["sl", "desc", "qty", "unit", "price", "amount"];
      while (res.length < colCount) res.push("ignore");
      return res;
    }

    // Fallback inspect sample values
    unmappedIndices.forEach((colIdx) => {
      const colValues = sampleRows.map((r) => (r[colIdx] || "").trim()).filter(Boolean);
      if (colValues.length === 0) return;

      const isAllDigits = colValues.every((v) => /^\d+$/.test(v));
      const isNumeric = colValues.every((v) => /^-?\d+(\.\d+)?$/.test(v.replace(/,/g, "")));
      const hasUnits = colValues.some((v) => /^(pcs|pc|nos|kg|mtr|mtrs|set|sets|ltr|drum|can|pkt|box|roll|pair|coil|feet|ft)$/i.test(v));

      if (hasUnits && !mappings.includes("unit")) {
        mappings[colIdx] = "unit";
      } else if (isAllDigits && colIdx === 0 && !mappings.includes("sl")) {
        mappings[colIdx] = "sl";
      } else if (isNumeric && !mappings.includes("qty")) {
        mappings[colIdx] = "qty";
      } else if (isNumeric && !mappings.includes("price") && currentDocType !== "challan") {
        mappings[colIdx] = "price";
      } else if (!mappings.includes("desc")) {
        mappings[colIdx] = "desc";
      }
    });

    // Ensure at least description is mapped
    if (!mappings.includes("desc") && colCount > 0) {
      mappings[0] = "desc";
    }

    return mappings;
  };

  const handlePasteEvent = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    e.preventDefault();
    const text = e.clipboardData.getData("text/plain") || e.clipboardData.getData("text");
    const html = e.clipboardData.getData("text/html");
    setRawText(text);
    processRawInput(text, html);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to 2D array
        const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        const stringGrid: string[][] = jsonData.map((row) =>
          row.map((cell) => (cell !== null && cell !== undefined ? String(cell).trim() : ""))
        ).filter(r => r.some(c => c.length > 0));

        if (stringGrid.length > 0) {
          setParsedGrid(stringGrid);
          const tsv = stringGrid.map((r) => r.join("\t")).join("\n");
          setRawText(tsv);
          const result = parseClipboardData({ text: tsv });
          setHasHeaderRow(result.hasHeader);
          setColumnMappings(autoMapColumns(stringGrid, result.hasHeader, docType));
          setStatusMessage(`Loaded file "${file.name}" with ${stringGrid.length} row(s).`);
        }
      } catch (err) {
        console.error("Failed to parse Excel file:", err);
        setStatusMessage("Could not read file. Please ensure it is a valid .xlsx, .xls, or .csv file.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleReadClipboard = async () => {
    try {
      if (navigator.clipboard && navigator.clipboard.readText) {
        const text = await navigator.clipboard.readText();
        if (text) {
          setRawText(text);
          processRawInput(text);
        } else {
          setStatusMessage("Clipboard is currently empty.");
        }
      } else {
        setStatusMessage("Please press Ctrl+V directly into the text box below.");
      }
    } catch (err) {
      setStatusMessage("Clipboard access denied. Please click the text area and press Ctrl+V.");
    }
  };

  const handleApplyImport = () => {
    if (parsedGrid.length === 0) return;

    const dataRows = hasHeaderRow ? parsedGrid.slice(1) : parsedGrid;
    if (dataRows.length === 0) return;

    const descColIdx = columnMappings.indexOf("desc");
    const qtyColIdx = columnMappings.indexOf("qty");
    const unitColIdx = columnMappings.indexOf("unit");
    const priceColIdx = columnMappings.indexOf("price");

    const convertedRows: QuotationRow[] = dataRows.map((row, idx) => {
      const desc = descColIdx >= 0 ? (row[descColIdx] || "") : "";
      const qty = qtyColIdx >= 0 ? (row[qtyColIdx] || "") : "";
      const unit = unitColIdx >= 0 ? (row[unitColIdx] || "") : "";
      const price = priceColIdx >= 0 ? (row[priceColIdx] || "") : "";

      const q = parseFloat(qty) || 0;
      const p = parseFloat(price.replace(/,/g, "")) || 0;
      const amount = docType === "challan" ? 0 : q * p;

      return {
        sl: idx + 1,
        desc,
        qty,
        unit,
        price,
        amount,
      };
    });

    // Add extra blank rows if requested
    for (let i = 0; i < extraBlankRows; i++) {
      convertedRows.push({
        sl: convertedRows.length + 1,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      });
    }

    onImportRows(convertedRows, importMode, selectedRowIndex);
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-3 sm:p-5 z-[99999] animate-in fade-in duration-200">
      <div className="bg-white rounded-xl shadow-2xl border border-slate-200 w-full max-w-4xl max-h-[92vh] flex flex-col overflow-hidden animate-in zoom-in-95 duration-200">
        
        {/* Modal Header */}
        <div className="bg-gradient-to-r from-indigo-700 via-indigo-800 to-slate-900 text-white px-5 py-3.5 flex items-center justify-between shadow-xs">
          <div className="flex items-center gap-2.5">
            <div className="bg-white/10 p-2 rounded-lg text-indigo-200">
              <FileSpreadsheet className="h-5 w-5" />
            </div>
            <div>
              <h2 className="text-sm sm:text-base font-black tracking-wide flex items-center gap-2">
                EXCEL &amp; CLIPBOARD SMART IMPORTER
                <span className="bg-emerald-500 text-white text-[9px] font-bold px-1.5 py-0.5 rounded uppercase tracking-wider">
                  Universal Paste
                </span>
              </h2>
              <p className="text-[11px] text-indigo-100/90 font-medium">
                Copy columns from Microsoft Excel, Google Sheets, or CSV and paste them directly into your sheet.
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-indigo-200 hover:text-white p-1 hover:bg-white/10 rounded-lg transition-colors cursor-pointer"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-4 sm:p-5 overflow-y-auto flex-1 space-y-4 text-slate-800">
          
          {/* Top Quick Actions Bar */}
          <div className="flex flex-wrap items-center justify-between gap-2.5 bg-slate-50 p-3 rounded-lg border border-slate-200">
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={handleReadClipboard}
                className="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold px-3 py-1.5 rounded-lg shadow-sm flex items-center gap-1.5 transition-all cursor-pointer"
              >
                <FileSpreadsheet className="h-3.5 w-3.5" />
                <span>Paste from Clipboard</span>
              </button>
              
              <label className="bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 text-xs font-bold px-3 py-1.5 rounded-lg shadow-xs flex items-center gap-1.5 transition-all cursor-pointer">
                <Upload className="h-3.5 w-3.5 text-slate-500" />
                <span>Upload .xlsx / .csv</span>
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv, .tsv"
                  onChange={handleFileUpload}
                  className="sr-only"
                />
              </label>
            </div>

            {statusMessage && (
              <span className="text-xs text-indigo-800 font-semibold bg-indigo-50 px-2.5 py-1 rounded border border-indigo-200">
                {statusMessage}
              </span>
            )}
          </div>

          {/* Paste Input Box */}
          <div>
            <label className="block text-xs font-extrabold uppercase tracking-wider text-slate-700 mb-1 flex items-center justify-between">
              <span>Paste Excel Data Here (Ctrl + V):</span>
              <span className="text-[11px] font-normal text-slate-500 lowercase">
                paste raw cells or type comma/tab separated items
              </span>
            </label>
            <textarea
              rows={4}
              value={rawText}
              onChange={(e) => {
                setRawText(e.target.value);
                processRawInput(e.target.value);
              }}
              onPaste={handlePasteEvent}
              placeholder="Click here and press Ctrl+V (or Cmd+V on Mac) to paste copied Excel table or columns..."
              className="w-full font-mono text-xs p-3 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none bg-slate-50/50"
            />
          </div>

          {/* Preview & Column Mapping Section */}
          {parsedGrid.length > 0 && (
            <div className="space-y-3 pt-1">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-3">
                  <h3 className="text-xs font-extrabold uppercase tracking-wider text-slate-900 flex items-center gap-1.5">
                    <span>Parsed Preview &amp; Column Mapping</span>
                    <span className="bg-indigo-100 text-indigo-800 text-[10px] px-2 py-0.5 rounded-full font-bold">
                      {hasHeaderRow ? parsedGrid.length - 1 : parsedGrid.length} Rows
                    </span>
                  </h3>
                </div>

                <label className="flex items-center gap-2 text-xs font-semibold text-slate-700 select-none cursor-pointer bg-slate-50 px-2.5 py-1 rounded border border-slate-200">
                  <input
                    type="checkbox"
                    checked={hasHeaderRow}
                    onChange={(e) => setHasHeaderRow(e.target.checked)}
                    className="rounded text-indigo-600 focus:ring-indigo-500 h-3.5 w-3.5"
                  />
                  <span>First row contains table headers (skip when inserting)</span>
                </label>
              </div>

              {/* Grid Preview with Column Dropdowns */}
              <div className="border border-slate-300 rounded-lg overflow-x-auto max-h-56 shadow-inner bg-slate-50">
                <table className="w-full text-left text-xs border-collapse font-mono">
                  <thead className="bg-slate-100 border-b border-slate-300 sticky top-0 z-10">
                    <tr>
                      <th className="p-2 border-r border-slate-300 w-10 text-center text-slate-500 font-bold">
                        #
                      </th>
                      {Array.from({ length: Math.max(...parsedGrid.map((r) => r.length)) }).map((_, cIdx) => (
                        <th key={cIdx} className="p-1.5 border-r border-slate-300 min-w-[140px]">
                          <div className="space-y-1">
                            <span className="text-[10px] font-bold text-slate-500 uppercase block">
                              Column {cIdx + 1}
                            </span>
                            <select
                              value={columnMappings[cIdx] || "ignore"}
                              onChange={(e) => {
                                const newMappings = [...columnMappings];
                                newMappings[cIdx] = e.target.value as ColumnTarget;
                                setColumnMappings(newMappings);
                              }}
                              className="w-full text-[11px] font-sans font-bold bg-white border border-slate-300 rounded px-1.5 py-1 text-indigo-900 focus:ring-1 focus:ring-indigo-500 outline-none"
                            >
                              <option value="ignore">Skip / Ignore</option>
                              <option value="sl">Serial No (SL)</option>
                              <option value="desc">Description of Item</option>
                              <option value="qty">Quantity (Qty)</option>
                              <option value="unit">Unit (PCS, KG, etc.)</option>
                              {docType !== "challan" && <option value="price">Unit Price</option>}
                              {docType !== "challan" && <option value="amount">Amount (Calculated)</option>}
                            </select>
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 bg-white">
                    {parsedGrid.slice(0, 15).map((row, rIdx) => {
                      const isHeader = hasHeaderRow && rIdx === 0;
                      return (
                        <tr
                          key={rIdx}
                          className={`${
                            isHeader
                              ? "bg-amber-50/70 font-bold text-amber-900 italic"
                              : "hover:bg-indigo-50/30 text-slate-800"
                          }`}
                        >
                          <td className="p-1.5 border-r border-slate-200 text-center text-[10px] text-slate-400 font-bold">
                            {isHeader ? "HDR" : rIdx + 1}
                          </td>
                          {row.map((cell, cIdx) => (
                            <td
                              key={cIdx}
                              className="p-1.5 border-r border-slate-200 truncate max-w-[200px] text-[11px]"
                              title={cell}
                            >
                              {cell || <span className="text-slate-300 italic">(empty)</span>}
                            </td>
                          ))}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {parsedGrid.length > 15 && (
                <p className="text-[11px] text-slate-500 text-right italic">
                  Showing first 15 of {parsedGrid.length} rows... All rows will be imported.
                </p>
              )}

              {/* Import Insertion Mode Options */}
              <div className="bg-slate-50 p-3.5 rounded-lg border border-slate-200 space-y-3">
                <div className="text-xs font-bold uppercase tracking-wider text-slate-700">
                  Import Placement Options:
                </div>
                
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs">
                  <label
                    className={`flex items-start gap-2.5 p-2.5 rounded-lg border cursor-pointer transition-all ${
                      importMode === "replace"
                        ? "bg-indigo-50/80 border-indigo-400 text-indigo-950 font-bold shadow-xs"
                        : "bg-white border-slate-200 text-slate-700 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="importMode"
                      checked={importMode === "replace"}
                      onChange={() => setImportMode("replace")}
                      className="mt-0.5 text-indigo-600 focus:ring-indigo-500"
                    />
                    <div>
                      <div>Replace Entire Sheet</div>
                      <div className="text-[10px] font-normal text-slate-500">
                        Overwrites current {totalCurrentRows} rows with imported data
                      </div>
                    </div>
                  </label>

                  <label
                    className={`flex items-start gap-2.5 p-2.5 rounded-lg border cursor-pointer transition-all ${
                      importMode === "append"
                        ? "bg-indigo-50/80 border-indigo-400 text-indigo-950 font-bold shadow-xs"
                        : "bg-white border-slate-200 text-slate-700 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="importMode"
                      checked={importMode === "append"}
                      onChange={() => setImportMode("append")}
                      className="mt-0.5 text-indigo-600 focus:ring-indigo-500"
                    />
                    <div>
                      <div>Append to End</div>
                      <div className="text-[10px] font-normal text-slate-500">
                        Adds new rows after row #{totalCurrentRows}
                      </div>
                    </div>
                  </label>

                  <label
                    className={`flex items-start gap-2.5 p-2.5 rounded-lg border cursor-pointer transition-all ${
                      importMode === "insert_at"
                        ? "bg-indigo-50/80 border-indigo-400 text-indigo-950 font-bold shadow-xs"
                        : "bg-white border-slate-200 text-slate-700 hover:border-slate-300"
                    }`}
                  >
                    <input
                      type="radio"
                      name="importMode"
                      checked={importMode === "insert_at"}
                      onChange={() => setImportMode("insert_at")}
                      className="mt-0.5 text-indigo-600 focus:ring-indigo-500"
                    />
                    <div>
                      <div>Insert at Row #{selectedRowIndex + 1}</div>
                      <div className="text-[10px] font-normal text-slate-500">
                        Overwrites/inserts starting at current selected cell
                      </div>
                    </div>
                  </label>
                </div>

                {/* Extra blank rows adder */}
                <div className="flex flex-wrap items-center justify-between gap-2 pt-1 text-xs border-t border-slate-200/80">
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-slate-700">Add Extra Blank Rows after import:</span>
                    <div className="flex items-center gap-1">
                      {[0, 5, 10, 20].map((count) => (
                        <button
                          key={count}
                          type="button"
                          onClick={() => setExtraBlankRows(count)}
                          className={`px-2 py-0.5 rounded text-xs font-bold transition-all cursor-pointer ${
                            extraBlankRows === count
                              ? "bg-indigo-600 text-white"
                              : "bg-white border border-slate-300 text-slate-700 hover:bg-slate-100"
                          }`}
                        >
                          +{count}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="bg-slate-100 px-5 py-3 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
          <div className="text-xs text-slate-500 flex items-center gap-1.5">
            <HelpCircle className="h-3.5 w-3.5 text-indigo-600" />
            <span>Direct in-cell pasting (Ctrl+V) is also active anywhere on the sheet table!</span>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 font-bold text-xs px-4 py-2 rounded-lg transition-all cursor-pointer"
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={handleApplyImport}
              disabled={parsedGrid.length === 0}
              className="bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white font-extrabold text-xs px-5 py-2 rounded-lg shadow transition-all cursor-pointer flex items-center gap-1.5"
            >
              <Check className="h-4 w-4" />
              <span>
                Insert {hasHeaderRow ? Math.max(0, parsedGrid.length - 1) : parsedGrid.length} Lines into Sheet
              </span>
            </button>
          </div>
        </div>

      </div>
    </div>
  );
}
