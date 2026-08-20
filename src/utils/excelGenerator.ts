import ExcelJS from "exceljs";
import { QuotationRow, MergedRegion } from "../types";
import { numberToWords } from "./numberToWords";

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
  if (!desc || !desc.trim()) return 1;
  const maxCharsPerLine = isChallan ? 54 : 46;
  const parts = desc.split(/\r?\n/);
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
  // Regular intermediate pages: Header (157pt) + Footer (17pt) = 174pt. Available = 560pt.
  // Last page: Available accounts for Totals and Signature/Stamp blocks.
  const REGULAR_PAGE_BUDGET = 560; 
  const LAST_PAGE_BUDGET = isChallan ? 560 : isInvoice ? 475 : 515;

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
  allDocumentRows: QuotationRow[] = []
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

  // 1. Leave 10 average size blank rows at the top (Rows 1 to 10) instead of the previous company header
  for (let r = 1; r <= 10; r++) {
    worksheet.getRow(r).height = 15;
  }

  // 2. Document Title in Row 11 (Clean title)
  worksheet.getRow(11).height = 19;
  worksheet.mergeCells(`A11:${lastColLetter}11`);
  const titleCell = worksheet.getCell("A11");
  const baseTitle = isChallan
    ? "DELIVERY CHALLAN"
    : isInvoice
    ? "INVOICE"
    : "QUOTATION";

  titleCell.value = baseTitle;
  titleCell.font = { name: "Arial", size: 11.5, bold: true, color: { argb: "000000" } };
  titleCell.alignment = { vertical: "middle", horizontal: "center" };

  // 3. Metadata Boxes (Rows 12 to 16 - Compact, tightly placed)
  worksheet.mergeCells("A12:B12");
  worksheet.getCell("A12").value = "MESSERS:";
  worksheet.getCell("A12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };

  worksheet.mergeCells("A13:B13");
  const messersCell = worksheet.getCell("A13");
  messersCell.value = messers || "";
  messersCell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
  messersCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  worksheet.mergeCells("A14:B14");
  worksheet.getCell("A14").value = "ADDRESS:";
  worksheet.getCell("A14").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };

  worksheet.mergeCells("A15:B16");
  const addrCell = worksheet.getCell("A15");
  addrCell.value = address || "";
  addrCell.font = { name: "Arial", size: 8, color: { argb: "000000" } };
  addrCell.alignment = { vertical: "top", wrapText: true };
  addrCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  // Left Box outer borders
  for (let r = 12; r <= 16; r++) {
    worksheet.getRow(r).height = 13;
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

  // Right Box (Rows 12 to 16)
  if (isInvoice) {
    // Row 12: INVOICE NO
    worksheet.mergeCells("C12:D12");
    worksheet.getCell("C12").value = "INVOICE NO.:";
    worksheet.getCell("C12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E12:F12");
    const iVal = worksheet.getCell("E12");
    iVal.value = invoiceNo || "";
    iVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    iVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 13: DATE
    worksheet.mergeCells("C13:D13");
    worksheet.getCell("C13").value = "DATE:";
    worksheet.getCell("C13").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E13:F13");
    const dVal = worksheet.getCell("E13");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

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
    // Row 12: CHALLAN NO
    worksheet.getCell("C12").value = "CHALLAN NO.:";
    worksheet.getCell("C12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    const cVal = worksheet.getCell("D12");
    cVal.value = challanNo || "";
    cVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    cVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 13: DATE
    worksheet.getCell("C13").value = "DATE:";
    worksheet.getCell("C13").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    const dVal = worksheet.getCell("D13");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

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
    // Quotation
    worksheet.mergeCells("C12:D12");
    worksheet.getCell("C12").value = "DATE:";
    worksheet.getCell("C12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E12:F12");
    const dVal = worksheet.getCell("E12");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

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

  // Right Box outer borders
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

  // 5. Table rows filling - Uses space strictly as per text only
  let currentRowNum = 18;
  const numItemsOnThisPage = pageRows.length;
  // Display only the actual items on this page (or 1 minimal row if totally empty)
  const displayRowCount = Math.max(numItemsOnThisPage, 1);

  for (let idx = 0; idx < displayRowCount; idx++) {
    const r = worksheet.getRow(currentRowNum);
    const rowData = pageRows[idx];

    const slVal = startSlIndex + idx;
    const descVal = rowData ? rowData.desc : "";
    const qtyVal = rowData && rowData.qty ? parseFloat(String(rowData.qty).replace(/,/g, "")) : "";
    const unitVal = rowData ? rowData.unit : "";

    // Dynamic compact row height based strictly on actual visual lines
    const visualLines = calculateItemVisualLines(descVal, isChallan);
    r.height = getItemRowHeight(visualLines);

    // Set SL
    const cellSL = r.getCell(1);
    cellSL.value = rowData ? slVal : "";
    cellSL.alignment = { vertical: "middle", horizontal: "center" };
    cellSL.font = { name: "Arial", size: 8 };

    // Set Description
    const cellDesc = r.getCell(2);
    cellDesc.value = descVal;
    cellDesc.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    cellDesc.font = { name: "Arial", size: 8 };

    // Set Qty
    const cellQty = r.getCell(3);
    cellQty.value = qtyVal === "" || isNaN(qtyVal) ? (rowData?.qty || "") : qtyVal;
    cellQty.alignment = { vertical: "middle", horizontal: "center" };
    cellQty.font = { name: "Arial", size: 8 };
    if (typeof qtyVal === "number" && !isNaN(qtyVal)) {
      cellQty.numFmt = "#,##0.00";
    }

    // Set Unit
    const cellUnit = r.getCell(4);
    cellUnit.value = unitVal;
    cellUnit.alignment = { vertical: "middle", horizontal: "center" };
    cellUnit.font = { name: "Arial", size: 8 };

    // Solid Table Cell Borders
    for (let col = 1; col <= totalCols; col++) {
      const cell = r.getCell(col);
      cell.border = {
        top: { style: "thin", color: { argb: "000000" } },
        bottom: { style: "thin", color: { argb: "000000" } },
        left: { style: "thin", color: { argb: "000000" } },
        right: { style: "thin", color: { argb: "000000" } },
      };
    }

    // Set Price & Amount for Quotation / Invoice
    if (!isChallan) {
      const priceVal = rowData && rowData.price ? parseFloat(String(rowData.price).replace(/,/g, "")) : "";

      const cellPrice = r.getCell(5);
      cellPrice.value = priceVal === "" || isNaN(priceVal) ? (rowData?.price || "") : priceVal;
      cellPrice.alignment = { vertical: "middle", horizontal: "right" };
      cellPrice.font = { name: "Arial", size: 8 };
      if (typeof priceVal === "number" && !isNaN(priceVal)) {
        cellPrice.numFmt = "#,##0.00";
      }

      const cellAmount = r.getCell(6);
      if (rowData && (rowData.desc.trim() || rowData.qty.trim() || rowData.price.trim())) {
        cellAmount.value = {
          formula: `=IF(OR(C${currentRowNum}="", E${currentRowNum}=""), 0, C${currentRowNum}*E${currentRowNum})`,
        } as any;
      } else {
        cellAmount.value = "";
      }
      cellAmount.alignment = { vertical: "middle", horizontal: "right" };
      cellAmount.font = { name: "Arial", size: 8 };
      cellAmount.numFmt = "#,##0.00";
    }

    currentRowNum++;
  }

  // Draw medium bottom border on the last grid row
  const lastGridRow = worksheet.getRow(currentRowNum - 1);
  for (let col = 1; col <= totalCols; col++) {
    const cell = lastGridRow.getCell(col);
    cell.border = {
      ...cell.border,
      bottom: { style: "medium", color: { argb: "000000" } },
    };
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

  // 9. Signatures Area - Compact & Centered
  worksheet.getRow(currentRowNum).height = 6;
  currentRowNum++;

  // "For Comilla Traders" row
  const authHeaderRow = currentRowNum;
  if (!isChallan) {
    worksheet.getRow(authHeaderRow).height = 13;
    worksheet.mergeCells(`E${authHeaderRow}:F${authHeaderRow}`);
    const authTitle = worksheet.getCell(`E${authHeaderRow}`);
    authTitle.value = "For Comilla Traders";
    authTitle.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    authTitle.alignment = { vertical: "middle", horizontal: "center" };
    currentRowNum++;
  }

  // Spacer between header & signature line
  worksheet.getRow(currentRowNum).height = 15;
  const sigRow = currentRowNum + 1;
  worksheet.getRow(sigRow).height = 14;

  // Receiver's Signature
  worksheet.mergeCells(`A${sigRow}:B${sigRow}`);
  const recSig = worksheet.getCell(`A${sigRow}`);
  recSig.value = "Receiver's Signature";
  recSig.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
  recSig.alignment = { vertical: "middle", horizontal: "center" };
  recSig.border = { top: { style: "thin", color: { argb: "000000" } } };

  // Authorized Signature
  const authColStart = isChallan ? "C" : "E";
  const authColEnd = isChallan ? "D" : "F";
  worksheet.mergeCells(`${authColStart}${sigRow}:${authColEnd}${sigRow}`);
  const authSig = worksheet.getCell(`${authColStart}${sigRow}`);
  authSig.value = "Authorized Signature";
  authSig.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
  authSig.alignment = { vertical: "middle", horizontal: "center" };
  authSig.border = { top: { style: "thin", color: { argb: "000000" } } };

  // Centered Official Stamp positioned directly between "For Comilla Traders" and "Authorized Signature"
  if (!isChallan && stampId !== null) {
    worksheet.addImage(stampId, {
      tl: { col: 4.72, row: sigRow - 2.05 },
      ext: { width: 66, height: 66 },
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
  transportationFee?: number
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
      effectiveRows
    );
  });

  return workbook;
};
