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
  const maxCharsPerLine = isChallan ? 48 : 42;
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
 * Computes exact Excel row height (in points) ensuring 25 items fit cleanly on an A4 page.
 */
export const getItemRowHeight = (visualLines: number): number => {
  if (visualLines <= 1) return 17.5;
  if (visualLines === 2) return 29;
  if (visualLines === 3) return 40;
  return 17.5 + (visualLines - 1) * 12;
};

export interface ExcelPageChunk {
  rows: QuotationRow[];
  startSlIndex: number;
  isLastPage: boolean;
}

/**
 * Dynamically paginates items so up to 25 items fit per page perfectly on standard A4.
 * When items reach maximum page height capacity (e.g. multi-line entries),
 * remaining items are automatically transferred to the next page.
 */
export const paginateRowsForExcel = (
  allRows: QuotationRow[],
  docType: "quotation" | "challan" | "invoice"
): ExcelPageChunk[] => {
  // If no rows, produce 1 standard empty sheet
  if (allRows.length === 0) {
    return [{ rows: [], startSlIndex: 1, isLastPage: true }];
  }

  const isChallan = docType === "challan";
  const isInvoice = docType === "invoice";
  const REGULAR_PAGE_BUDGET = 460; // vertical points budget for items
  const LAST_PAGE_BUDGET = isChallan ? 460 : isInvoice ? 410 : 445; // budget allowing totals & signature block
  const MAX_ITEMS_PER_PAGE = 25; // Standard 25 items capacity per page

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

    // Check if adding this item exceeds budget or max 25 items limit
    if (
      currentChunk.length > 0 &&
      (currentHeight + rowHeight > budgetForCurrentPage || currentChunk.length >= MAX_ITEMS_PER_PAGE)
    ) {
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
    // If last chunk alone exceeds the last page budget (due to several long multi-line items)
    if (!isChallan && currentHeight > LAST_PAGE_BUDGET && currentChunk.length > 1) {
      let splitIdx = currentChunk.length - 1;
      let runningHeight = 0;
      for (let j = 0; j < currentChunk.length; j++) {
        const lines = calculateItemVisualLines(currentChunk[j].desc, isChallan);
        runningHeight += getItemRowHeight(lines);
        if (runningHeight > REGULAR_PAGE_BUDGET) {
          splitIdx = Math.max(1, j);
          break;
        }
      }
      const part1 = currentChunk.slice(0, splitIdx);
      const part2 = currentChunk.slice(splitIdx);
      if (part1.length > 0) {
        chunks.push({
          rows: part1,
          startSlIndex: currentSl,
          isLastPage: false,
        });
        currentSl += part1.length;
      }
      chunks.push({
        rows: part2,
        startSlIndex: currentSl,
        isLastPage: true,
      });
    } else {
      chunks.push({
        rows: currentChunk,
        startSlIndex: currentSl,
        isLastPage: true,
      });
    }
  }

  // Update strictly the last chunk
  chunks.forEach((chunk, idx) => {
    chunk.isLastPage = idx === chunks.length - 1;
  });

  return chunks;
};

/**
 * Builds a single pristine worksheet matching the exact visual component.
 * Dynamic item height, clean non-overlapping text, standard editable grid lines,
 * and authentic document borders.
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
      top: 0.3,
      bottom: 0.3,
      header: 0.1,
      footer: 0.15,
    },
    horizontalCentered: true,
    showGridLines: false, // Clean print output
  };

  // Add centered Page number in Excel print footer
  worksheet.headerFooter = {
    oddFooter: totalPages > 1 ? `&CPage ${pageIndex} of ${totalPages}` : "&CPage &P of &N",
    evenFooter: totalPages > 1 ? `&CPage ${pageIndex} of ${totalPages}` : "&CPage &P of &N",
  };

  (worksheet.properties as any).pageSetUpPr = { fitToPage: true };

  const isChallan = docType === "challan";
  const isInvoice = docType === "invoice";
  const totalCols = isChallan ? 4 : 6;
  const lastColLetter = isChallan ? "D" : "F";

  // Set precise column widths to ensure full readability and no clipped text
  if (isChallan) {
    worksheet.columns = [
      { key: "A", width: 7.5 },  // SL / Logo
      { key: "B", width: 56.0 }, // Description / Company info
      { key: "C", width: 16.0 }, // Qty / Right Box Label
      { key: "D", width: 22.0 }, // Unit / Right Box Value
    ];
  } else {
    worksheet.columns = [
      { key: "A", width: 7.0 },  // SL / Logo
      { key: "B", width: 50.0 }, // Description / Company info
      { key: "C", width: 9.0 },  // Qty / Right Box Label (Part 1)
      { key: "D", width: 10.5 }, // Unit / Right Box Label (Part 2) -> C+D = 19.5 (Fits "REQUISITION NO.:" easily)
      { key: "E", width: 12.0 }, // Price / Right Box Value (Part 1)
      { key: "F", width: 14.5 }, // Amount / Right Box Value (Part 2) -> E+F = 26.5
    ];
  }

  // Keep Excel standard grid lines visible on screen so the user can easily view and edit cells
  worksheet.views = [{ showGridLines: true }];

  // 1. Add Logo in merged A1:A4
  worksheet.mergeCells("A1:A4");
  if (logoId !== null) {
    worksheet.addImage(logoId, {
      tl: { col: 0.05, row: 0.1 },
      ext: { width: 50, height: 50 },
    });
  }

  // 2. Company title & services on the left header (Rows 1 to 4)
  worksheet.getRow(1).height = 20;
  worksheet.getRow(2).height = 14;
  worksheet.getRow(3).height = 13;
  worksheet.getRow(4).height = 13;

  const companyTitleCell = worksheet.getCell("B1");
  companyTitleCell.value = "COMILLA TRADERS";
  companyTitleCell.font = { name: "Arial", size: 15, bold: true, color: { argb: "000000" } };
  companyTitleCell.alignment = { vertical: "middle", horizontal: "left" };

  const companySub1 = worksheet.getCell("B2");
  companySub1.value = "Ship Chandlers, General Order Suppliers & Importers";
  companySub1.font = { name: "Arial", size: 8, bold: true, color: { argb: "334155" } };
  companySub1.alignment = { vertical: "middle", horizontal: "left" };

  const companySub2 = worksheet.getCell("B3");
  companySub2.value = "All kinds of Marine Safety, Deck, Engine, Cabin, Electrical, Provision & Bond Store";
  companySub2.font = { name: "Arial", size: 7.5, color: { argb: "64748B" } };
  companySub2.alignment = { vertical: "middle", horizontal: "left" };

  // 3. Contact information on the right header (Merged across C..F or C..D for broad unclipped display)
  const contactStartCol = "C";

  worksheet.mergeCells(`${contactStartCol}1:${lastColLetter}1`);
  const contact1 = worksheet.getCell(`${contactStartCol}1`);
  contact1.value = "Office: Jubilee Road, Chattogram, Bangladesh";
  contact1.font = { name: "Arial", size: 7.5, color: { argb: "1E293B" } };
  contact1.alignment = { vertical: "middle", horizontal: "right" };

  worksheet.mergeCells(`${contactStartCol}2:${lastColLetter}2`);
  const contact2 = worksheet.getCell(`${contactStartCol}2`);
  contact2.value = "Helplines: 01819315746, 01712-900431";
  contact2.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "1E293B" } };
  contact2.alignment = { vertical: "middle", horizontal: "right" };

  worksheet.mergeCells(`${contactStartCol}3:${lastColLetter}3`);
  const contact3 = worksheet.getCell(`${contactStartCol}3`);
  contact3.value = "Official Email: comillatraders@gmail.com";
  contact3.font = { name: "Arial", size: 7.5, color: { argb: "1E293B" } };
  contact3.alignment = { vertical: "middle", horizontal: "right" };

  worksheet.mergeCells(`${contactStartCol}4:${lastColLetter}4`);
  const contact4 = worksheet.getCell(`${contactStartCol}4`);
  contact4.value = "CHATTOGRAM • BANGLADESH";
  contact4.font = { name: "Arial", size: 8, bold: true, color: { argb: "4338CA" } };
  contact4.alignment = { vertical: "middle", horizontal: "right" };

  // Solid line under the header (Row 5)
  for (let c = 1; c <= totalCols; c++) {
    const cell = worksheet.getCell(5, c);
    cell.border = {
      bottom: { style: "medium", color: { argb: "000000" } },
    };
  }
  worksheet.getRow(5).height = 4;
  worksheet.getRow(6).height = 5;

  // 4. Document Title in Row 7
  worksheet.mergeCells(`A7:${lastColLetter}7`);
  const titleCell = worksheet.getCell("A7");
  const baseTitle = isChallan
    ? "DELIVERY CHALLAN"
    : isInvoice
    ? "INVOICE"
    : "QUOTATION";

  titleCell.value = totalPages > 1 ? `${baseTitle} (PAGE ${pageIndex} OF ${totalPages})` : baseTitle;
  titleCell.font = { name: "Arial", size: 12, bold: true, color: { argb: "000000" } };
  titleCell.alignment = { vertical: "middle", horizontal: "center" };
  worksheet.getRow(7).height = 20;
  worksheet.getRow(8).height = 5;

  // 5. Metadata Boxes (Rows 10 to 14)
  worksheet.mergeCells("A10:B10");
  worksheet.getCell("A10").value = "MESSERS:";
  worksheet.getCell("A10").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };

  worksheet.mergeCells("A11:B11");
  const messersCell = worksheet.getCell("A11");
  messersCell.value = messers || "";
  messersCell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
  messersCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  worksheet.mergeCells("A12:B12");
  worksheet.getCell("A12").value = "ADDRESS:";
  worksheet.getCell("A12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };

  worksheet.mergeCells("A13:B14");
  const addrCell = worksheet.getCell("A13");
  addrCell.value = address || "";
  addrCell.font = { name: "Arial", size: 8, color: { argb: "000000" } };
  addrCell.alignment = { vertical: "top", wrapText: true };
  addrCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  // Left Box outer borders
  for (let r = 10; r <= 14; r++) {
    worksheet.getRow(r).height = 14;
    for (let c = 1; c <= 2; c++) {
      const cell = worksheet.getCell(r, c);
      const cellBorders: any = { ...cell.border };
      if (r === 10) cellBorders.top = { style: "thin", color: { argb: "000000" } };
      if (r === 14) cellBorders.bottom = { style: "thin", color: { argb: "000000" } };
      if (c === 1) cellBorders.left = { style: "thin", color: { argb: "000000" } };
      if (c === 2) cellBorders.right = { style: "thin", color: { argb: "000000" } };
      cell.border = cellBorders;
    }
  }

  // Right Box:
  if (isInvoice) {
    // Row 10: INVOICE NO
    worksheet.mergeCells("C10:D10");
    worksheet.getCell("C10").value = "INVOICE NO.:";
    worksheet.getCell("C10").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E10:F10");
    const iVal = worksheet.getCell("E10");
    iVal.value = invoiceNo || "";
    iVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    iVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 11: DATE
    worksheet.mergeCells("C11:D11");
    worksheet.getCell("C11").value = "DATE:";
    worksheet.getCell("C11").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E11:F11");
    const dVal = worksheet.getCell("E11");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 12: CHALLAN NO
    worksheet.mergeCells("C12:D12");
    worksheet.getCell("C12").value = "CHALLAN NO.:";
    worksheet.getCell("C12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E12:F12");
    const cVal = worksheet.getCell("E12");
    cVal.value = challanNo || "";
    cVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    cVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 13: REQUISITION NO
    worksheet.mergeCells("C13:D13");
    worksheet.getCell("C13").value = "REQUISITION NO.:";
    worksheet.getCell("C13").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E13:F13");
    const rVal = worksheet.getCell("E13");
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 14: P.O. NUMBER
    worksheet.mergeCells("C14:D14");
    worksheet.getCell("C14").value = "P.O. NUMBER:";
    worksheet.getCell("C14").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E14:F14");
    const pVal = worksheet.getCell("E14");
    pVal.value = poNumber || "";
    pVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    pVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
  } else if (isChallan) {
    // Row 10: CHALLAN NO
    worksheet.getCell("C10").value = "CHALLAN NO.:";
    worksheet.getCell("C10").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    const cVal = worksheet.getCell("D10");
    cVal.value = challanNo || "";
    cVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    cVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 11: DATE
    worksheet.getCell("C11").value = "DATE:";
    worksheet.getCell("C11").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    const dVal = worksheet.getCell("D11");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Row 12: REQUISITION NO
    worksheet.getCell("C12").value = "REQUISITION NO.:";
    worksheet.getCell("C12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    const rVal = worksheet.getCell("D12");
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    // Rows 13, 14: blank dotted lines
    for (let r = 13; r <= 14; r++) {
      worksheet.getCell(`D${r}`).border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    }
  } else {
    // Quotation
    worksheet.mergeCells("C10:D10");
    worksheet.getCell("C10").value = "DATE:";
    worksheet.getCell("C10").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E10:F10");
    const dVal = worksheet.getCell("E10");
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    worksheet.mergeCells("C11:D11");
    worksheet.getCell("C11").value = "REQUISITION NO.:";
    worksheet.getCell("C11").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    worksheet.mergeCells("E11:F11");
    const rVal = worksheet.getCell("E11");
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    for (let r = 12; r <= 14; r++) {
      worksheet.mergeCells(`E${r}:F${r}`);
      worksheet.getCell(`E${r}`).border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    }
  }

  // Right Box outer borders
  const rStartCol = 3;
  for (let r = 10; r <= 14; r++) {
    for (let c = rStartCol; c <= totalCols; c++) {
      const cell = worksheet.getCell(r, c);
      const cellBorders: any = { ...cell.border };
      if (r === 10) cellBorders.top = { style: "thin", color: { argb: "000000" } };
      if (r === 14) cellBorders.bottom = { style: "thin", color: { argb: "000000" } };
      if (c === rStartCol) cellBorders.left = { style: "thin", color: { argb: "000000" } };
      if (c === totalCols) cellBorders.right = { style: "thin", color: { argb: "000000" } };
      cell.border = cellBorders;
    }
  }

  worksheet.getRow(15).height = 5;

  // 6. Main Table Header (Row 16)
  const headerRow = worksheet.getRow(16);
  headerRow.height = 19;

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

  // 7. Table rows filling (Dynamic height & clean text wrap)
  let currentRowNum = 17;
  const numItemsOnThisPage = pageRows.length;
  // Pad table to 25 rows on single-page document for standard 25-item template look
  const displayRowCount = totalPages === 1 ? Math.max(numItemsOnThisPage, 25) : Math.max(numItemsOnThisPage, 15);

  for (let idx = 0; idx < displayRowCount; idx++) {
    const r = worksheet.getRow(currentRowNum);
    const rowData = pageRows[idx];

    const slVal = startSlIndex + idx;
    const descVal = rowData ? rowData.desc : "";
    const qtyVal = rowData && rowData.qty ? parseFloat(String(rowData.qty).replace(/,/g, "")) : "";
    const unitVal = rowData ? rowData.unit : "";

    // Dynamic row height based on visual lines in description
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
      const excelStartRow = (region.startRow - pageStartRowIdx) + 17;
      const excelEndRow = (region.endRow - pageStartRowIdx) + 17;
      const excelStartCol = region.startCol + 2;
      const excelEndCol = region.endCol + 2;

      if (
        excelStartRow >= 17 &&
        excelStartCol >= 1 &&
        excelEndRow < 17 + displayRowCount &&
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

  // 8. Totals & Words Block (Strictly rendered ONLY on the FINAL page, NEVER on intermediate pages)
  if (!isChallan && isLastPage) {
    const totalRow = currentRowNum;
    const numTotalRows = isInvoice ? 4 : 1;

    for (let rOffset = 0; rOffset < numTotalRows; rOffset++) {
      worksheet.getRow(totalRow + rOffset).height = 18;
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

    // Apply medium borders to word block
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

    const sumRange = `F17:F${totalRow - 1}`;

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
      grandValCell.font = { name: "Arial", size: 9, bold: true };
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
      // Row 1: TOTAL for Quotation
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

  // 9. Signatures Area
  worksheet.getRow(currentRowNum).height = 10;
  currentRowNum++;

  // "For Comilla Traders" row (Quotation & Invoice)
  if (!isChallan) {
    const authHeaderRow = currentRowNum;
    worksheet.getRow(authHeaderRow).height = 14;
    worksheet.mergeCells(`E${authHeaderRow}:F${authHeaderRow}`);
    const authTitle = worksheet.getCell(`E${authHeaderRow}`);
    authTitle.value = "For Comilla Traders";
    authTitle.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    authTitle.alignment = { vertical: "middle", horizontal: "center" };
    currentRowNum++;
  }

  // Spacer between header & signature lines
  worksheet.getRow(currentRowNum).height = 18;
  const sigRow = currentRowNum + 1;
  worksheet.getRow(sigRow).height = 16;

  // Receiver's Signature (Merged A to B with a single top line for signing)
  worksheet.mergeCells(`A${sigRow}:B${sigRow}`);
  const recSig = worksheet.getCell(`A${sigRow}`);
  recSig.value = "Receiver's Signature";
  recSig.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
  recSig.alignment = { vertical: "middle", horizontal: "center" };
  recSig.border = { top: { style: "thin", color: { argb: "000000" } } };

  // Authorized Signature (Merged right columns with a single top line for signing)
  const authColStart = isChallan ? "C" : "E";
  const authColEnd = isChallan ? "D" : "F";
  worksheet.mergeCells(`${authColStart}${sigRow}:${authColEnd}${sigRow}`);
  const authSig = worksheet.getCell(`${authColStart}${sigRow}`);
  authSig.value = "Authorized Signature";
  authSig.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
  authSig.alignment = { vertical: "middle", horizontal: "center" };
  authSig.border = { top: { style: "thin", color: { argb: "000000" } } };

  // Add Official Stamp precisely centered over the Authorized Signature section (between "For Comilla Traders" and "Authorized Signature")
  if (!isChallan && stampId !== null) {
    // Columns E and F span from col 4.0 to 6.0 (total width 26.5 chars). Center is around col 4.85
    worksheet.addImage(stampId, {
      tl: { col: 4.85, row: sigRow - 2.05 },
      ext: { width: 68, height: 68 },
    });
  }

  // 10. Footer Disclaimer Notice
  const noticeRow = sigRow + 2;
  worksheet.getRow(sigRow + 1).height = 6;
  worksheet.getRow(noticeRow).height = 14;
  worksheet.mergeCells(`A${noticeRow}:${lastColLetter}${noticeRow}`);
  const noticeCell = worksheet.getCell(`A${noticeRow}`);
  noticeCell.value = "ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE.";
  noticeCell.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "000000" } };
  noticeCell.alignment = { vertical: "middle", horizontal: "center" };
};

/**
 * Generates complete multi-sheet Excel Workbook for the currently selected document type.
 * Automatically analyzes content height and paginates items dynamically when maximum page capacity is reached.
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

  // Filter out purely blank trailing rows if user has extra empty slots at the bottom,
  // but keep all filled rows
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

  // Dynamic capacity-aware pagination
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
