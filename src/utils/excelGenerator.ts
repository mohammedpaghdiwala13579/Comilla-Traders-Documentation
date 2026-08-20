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
 * Builds a single pristine worksheet matching the exact visual component.
 * Configured for 30 items per page with perfect A4 fit-to-1-page print setup.
 * Completely eliminates overlapping text/logos, removes intermediate subtotal rows,
 * and cleans up any gridlines in the footer/signature zone.
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
    fitToHeight: 1, // Fits exactly on 1 page of A4
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.3,
      bottom: 0.3,
      header: 0.1,
      footer: 0.15,
    },
    horizontalCentered: true,
    showGridLines: false, // Clean look without default gridlines spilling into footer
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

  // Set precise column widths
  if (isChallan) {
    worksheet.columns = [
      { key: "A", width: 7.5 },  // SL (Dedicated logo column in header, SL in table)
      { key: "B", width: 68.0 }, // Description
      { key: "C", width: 11.5 }, // Qty
      { key: "D", width: 13.0 }, // Unit
    ];
  } else {
    worksheet.columns = [
      { key: "A", width: 7.0 },  // SL
      { key: "B", width: 51.0 }, // Description
      { key: "C", width: 8.5 },  // Qty
      { key: "D", width: 9.5 },  // Unit
      { key: "E", width: 11.5 }, // Price
      { key: "F", width: 13.5 }, // Amount
    ];
  }

  // Turn off default raw spreadsheet grid lines so our clean borders and blank signature areas stay crisp
  worksheet.views = [{ showGridLines: false }];

  // 1. Add Logo in merged A1:A4 (perfect fit, absolutely no overlap with Company Name in B1)
  worksheet.mergeCells("A1:A4");
  if (logoId !== null) {
    worksheet.addImage(logoId, {
      tl: { col: 0.08, row: 0.1 },
      ext: { width: 52, height: 52 },
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
  companySub2.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "64748B" } };
  companySub2.alignment = { vertical: "middle", horizontal: "left" };

  // 3. Contact information on the right header
  const firstColOfContact = isChallan ? "C" : "E";

  const contact1 = worksheet.getCell(`${firstColOfContact}1`);
  contact1.value = "Office: Jubilee Road, Chattogram, Bangladesh";
  contact1.font = { name: "Arial", size: 7.5, color: { argb: "1E293B" } };
  contact1.alignment = { vertical: "middle", horizontal: "right" };
  worksheet.mergeCells(`${firstColOfContact}1:${lastColLetter}1`);

  const contact2 = worksheet.getCell(`${firstColOfContact}2`);
  contact2.value = "Helplines: 01819315746, 01712-900431";
  contact2.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "1E293B" } };
  contact2.alignment = { vertical: "middle", horizontal: "right" };
  worksheet.mergeCells(`${firstColOfContact}2:${lastColLetter}2`);

  const contact3 = worksheet.getCell(`${firstColOfContact}3`);
  contact3.value = "Official Email: comillatraders@gmail.com";
  contact3.font = { name: "Arial", size: 7.5, color: { argb: "1E293B" } };
  contact3.alignment = { vertical: "middle", horizontal: "right" };
  worksheet.mergeCells(`${firstColOfContact}3:${lastColLetter}3`);

  const contact4 = worksheet.getCell(`${firstColOfContact}4`);
  contact4.value = "CHATTOGRAM • BANGLADESH";
  contact4.font = { name: "Arial", size: 8, bold: true, color: { argb: "4338CA" } };
  contact4.alignment = { vertical: "middle", horizontal: "right" };
  worksheet.mergeCells(`${firstColOfContact}4:${lastColLetter}4`);

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
  const leftBoxWidth = Math.floor(totalCols / 2);
  const leftBoxEndColLetter = leftBoxWidth === 2 ? "B" : "C";
  const rightBoxStartCol = leftBoxWidth + 1;
  const rightBoxStartColLetter = String.fromCharCode(64 + rightBoxStartCol);

  // Left Box: Messers & Address
  worksheet.getCell("A10").value = "MESSERS:";
  worksheet.getCell("A10").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
  worksheet.mergeCells(`A11:${leftBoxEndColLetter}11`);
  const messersCell = worksheet.getCell("A11");
  messersCell.value = messers || "";
  messersCell.font = { name: "Arial", size: 8.5, bold: true, color: { argb: "000000" } };
  messersCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  worksheet.getCell("A12").value = "ADDRESS:";
  worksheet.getCell("A12").font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
  worksheet.mergeCells(`A13:${leftBoxEndColLetter}14`);
  const addrCell = worksheet.getCell("A13");
  addrCell.value = address || "";
  addrCell.font = { name: "Arial", size: 8, color: { argb: "000000" } };
  addrCell.alignment = { vertical: "top", wrapText: true };
  addrCell.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

  for (let r = 10; r <= 14; r++) {
    worksheet.getRow(r).height = 14;
    for (let c = 1; c <= leftBoxWidth; c++) {
      const cell = worksheet.getCell(r, c);
      const cellBorders: any = { ...cell.border };
      if (r === 10) cellBorders.top = { style: "thin", color: { argb: "000000" } };
      if (r === 14) cellBorders.bottom = { style: "thin", color: { argb: "000000" } };
      if (c === 1) cellBorders.left = { style: "thin", color: { argb: "000000" } };
      if (c === leftBoxWidth) cellBorders.right = { style: "thin", color: { argb: "000000" } };
      cell.border = cellBorders;
    }
  }

  // Right Box: Info depending on docType
  if (isInvoice) {
    worksheet.getCell(`${rightBoxStartColLetter}10`).value = "INVOICE NO.:";
    const iVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}10`);
    iVal.value = invoiceNo || "";
    iVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}10:${lastColLetter}10`);
    iVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    worksheet.getCell(`${rightBoxStartColLetter}11`).value = "DATE:";
    const dVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}11`);
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}11:${lastColLetter}11`);
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    worksheet.getCell(`${rightBoxStartColLetter}12`).value = "CHALLAN NO.:";
    const cVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}12`);
    cVal.value = challanNo || "";
    cVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}12:${lastColLetter}12`);
    cVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    worksheet.getCell(`${rightBoxStartColLetter}13`).value = "REQUISITION NO.:";
    const rVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}13`);
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}13:${lastColLetter}13`);
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    worksheet.getCell(`${rightBoxStartColLetter}14`).value = "P.O. NUMBER:";
    const pVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}14`);
    pVal.value = poNumber || "";
    pVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}14:${lastColLetter}14`);
    pVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
  } else if (isChallan) {
    worksheet.getCell(`${rightBoxStartColLetter}10`).value = "CHALLAN NO.:";
    const cVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}10`);
    cVal.value = challanNo || "";
    cVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}10:${lastColLetter}10`);
    cVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    worksheet.getCell(`${rightBoxStartColLetter}11`).value = "DATE:";
    const dVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}11`);
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}11:${lastColLetter}11`);
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    worksheet.getCell(`${rightBoxStartColLetter}12`).value = "REQUISITION NO.:";
    const rVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}12`);
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}12:${lastColLetter}12`);
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    for (let r = 13; r <= 14; r++) {
      worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}${r}:${lastColLetter}${r}`);
      worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}${r}`).border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    }
  } else {
    // Quotation
    worksheet.getCell(`${rightBoxStartColLetter}10`).value = "DATE:";
    const dVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}10`);
    dVal.value = dateVal || "";
    dVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}10:${lastColLetter}10`);
    dVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    worksheet.getCell(`${rightBoxStartColLetter}11`).value = "REQUISITION NO.:";
    const rVal = worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}11`);
    rVal.value = requisitionNo || "";
    rVal.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
    worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}11:${lastColLetter}11`);
    rVal.border = { bottom: { style: "dotted", color: { argb: "64748B" } } };

    for (let r = 12; r <= 14; r++) {
      worksheet.mergeCells(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}${r}:${lastColLetter}${r}`);
      worksheet.getCell(`${String.fromCharCode(rightBoxStartColLetter.charCodeAt(0) + 1)}${r}`).border = { bottom: { style: "dotted", color: { argb: "64748B" } } };
    }
  }

  // Label formatting & borders for Right Box
  for (let r = 10; r <= 14; r++) {
    const lbl = worksheet.getCell(r, rightBoxStartCol);
    lbl.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "475569" } };
    lbl.alignment = { vertical: "middle", horizontal: "left" };

    for (let c = rightBoxStartCol; c <= totalCols; c++) {
      const cell = worksheet.getCell(r, c);
      const cellBorders: any = { ...cell.border };
      if (r === 10) cellBorders.top = { style: "thin", color: { argb: "000000" } };
      if (r === 14) cellBorders.bottom = { style: "thin", color: { argb: "000000" } };
      if (c === rightBoxStartCol) cellBorders.left = { style: "thin", color: { argb: "000000" } };
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

  // 7. Table rows filling (Fixed 30 items per page)
  const ITEMS_PER_PAGE = 30;
  let currentRowNum = 17;

  for (let idx = 0; idx < ITEMS_PER_PAGE; idx++) {
    const r = worksheet.getRow(currentRowNum);
    r.height = 16.5; // Perfect height for 30 items fitting cleanly on A4

    const rowData = pageRows[idx];
    const slVal = startSlIndex + idx;
    const descVal = rowData ? rowData.desc : "";
    const qtyVal = rowData && rowData.qty ? parseFloat(String(rowData.qty).replace(/,/g, "")) : "";
    const unitVal = rowData ? rowData.unit : "";

    // Set SL
    const cellSL = r.getCell(1);
    cellSL.value = rowData || idx < pageRows.length ? slVal : "";
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

    // Cell Borders
    for (let col = 1; col <= totalCols; col++) {
      const cell = r.getCell(col);
      cell.border = {
        top: { style: "thin", color: { argb: "D1D5DB" } },
        bottom: { style: "thin", color: { argb: "D1D5DB" } },
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
  const pageEndRowIdx = pageStartRowIdx + ITEMS_PER_PAGE - 1;

  mergedRegions.forEach((region) => {
    if (region.startRow >= pageStartRowIdx && region.endRow <= pageEndRowIdx) {
      const excelStartRow = (region.startRow - pageStartRowIdx) + 17;
      const excelEndRow = (region.endRow - pageStartRowIdx) + 17;
      const excelStartCol = region.startCol + 2;
      const excelEndCol = region.endCol + 2;

      if (
        excelStartRow >= 17 &&
        excelStartCol >= 1 &&
        excelEndRow < 17 + ITEMS_PER_PAGE &&
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
            top: rOffset === 0 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "CCCCCC" } },
            bottom: rOffset === numTotalRows - 1 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "CCCCCC" } },
            left: c === 5 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "CCCCCC" } },
            right: c === 6 ? { style: "medium", color: { argb: "000000" } } : { style: "thin", color: { argb: "CCCCCC" } },
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
  // Blank spacer row (no borders, clean spacing)
  worksheet.getRow(currentRowNum).height = 12;
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
  worksheet.getRow(currentRowNum).height = 20;
  const sigRow = currentRowNum + 1;
  worksheet.getRow(sigRow).height = 16;

  // Receiver's Signature (Merged A to B with a single clean top line)
  worksheet.mergeCells(`A${sigRow}:B${sigRow}`);
  const recSig = worksheet.getCell(`A${sigRow}`);
  recSig.value = "Receiver's Signature";
  recSig.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
  recSig.alignment = { vertical: "middle", horizontal: "center" };
  recSig.border = { top: { style: "thin", color: { argb: "000000" } } };

  // Authorized Signature (Right side with a single clean top line)
  const authColStart = isChallan ? "C" : "E";
  const authColEnd = isChallan ? "D" : "F";
  worksheet.mergeCells(`${authColStart}${sigRow}:${authColEnd}${sigRow}`);
  const authSig = worksheet.getCell(`${authColStart}${sigRow}`);
  authSig.value = "Authorized Signature";
  authSig.font = { name: "Arial", size: 8, bold: true, color: { argb: "000000" } };
  authSig.alignment = { vertical: "middle", horizontal: "center" };
  authSig.border = { top: { style: "thin", color: { argb: "000000" } } };

  // Add Official Stamp cleanly positioned above the signature line without obscuring text
  if (!isChallan && stampId !== null) {
    worksheet.addImage(stampId, {
      tl: { col: 4.2, row: sigRow - 2.2 },
      ext: { width: 62, height: 62 },
    });
  }

  // 10. Footer Disclaimer Notice
  const noticeRow = sigRow + 2;
  worksheet.getRow(sigRow + 1).height = 8;
  worksheet.getRow(noticeRow).height = 14;
  worksheet.mergeCells(`A${noticeRow}:${lastColLetter}${noticeRow}`);
  const noticeCell = worksheet.getCell(`A${noticeRow}`);
  noticeCell.value = "ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE.";
  noticeCell.font = { name: "Arial", size: 7.5, bold: true, color: { argb: "000000" } };
  noticeCell.alignment = { vertical: "middle", horizontal: "center" };
};

/**
 * Generates complete multi-sheet Excel Workbook for the currently selected document type.
 * Automatically divides items into batches of 30 per sheet with clean headers, no overlapping elements,
 * no intermediate subtotals, and no unwanted gridlines in signature areas.
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

  // Export only the currently selected document type
  const sheetTypes: ("quotation" | "challan" | "invoice")[] = [docType];

  const baseSheetNames: Record<string, string> = {
    quotation: "Quotation",
    challan: "Challan",
    invoice: "Invoice",
  };

  const ITEMS_PER_PAGE = 30;

  // Trim or slice rows - ensure at least 1 page
  const totalRowsCount = rows.length;
  const totalPages = Math.max(1, Math.ceil(totalRowsCount / ITEMS_PER_PAGE));

  sheetTypes.forEach((type) => {
    for (let pageIdx = 0; pageIdx < totalPages; pageIdx++) {
      const startIdx = pageIdx * ITEMS_PER_PAGE;
      const endIdx = startIdx + ITEMS_PER_PAGE;
      const pageRows = rows.slice(startIdx, endIdx);
      const isLastPage = pageIdx === totalPages - 1;
      const pageNumber = pageIdx + 1;

      // Sheet name (e.g., "Quotation" if 1 page, or "Quotation - Pg 1", "Quotation - Pg 2")
      const sheetName =
        totalPages === 1
          ? baseSheetNames[type]
          : `${baseSheetNames[type]} - Pg ${pageNumber}`;

      const ws = workbook.addWorksheet(sheetName);
      buildDocumentWorksheet(
        workbook,
        ws,
        type,
        messers,
        address,
        challanNo,
        dateVal,
        requisitionNo,
        pageRows,
        mergedRegions,
        invoiceNo,
        poNumber,
        vatPercent || 0,
        transportationFee || 0,
        logoId,
        stampId,
        startIdx + 1, // startSlIndex (e.g. 1, 31, 61...)
        pageNumber,
        totalPages,
        isLastPage,
        rows
      );
    }
  });

  return workbook;
};
