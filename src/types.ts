export interface QuotationRow {
  sl: number;
  desc: string;
  qty: string;
  unit: string;
  price: string;
  amount: number;
}

export interface MergedRegion {
  id: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

export interface CellBorders {
  top?: string;
  bottom?: string;
  left?: string;
  right?: string;
}

export interface CellFormat {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: "none" | "single" | "double";
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
  indent?: number;
  orientation?: "horizontal" | "angle-up" | "angle-down" | "vertical" | "rotate-up" | "rotate-down";
  color?: string;
  bgColor?: string;
  borders?: CellBorders;
}

export type CellFormatMap = Record<string, CellFormat>;

export interface SavedDocument {
  id: string;
  name: string;
  createdAt: string;
  updatedAt: string;
  docType: "quotation" | "challan" | "invoice";
  dateVal: string;
  messers: string;
  address: string;
  requisitionNo: string;
  challanNo?: string;
  invoiceNo?: string;
  poNumber?: string;
  rows: QuotationRow[];
  mergedRegions: MergedRegion[];
  cellFormats?: CellFormatMap;
  vatPercent?: number;
  transportationFee?: number;
}
