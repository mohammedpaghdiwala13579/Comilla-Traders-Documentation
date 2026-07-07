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
}
