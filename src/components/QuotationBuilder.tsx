import React, { useState, useEffect, useRef } from "react";
// @ts-ignore
import html2pdf from "html2pdf.js";
import { Download, Printer, Calendar, Save, Trash2, Plus, Minus, Check, RefreshCw, Copy, X, FileSpreadsheet, Layers, ListPlus, ArrowDownToLine, CheckCheck, Scissors, WrapText } from "lucide-react";
import { db } from "../lib/firebase";
import { collection, doc, setDoc, deleteDoc, onSnapshot, query, orderBy } from "firebase/firestore";
import { numberToWords } from "../utils/numberToWords";
import { parseClipboardData, parseTSV, cleanCellText } from "../utils/tsvParser";
import { generateExcelWorkbook } from "../utils/excelGenerator";
import { QuotationRow, MergedRegion, SavedDocument, CellFormat, CellFormatMap, CellBorders } from "../types";
import SavedDocumentsPanel from "./SavedDocumentsPanel";
import ExcelPasteModal from "./ExcelPasteModal";
import ExcelRibbonToolbar from "./ExcelRibbonToolbar";
import RichTextCell from "./RichTextCell";
import { stripHtml, parseNumericInput } from "../utils/textFormatter";

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: null,
      email: null
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// Translate OKLCH colors to standard sRGB for canvas compatibility (used for html2pdf rendering)
function oklchToRgb(l: number, c: number, h: number): [number, number, number] {
  const hRad = (h * Math.PI) / 180;
  const a = c * Math.cos(hRad);
  const b = c * Math.sin(hRad);
  
  const l_ = l + 0.3963377774 * a + 0.2158037573 * b;
  const m_ = l - 0.1055613458 * a - 0.0638541728 * b;
  const s_ = l - 0.0894841775 * a - 1.2914855414 * b;
  
  const l3 = l_ * l_ * l_;
  const m3 = m_ * m_ * m_;
  const s3 = s_ * s_ * s_;
  
  const rLinear = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const gLinear = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const bLinear = -0.0041960863 * l3 - 0.703418614 * m3 + 1.7076146995 * s3;
  
  const toSRGB = (x: number) => {
    const clamped = Math.max(0, Math.min(1, x));
    return clamped <= 0.0031308
      ? Math.round(clamped * 12.92 * 255)
      : Math.round((1.055 * Math.pow(clamped, 1 / 2.4) - 0.055) * 255);
  };
  
  return [toSRGB(rLinear), toSRGB(gLinear), toSRGB(bLinear)];
}

function replaceOklchInCss(cssText: string): string {
  if (!cssText || typeof cssText !== 'string' || !cssText.includes("oklch")) {
    return cssText;
  }
  
  const oklchRegex = /oklch\(([^)]+)\)/g;
  return cssText.replace(oklchRegex, (match, innerText) => {
    try {
      const parts = innerText.trim().split(/\s+/).filter((p: string) => p !== '/');
      if (parts.length >= 3) {
        const lStr = parts[0];
        const cStr = parts[1];
        const hStr = parts[2];
        const aStr = parts[3];
        
        let l = lStr.endsWith('%') ? parseFloat(lStr) / 100 : parseFloat(lStr);
        let c = cStr.endsWith('%') ? parseFloat(cStr) / 100 : parseFloat(cStr);
        let h = hStr.endsWith('deg') ? parseFloat(hStr) : parseFloat(hStr);
        
        let alpha = 1;
        if (aStr) {
          alpha = aStr.endsWith('%') ? parseFloat(aStr) / 100 : parseFloat(aStr);
        }
        
        if (isNaN(l) || isNaN(c) || isNaN(h)) return match;
        
        const [r, g, b] = oklchToRgb(l, c, h);
        if (aStr !== undefined) {
          return `rgba(${r}, ${g}, ${b}, ${alpha})`;
        } else {
          return `rgb(${r}, ${g}, ${b})`;
        }
      }
      return match;
    } catch (e) {
      return match;
    }
  });
}

const DRAFT_STORAGE_KEY = "comilla_active_draft_v2";

const loadInitialDraft = () => {
  if (typeof window === "undefined") return null;
  try {
    const raw = localStorage.getItem(DRAFT_STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
};

export default function QuotationBuilder() {
  const initialDraft = useRef(loadInitialDraft()).current;

  const [docType, setDocType] = useState<"quotation" | "challan" | "invoice">(() => initialDraft?.docType || "quotation");
  const [dateVal, setDateVal] = useState(() => {
    if (initialDraft?.dateVal) return initialDraft.dateVal;
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, "0");
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy = today.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  });
  const [messers, setMessers] = useState(() => initialDraft?.messers || "");
  const [address, setAddress] = useState(() => initialDraft?.address || "");
  const [challanNo, setChallanNo] = useState(() => initialDraft?.challanNo || "");
  const [requisitionNo, setRequisitionNo] = useState(() => initialDraft?.requisitionNo || "");
  const [invoiceNo, setInvoiceNo] = useState(() => initialDraft?.invoiceNo || "");
  const [poNumber, setPoNumber] = useState(() => initialDraft?.poNumber || "");
  const [vatPercent, setVatPercent] = useState<string>(() => initialDraft?.vatPercent !== undefined ? String(initialDraft.vatPercent) : "0");
  const [transportationFee, setTransportationFee] = useState<string>(() => initialDraft?.transportationFee !== undefined ? String(initialDraft.transportationFee) : "0");
  const [isGeneratingExcel, setIsGeneratingExcel] = useState(false);
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);

  // In-app storage & Auto-Save states
  const [savedDocs, setSavedDocs] = useState<SavedDocument[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedTypeFilter, setSelectedTypeFilter] = useState<"all" | "quotation" | "challan" | "invoice">("all");
  const [currentDocId, setCurrentDocId] = useState<string | null>(() => initialDraft?.currentDocId || null);
  const [autoSaveEnabled, setAutoSaveEnabled] = useState<boolean>(() => {
    if (typeof window !== "undefined") {
      const val = localStorage.getItem("comilla_autosave_enabled");
      return val === null ? true : val === "true";
    }
    return true;
  });
  const [lastSavedTime, setLastSavedTime] = useState<string | null>(null);
  const [saveStatus, setSaveStatus] = useState<"idle" | "saving" | "saved" | "error">("idle");

  // Excel Paste Modal and Batch Row Adder states
  const [isExcelModalOpen, setIsExcelModalOpen] = useState(false);
  const [customRowCountInput, setCustomRowCountInput] = useState<string>("10");
  const [targetTotalRowCountInput, setTargetTotalRowCountInput] = useState<string>("");
  const [toastMessage, setToastMessage] = useState<{ text: string; type?: "info" | "success" } | null>(null);

  const showToast = (text: string, type: "info" | "success" = "success") => {
    setToastMessage({ text, type });
    setTimeout(() => {
      setToastMessage(null);
    }, 3500);
  };

  // Grid rows: starts with saved draft rows or 20 blank rows by default
  const [rows, setRows] = useState<QuotationRow[]>(() => {
    if (initialDraft?.rows && Array.isArray(initialDraft.rows) && initialDraft.rows.length > 0) {
      return initialDraft.rows;
    }
    const initialRows: QuotationRow[] = [];
    for (let i = 1; i <= 20; i++) {
      initialRows.push({
        sl: i,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      });
    }
    return initialRows;
  });

  const [mergedRegions, setMergedRegions] = useState<MergedRegion[]>(() => initialDraft?.mergedRegions || []);
  const [cellFormats, setCellFormats] = useState<CellFormatMap>(() => initialDraft?.cellFormats || {});
  const [selectedRowIndex, setSelectedRowIndex] = useState<number>(0);
  const [selectedCell, setSelectedCell] = useState<{ rowIndex: number; colIndex: number } | null>({ rowIndex: 0, colIndex: 0 });
  const [isSelecting, setIsSelecting] = useState<boolean>(false);
  const [selectionStart, setSelectionStart] = useState<{ rowIndex: number; colIndex: number } | null>({ rowIndex: 0, colIndex: 0 });
  const [selectionEnd, setSelectionEnd] = useState<{ rowIndex: number; colIndex: number } | null>({ rowIndex: 0, colIndex: 0 });

  const [contextMenu, setContextMenu] = useState<{
    visible: boolean;
    x: number;
    y: number;
    rowIndex: number;
    colIndex: number;
  } | null>(null);

  const dateRef = useRef<HTMLInputElement>(null);
  const autoSaveTimerRef = useRef<NodeJS.Timeout | null>(null);

  const triggerDatePicker = () => {
    if (dateRef.current) {
      try {
        dateRef.current.showPicker();
      } catch (e) {
        dateRef.current.focus();
        dateRef.current.click();
      }
    }
  };

  const handleDatePickerChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const dateStr = e.target.value;
    if (!dateStr) return;
    const [yyyy, mm, dd] = dateStr.split("-");
    setDateVal(`${dd}/${mm}/${yyyy}`);
  };

  // Listen to Firestore documents
  useEffect(() => {
    const q = query(collection(db, "documents"), orderBy("updatedAt", "desc"));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const docs: SavedDocument[] = [];
      snapshot.forEach((doc) => {
        const data = doc.data();
        
        const docRows = (data.rows || []).map((r: any) => ({
          sl: Number(r.sl) || 0,
          desc: String(r.desc ?? ""),
          qty: String(r.qty ?? ""),
          unit: String(r.unit ?? ""),
          price: String(r.price ?? ""),
          amount: Number(r.amount) || 0,
        }));
        
        const docMergedRegions: MergedRegion[] = Array.isArray(data.mergedRegions)
          ? data.mergedRegions.map((m: any) => ({
              id: String(m.id ?? `region-${Math.random().toString(36).substring(2, 9)}`),
              startRow: Number(m.startRow) || 0,
              endRow: Number(m.endRow) || 0,
              startCol: Number(m.startCol) ?? 0,
              endCol: Number(m.endCol) ?? 0,
            }))
          : [];
          
        docs.push({
          id: doc.id,
          name: data.name || "",
          createdAt: data.createdAt || "",
          updatedAt: data.updatedAt || "",
          docType: data.docType || "quotation",
          dateVal: data.dateVal || "",
          messers: data.messers || "",
          address: data.address || "",
          challanNo: data.challanNo || "",
          requisitionNo: data.requisitionNo || "",
          invoiceNo: data.invoiceNo || "",
          poNumber: data.poNumber || "",
          rows: docRows,
          mergedRegions: docMergedRegions,
          cellFormats: (data.cellFormats as CellFormatMap) || {},
          vatPercent: data.vatPercent,
          transportationFee: data.transportationFee
        });
      });
      setSavedDocs(docs);
    }, (error) => {
      handleFirestoreError(error, OperationType.GET, "documents");
    });

    return () => unsubscribe();
  }, []);

  const generateUUID = () => {
    return 'doc-' + Math.random().toString(36).substring(2, 11) + '-' + Date.now();
  };

  const saveCurrentDocToApp = async (customName?: string) => {
    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = null;
    }

    const now = new Date().toISOString();
    let docIdentifier = "";
    if (docType === "challan" && challanNo) {
      docIdentifier = ` (Challan #${challanNo})`;
    } else if (docType === "invoice" && invoiceNo) {
      docIdentifier = ` (Invoice #${invoiceNo})`;
    }

    const docTypeLabel = docType === "invoice" ? "Invoice" : docType === "challan" ? "Challan" : "Quotation";
    const defaultName = `${docTypeLabel}${docIdentifier} - ${messers || "Unnamed Client"} (${dateVal})`;
    const nameToUse = customName || savedDocs.find(d => d.id === currentDocId)?.name || defaultName;

    const docId = currentDocId || generateUUID();

    const sanitizedRows = rows.map(r => ({
      sl: Number(r.sl) || 0,
      desc: String(r.desc ?? ""),
      qty: String(r.qty ?? ""),
      unit: String(r.unit ?? ""),
      price: String(r.price ?? ""),
      amount: Number(r.amount) || 0
    }));

    const sanitizedMergedRegions = mergedRegions.map(m => ({
      id: String(m.id),
      startRow: Number(m.startRow) || 0,
      endRow: Number(m.endRow) || 0,
      startCol: Number(m.startCol) ?? 0,
      endCol: Number(m.endCol) ?? 0
    }));

    const docData: SavedDocument = {
      id: docId,
      name: String(nameToUse || "Unnamed Document"),
      createdAt: String(savedDocs.find(d => d.id === currentDocId)?.createdAt || now),
      updatedAt: String(now),
      docType: docType as "quotation" | "challan" | "invoice",
      dateVal: String(dateVal || ""),
      messers: String(messers || ""),
      address: String(address || ""),
      challanNo: String(challanNo || ""),
      requisitionNo: String(requisitionNo || ""),
      invoiceNo: String(invoiceNo || ""),
      poNumber: String(poNumber || ""),
      rows: sanitizedRows,
      mergedRegions: sanitizedMergedRegions,
      cellFormats: { ...cellFormats },
      vatPercent: parseFloat(vatPercent) || 0,
      transportationFee: parseFloat(transportationFee) || 0
    };

    setSaveStatus("saving");
    try {
      await setDoc(doc(db, "documents", docId), docData);
      if (!currentDocId) {
        setCurrentDocId(docId);
      }
      const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
      setLastSavedTime(timeStr);
      setSaveStatus("saved");
      setTimeout(() => setSaveStatus("idle"), 3000);
    } catch (e) {
      console.error("Error saving document:", e);
      setSaveStatus("error");
      setTimeout(() => setSaveStatus("idle"), 3000);
      handleFirestoreError(e, OperationType.WRITE, `documents/${docId}`);
    }
  };

  const resetSheetFields = () => {
    setDocType("quotation");
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, "0");
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy = today.getFullYear();
    setDateVal(`${dd}/${mm}/${yyyy}`);
    setMessers("");
    setAddress("");
    setChallanNo("");
    setRequisitionNo("");
    setInvoiceNo("");
    setPoNumber("");
    setVatPercent("0");
    setTransportationFee("0");
    
    const initialRows: QuotationRow[] = [];
    for (let i = 1; i <= 20; i++) {
      initialRows.push({
        sl: i,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      });
    }
    setRows(initialRows);
    setMergedRegions([]);
    setCellFormats({});
    setCurrentDocId(null);
    setLastSavedTime(null);
    if (typeof window !== "undefined") {
      localStorage.removeItem(DRAFT_STORAGE_KEY);
    }
  };

  const loadSavedDoc = (doc: SavedDocument) => {
    setDocType(doc.docType);
    setDateVal(doc.dateVal);
    setMessers(doc.messers);
    setAddress(doc.address);
    setChallanNo(doc.challanNo || "");
    setRequisitionNo(doc.requisitionNo || "");
    setInvoiceNo(doc.invoiceNo || "");
    setPoNumber(doc.poNumber || "");
    setRows(doc.rows.map(r => ({ ...r })));
    setMergedRegions((doc.mergedRegions || []).map(m => ({ ...m })));
    setCellFormats(doc.cellFormats ? { ...doc.cellFormats } : {});
    setCurrentDocId(doc.id);
    setLastSavedTime(null);
    setVatPercent(doc.vatPercent !== undefined ? String(doc.vatPercent) : "0");
    setTransportationFee(doc.transportationFee !== undefined ? String(doc.transportationFee) : "0");

    if (typeof window !== "undefined") {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  const deleteSavedDoc = async (id: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    if (window.confirm("Are you sure you want to delete this saved document from the online database?")) {
      try {
        await deleteDoc(doc(db, "documents", id));
        if (currentDocId === id) {
          resetSheetFields();
        }
      } catch (e) {
        console.error("Error deleting document:", e);
        handleFirestoreError(e, OperationType.DELETE, `documents/${id}`);
      }
    }
  };

  const renameSavedDoc = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const documentObj = savedDocs.find(d => d.id === id);
    if (!documentObj) return;
    const newName = window.prompt("Rename this document:", documentObj.name);
    if (newName && newName.trim() !== "") {
      try {
        const updatedData = {
          ...documentObj,
          name: newName.trim(),
          updatedAt: new Date().toISOString()
        };
        await setDoc(doc(db, "documents", id), updatedData);
      } catch (e) {
        console.error("Error renaming document:", e);
        handleFirestoreError(e, OperationType.WRITE, `documents/${id}`);
      }
    }
  };

  const startNewDoc = () => {
    if (window.confirm("Start a new document? Unsaved changes on your active sheet will be overwritten.")) {
      resetSheetFields();
    }
  };

  const duplicateCurrentDoc = async () => {
    const defaultName = `Copy of ${messers ? messers.trim() : "Quotation"} (${dateVal})`;
    const docName = window.prompt("Enter a name for the duplicated copy:", defaultName);
    if (!docName || docName.trim() === "") return;

    setSaveStatus("saving");
    const newId = `doc_${Date.now()}`;
    try {
      const docPayload = {
        id: newId,
        name: docName.trim(),
        docType,
        dateVal,
        messers,
        address,
        challanNo: challanNo || "",
        requisitionNo: requisitionNo || "",
        invoiceNo: invoiceNo || "",
        poNumber: poNumber || "",
        rows: rows.map(r => ({
          sl: r.sl,
          desc: r.desc,
          qty: r.qty,
          unit: r.unit,
          price: r.price,
          amount: r.amount
        })),
        mergedRegions: mergedRegions.map(m => ({ ...m })),
        cellFormats: { ...cellFormats },
        vatPercent: parseFloat(vatPercent) || 0,
        transportationFee: parseFloat(transportationFee) || 0,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };

      await setDoc(doc(db, "documents", newId), docPayload);
      setCurrentDocId(newId);
      setSaveStatus("saved");
      setTimeout(() => setSaveStatus("idle"), 3000);
    } catch (err) {
      console.error("Error duplicating document:", err);
      setSaveStatus("error");
      setTimeout(() => setSaveStatus("idle"), 3000);
      handleFirestoreError(err, OperationType.WRITE, `documents/${newId}`);
    }
  };

  // Debounced Auto-Save
  useEffect(() => {
    if (!autoSaveEnabled) return;

    const hasAnyContent = 
      currentDocId !== null ||
      messers.trim() !== "" || 
      address.trim() !== "" || 
      challanNo.trim() !== "" || 
      invoiceNo.trim() !== "" ||
      poNumber.trim() !== "" ||
      rows.some(r => r.desc.trim() !== "");

    if (!hasAnyContent) return;

    const timer = setTimeout(async () => {
      const docId = currentDocId || generateUUID();
      const now = new Date().toISOString();
      let docIdentifier = "";
      if (docType === "challan" && challanNo) {
        docIdentifier = ` (Challan #${challanNo})`;
      } else if (docType === "invoice" && invoiceNo) {
        docIdentifier = ` (Invoice #${invoiceNo})`;
      }

      const docTypeLabel = docType === "invoice" ? "Invoice" : docType === "challan" ? "Challan" : "Quotation";
      const defaultName = `${docTypeLabel}${docIdentifier} - ${messers || "Unnamed Client"} (${dateVal})`;
      const nameToUse = savedDocs.find(d => d.id === currentDocId)?.name || defaultName;

      const sanitizedRows = rows.map(r => ({
        sl: Number(r.sl) || 0,
        desc: String(r.desc ?? ""),
        qty: String(r.qty ?? ""),
        unit: String(r.unit ?? ""),
        price: String(r.price ?? ""),
        amount: Number(r.amount) || 0
      }));

      const sanitizedMergedRegions = mergedRegions.map(m => ({
        id: String(m.id),
        startRow: Number(m.startRow) || 0,
        endRow: Number(m.endRow) || 0,
        startCol: Number(m.startCol) ?? 0,
        endCol: Number(m.endCol) ?? 0
      }));

      const docData: SavedDocument = {
        id: docId,
        name: String(nameToUse || "Unnamed Document"),
        createdAt: String(savedDocs.find(d => d.id || docId)?.createdAt || now),
        updatedAt: String(now),
        docType: docType as "quotation" | "challan" | "invoice",
        dateVal: String(dateVal || ""),
        messers: String(messers || ""),
        address: String(address || ""),
        challanNo: String(challanNo || ""),
        requisitionNo: String(requisitionNo || ""),
        invoiceNo: String(invoiceNo || ""),
        poNumber: String(poNumber || ""),
        rows: sanitizedRows,
        mergedRegions: sanitizedMergedRegions,
        cellFormats: { ...cellFormats },
        vatPercent: parseFloat(vatPercent) || 0,
        transportationFee: parseFloat(transportationFee) || 0
      };

      setSaveStatus("saving");
      try {
        await setDoc(doc(db, "documents", docId), docData);
        if (!currentDocId) {
          setCurrentDocId(docId);
        }
        const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        setLastSavedTime(timeStr);
        setSaveStatus("saved");
        setTimeout(() => setSaveStatus("idle"), 3000);
      } catch (e) {
        console.error("Auto-save failed:", e);
        setSaveStatus("error");
        setTimeout(() => setSaveStatus("idle"), 3000);
        handleFirestoreError(e, OperationType.WRITE, `documents/${docId}`);
      }
    }, 1500);

    autoSaveTimerRef.current = timer;
    return () => clearTimeout(timer);
  }, [
    docType,
    dateVal,
    messers,
    address,
    challanNo,
    requisitionNo,
    invoiceNo,
    poNumber,
    rows,
    mergedRegions,
    cellFormats,
    autoSaveEnabled,
    currentDocId,
    vatPercent,
    transportationFee
  ]);

  // Active sheet draft persistence to prevent any loss of transportation fee, rows, or details on app reopen
  useEffect(() => {
    if (typeof window === "undefined") return;
    try {
      const draftPayload = {
        docType,
        dateVal,
        messers,
        address,
        challanNo,
        requisitionNo,
        invoiceNo,
        poNumber,
        vatPercent,
        transportationFee,
        currentDocId,
        rows,
        mergedRegions,
        cellFormats
      };
      localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(draftPayload));
    } catch (e) {
      // Ignore localStorage storage quota errors
    }
  }, [
    docType,
    dateVal,
    messers,
    address,
    challanNo,
    requisitionNo,
    invoiceNo,
    poNumber,
    vatPercent,
    transportationFee,
    currentDocId,
    rows,
    mergedRegions,
    cellFormats
  ]);

  // Adjust textarea heights dynamically based on content
  useEffect(() => {
    const textareas = document.querySelectorAll("textarea[data-row]");
    textareas.forEach((ta: any) => {
      ta.style.height = "auto";
      ta.style.height = `${ta.scrollHeight}px`;
    });
  }, [rows]);

  useEffect(() => {
    const handleDismiss = () => setContextMenu(null);
    const handleGlobalMouseUp = () => setIsSelecting(false);
    document.addEventListener("click", handleDismiss);
    window.addEventListener("scroll", handleDismiss, true);
    window.addEventListener("mouseup", handleGlobalMouseUp);
    return () => {
      document.removeEventListener("click", handleDismiss);
      window.removeEventListener("scroll", handleDismiss, true);
      window.removeEventListener("mouseup", handleGlobalMouseUp);
    };
  }, []);

  const handleRowChange = (index: number, field: keyof QuotationRow, value: string) => {
    setRows((prevRows) => {
      const updated = [...prevRows];
      const targetRow = { ...updated[index] };
      
      if (field === "desc" || field === "unit" || field === "qty" || field === "price") {
        (targetRow as any)[field] = value;
      }

      const cleanQty = stripHtml(String(targetRow.qty || ""));
      const cleanPrice = stripHtml(String(targetRow.price || ""));
      const q = parseNumericInput(cleanQty);
      const p = parseNumericInput(cleanPrice);
      targetRow.amount = docType === "challan" ? 0 : q * p;

      updated[index] = targetRow;
      return updated;
    });
  };

  const addRow = () => {
    addRows(1);
  };

  const MAX_LINES_LIMIT = 1000;

  const addRows = (count: number) => {
    if (rows.length >= MAX_LINES_LIMIT) {
      showToast(`Maximum limit of ${MAX_LINES_LIMIT.toLocaleString()} lines reached`);
      return;
    }
    const availableSlots = MAX_LINES_LIMIT - rows.length;
    const requestedCount = Math.max(1, count);
    const qtyToAdd = Math.min(requestedCount, availableSlots);

    if (qtyToAdd <= 0) {
      showToast(`Maximum limit of ${MAX_LINES_LIMIT.toLocaleString()} lines reached`);
      return;
    }

    setRows((prevRows) => {
      const newRows = [...prevRows];
      const startIdx = newRows.length;
      for (let i = 1; i <= qtyToAdd; i++) {
        newRows.push({
          sl: startIdx + i,
          desc: "",
          qty: "",
          unit: "",
          price: "",
          amount: 0,
        });
      }
      return newRows;
    });

    if (qtyToAdd < requestedCount) {
      showToast(`Added ${qtyToAdd} lines (reached ${MAX_LINES_LIMIT.toLocaleString()} line limit)`);
    } else {
      showToast(`Added ${qtyToAdd} line${qtyToAdd > 1 ? "s" : ""}`);
    }
  };

  const setExactTotalRows = (targetCount: number) => {
    const count = Math.max(1, Math.min(MAX_LINES_LIMIT, targetCount));
    setRows((prevRows) => {
      if (count === prevRows.length) return prevRows;
      if (count > prevRows.length) {
        const needed = count - prevRows.length;
        const newRows = [...prevRows];
        const startIdx = newRows.length;
        for (let i = 1; i <= needed; i++) {
          newRows.push({
            sl: startIdx + i,
            desc: "",
            qty: "",
            unit: "",
            price: "",
            amount: 0,
          });
        }
        showToast(`Sheet resized to ${count} lines (+${needed} lines added)`);
        return newRows;
      } else {
        const removed = prevRows.slice(count);
        const hasData = removed.some(r => r.desc.trim() || r.qty.trim() || r.unit.trim() || r.price.trim());
        if (hasData) {
          const ok = window.confirm(`Trimming down to ${count} lines will delete rows containing data. Are you sure you want to proceed?`);
          if (!ok) return prevRows;
        }
        showToast(`Sheet resized to ${count} lines`);
        return prevRows.slice(0, count).map((r, i) => ({ ...r, sl: i + 1 }));
      }
    });
  };

  const trimTrailingBlankRows = () => {
    setRows((prevRows) => {
      let lastNonEmptyIndex = -1;
      for (let i = prevRows.length - 1; i >= 0; i--) {
        const r = prevRows[i];
        if (r.desc.trim() !== "" || r.qty.trim() !== "" || r.unit.trim() !== "" || r.price.trim() !== "") {
          lastNonEmptyIndex = i;
          break;
        }
      }
      if (lastNonEmptyIndex === -1) {
        showToast("Sheet has no filled rows");
        return prevRows.slice(0, 10);
      }
      // Keep at least up to last content row, minimum 5 rows for aesthetics
      const newCount = Math.max(lastNonEmptyIndex + 1, 5);
      const trimmed = prevRows.slice(0, newCount);
      showToast(`Cleaned blank lines at bottom (${trimmed.length} total lines remaining)`);
      return trimmed.map((r, i) => ({ ...r, sl: i + 1 }));
    });
  };

  const removeRow = () => {
    setRows((prevRows) => {
      if (prevRows.length <= 1) return prevRows;
      return prevRows.slice(0, -1);
    });
  };

  const insertRow = (index: number, position: 'above' | 'below') => {
    insertMultipleRows(index, position, 1);
  };

  const insertMultipleRows = (index: number, position: 'above' | 'below', count: number = 1) => {
    if (rows.length >= MAX_LINES_LIMIT) {
      showToast(`Maximum limit of ${MAX_LINES_LIMIT.toLocaleString()} lines reached`);
      return;
    }
    const availableSlots = MAX_LINES_LIMIT - rows.length;
    const qty = Math.min(Math.max(1, count), availableSlots);
    if (qty <= 0) {
      showToast(`Maximum limit of ${MAX_LINES_LIMIT.toLocaleString()} lines reached`);
      return;
    }
    const insertAt = position === 'above' ? index : index + 1;
    setRows((prevRows) => {
      const updated = [...prevRows];
      const newItems: QuotationRow[] = [];
      for (let i = 0; i < qty; i++) {
        newItems.push({
          sl: 0,
          desc: "",
          qty: "",
          unit: "",
          price: "",
          amount: 0,
        });
      }
      updated.splice(insertAt, 0, ...newItems);
      return updated.slice(0, MAX_LINES_LIMIT).map((r, i) => ({
        ...r,
        sl: i + 1
      }));
    });

    setMergedRegions((prevRegions) =>
      prevRegions.map((region) => {
        if (region.startRow >= insertAt) {
          return { ...region, startRow: region.startRow + qty, endRow: region.endRow + qty };
        }
        if (region.endRow >= insertAt) {
          return { ...region, endRow: region.endRow + qty };
        }
        return region;
      })
    );

    showToast(`Inserted ${qty} line${qty > 1 ? "s" : ""} ${position} line ${index + 1}`);
  };

  const handleImportFromExcel = (newRows: QuotationRow[], mode: "replace" | "append" | "insert_at", insertIndex: number = 0) => {
    if (newRows.length === 0) return;
    setRows((prevRows) => {
      let result: QuotationRow[] = [];
      if (mode === "replace") {
        result = newRows.slice(0, MAX_LINES_LIMIT);
        // Ensure at least 20 blank rows if fewer
        while (result.length < Math.min(20, MAX_LINES_LIMIT)) {
          result.push({
            sl: result.length + 1,
            desc: "",
            qty: "",
            unit: "",
            price: "",
            amount: 0,
          });
        }
        setMergedRegions([]);
      } else if (mode === "append") {
        const isCurrentSheetEmpty = prevRows.every(r => !r.desc.trim() && !r.qty.trim() && !r.unit.trim() && !r.price.trim());
        if (isCurrentSheetEmpty && prevRows.length <= 20) {
          result = newRows.slice(0, MAX_LINES_LIMIT);
          while (result.length < Math.min(20, MAX_LINES_LIMIT)) {
            result.push({
              sl: result.length + 1,
              desc: "",
              qty: "",
              unit: "",
              price: "",
              amount: 0,
            });
          }
        } else {
          result = [...prevRows, ...newRows].slice(0, MAX_LINES_LIMIT);
        }
      } else if (mode === "insert_at") {
        const updated = [...prevRows];
        updated.splice(insertIndex, 0, ...newRows);
        result = updated.slice(0, MAX_LINES_LIMIT);
      }
      showToast(`Imported ${Math.min(newRows.length, MAX_LINES_LIMIT)} lines from Excel successfully!`);
      return result.map((r, i) => ({ ...r, sl: i + 1 }));
    });
  };

  const deleteSpecificRow = (index: number) => {
    setRows((prevRows) => {
      if (prevRows.length <= 1) {
        return [{
          sl: 1,
          desc: "",
          qty: "",
          unit: "",
          price: "",
          amount: 0,
        }];
      }
      const updated = prevRows.filter((_, i) => i !== index);
      return updated.map((r, i) => ({
        ...r,
        sl: i + 1
      }));
    });

    setMergedRegions((prevRegions) =>
      prevRegions
        .map((region) => {
          let { startRow, endRow } = region;
          if (startRow > index) startRow -= 1;
          if (endRow >= index) endRow -= 1;
          return { ...region, startRow, endRow };
        })
        .filter((region) => region.endRow >= region.startRow)
    );
  };

  const clearSpecificRow = (index: number) => {
    setRows((prevRows) => {
      const updated = [...prevRows];
      updated[index] = {
        sl: index + 1,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      };
      return updated;
    });
  };

  const COLUMN_FIELD_BY_INDEX: Record<number, "desc" | "qty" | "unit" | "price" | null> = {
    [-1]: null,
    0: "desc",
    1: "qty",
    2: "unit",
    3: "price",
    4: null,
  };

  const getMergeRegionAt = (rowIndex: number, colIndex: number): MergedRegion | undefined => {
    return mergedRegions.find(
      (m) =>
        rowIndex >= m.startRow &&
        rowIndex <= m.endRow &&
        colIndex >= m.startCol &&
        colIndex <= m.endCol
    );
  };

  const getMergeInfo = (rowIndex: number, colIndex: number) => {
    const region = getMergeRegionAt(rowIndex, colIndex);
    if (!region) return { region: undefined, isAnchor: false };
    const isAnchor = rowIndex === region.startRow && colIndex === region.startCol;
    return { region, isAnchor };
  };

  const rangesOverlap = (a: MergedRegion, b: { startRow: number; endRow: number; startCol: number; endCol: number }) => {
    return a.startRow <= b.endRow && a.endRow >= b.startRow && a.startCol <= b.endCol && a.endCol >= b.startCol;
  };

  const mergeSelectedRange = () => {
    if (!selectionStart || !selectionEnd) return;

    const startRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
    const endRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
    const startCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
    const endCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);

    if (startRow === endRow && startCol === endCol) return;

    const candidateRegion = { startRow, endRow, startCol, endCol };
    const overlapping = mergedRegions.find((m) => rangesOverlap(m, candidateRegion));
    if (overlapping) {
      window.alert("Part of this selection is already merged. Unmerge it first, then try again.");
      return;
    }

    setRows((prevRows) => {
      const updated = prevRows.map((r) => ({ ...r }));
      const pieces: string[] = [];

      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const field = COLUMN_FIELD_BY_INDEX[c];
          if (field && updated[r]) {
            const val = String(updated[r][field] ?? "").trim();
            if (val !== "") pieces.push(val);
          }
        }
      }

      const combined = pieces.join(" ");

      for (let r = startRow; r <= endRow; r++) {
        if (!updated[r]) continue;
        for (let c = startCol; c <= endCol; c++) {
          const field = COLUMN_FIELD_BY_INDEX[c];
          if (!field) continue;
          if (r === startRow && c === startCol) {
            (updated[r] as any)[field] = combined;
          } else {
            (updated[r] as any)[field] = "";
          }
        }
        const q = parseNumericInput(updated[r].qty);
        const p = parseNumericInput(updated[r].price);
        updated[r].amount = docType === "challan" ? 0 : q * p;
      }
      return updated;
    });

    const newRegion: MergedRegion = {
      id: generateUUID(),
      startRow,
      endRow,
      startCol,
      endCol,
    };
    setMergedRegions((prev) => [...prev, newRegion]);

    setSelectionStart({ rowIndex: startRow, colIndex: startCol });
    setSelectionEnd({ rowIndex: endRow, colIndex: endCol });
    setSelectedCell({ rowIndex: startRow, colIndex: startCol });
    setSelectedRowIndex(startRow);
  };

  const unmergeRegionAt = (rowIndex: number, colIndex: number) => {
    const region = getMergeRegionAt(rowIndex, colIndex);
    if (!region) return;
    setMergedRegions((prev) => prev.filter((m) => m.id !== region.id));
  };

  const hasRangeSelectionMatchingRegion = (region: MergedRegion) => {
    if (!selectionStart || !selectionEnd) return false;
    const startRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
    const endRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
    const startCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
    const endCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);
    return (
      startRow === region.startRow &&
      endRow === region.endRow &&
      startCol === region.startCol &&
      endCol === region.endCol
    );
  };

  const toggleMergeSelectedRange = () => {
    if (!selectionStart || !selectionEnd) return;
    const { rowIndex, colIndex } = selectionStart;
    const existing = getMergeRegionAt(rowIndex, colIndex);
    if (existing && hasRangeSelectionMatchingRegion(existing)) {
      unmergeRegionAt(rowIndex, colIndex);
    } else if (existing) {
      unmergeRegionAt(rowIndex, colIndex);
    } else {
      mergeSelectedRange();
    }
  };

  const moveRow = (index: number, direction: 'up' | 'down') => {
    setRows((prevRows) => {
      if (direction === 'up' && index === 0) return prevRows;
      if (direction === 'down' && index === prevRows.length - 1) return prevRows;

      const updated = [...prevRows];
      const swapIndex = direction === 'up' ? index - 1 : index + 1;
      
      const temp = updated[index];
      updated[index] = updated[swapIndex];
      updated[swapIndex] = temp;

      return updated.map((r, i) => ({
        ...r,
        sl: i + 1
      }));
    });
  };

  const isCellSelected = (rowIndex: number, colIndex: number) => {
    if (!selectionStart || !selectionEnd) {
      return selectedCell?.rowIndex === rowIndex && selectedCell?.colIndex === colIndex;
    }
    const minRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
    const maxRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
    const minCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
    const maxCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);

    return rowIndex >= minRow && rowIndex <= maxRow && colIndex >= minCol && colIndex <= maxCol;
  };

  const hasRangeSelection = () => {
    if (!selectionStart || !selectionEnd) return false;
    return selectionStart.rowIndex !== selectionEnd.rowIndex || selectionStart.colIndex !== selectionEnd.colIndex;
  };

  const handleCellMouseDown = (e: React.MouseEvent, rowIndex: number, colIndex: number) => {
    if (e.button !== 0) return;

    // Check if user clicked an editable element or within one
    const target = e.target as HTMLElement;
    const isInsideEditable = target && (target.isContentEditable || !!target.closest("[contenteditable='true']") || target.tagName === "INPUT");

    setSelectedCell({ rowIndex, colIndex });
    setSelectedRowIndex(rowIndex);
    setSelectionStart({ rowIndex, colIndex });
    setSelectionEnd({ rowIndex, colIndex });

    if (isInsideEditable) {
      // User is selecting text or placing cursor in an editable cell!
      // Do NOT preventDefault, do NOT blur, and do not start multi-cell drag
      setIsSelecting(false);
      return;
    }

    const isActive = selectedCell?.rowIndex === rowIndex && selectedCell?.colIndex === colIndex;
    setIsSelecting(true);

    if (!isActive) {
      if (document.activeElement instanceof HTMLElement) {
        document.activeElement.blur();
      }
      e.preventDefault();
    }
  };

  const handleCellMouseEnter = (rowIndex: number, colIndex: number) => {
    if (isSelecting) {
      setSelectionEnd({ rowIndex, colIndex });
    }
  };

  const handleCellMouseUp = (e: React.MouseEvent, rowIndex: number, colIndex: number) => {
    setIsSelecting(false);

    const target = e.target as HTMLElement;
    const isInsideEditable = target && (target.isContentEditable || !!target.closest("[contenteditable='true']") || target.tagName === "INPUT");
    if (isInsideEditable) {
      // Preserve active text selection in editable cells
      return;
    }

    const isSingleCell = selectionStart && selectionStart.rowIndex === rowIndex && selectionStart.colIndex === colIndex;
    if (isSingleCell && !hasRangeSelection()) {
      if (colIndex >= 0 && colIndex <= 3) {
        const textarea = document.querySelector(`[data-row="${rowIndex}"][data-col="${colIndex}"]`) as HTMLTextAreaElement | null;
        if (textarea) {
          textarea.focus();
        }
      }
    }
  };

  const getCellClassName = (rowIndex: number, colIndex: number, baseClasses: string) => {
    const isSelected = isCellSelected(rowIndex, colIndex);
    const isActive = selectedCell?.rowIndex === rowIndex && selectedCell?.colIndex === colIndex;
    
    let highlightClass = "";
    if (isActive) {
      highlightClass = "outline outline-2 outline-indigo-600 outline-offset-[-2px] bg-indigo-50/15 print:!outline-none print:!bg-transparent print:!shadow-none z-10 print:z-auto relative";
    } else if (isSelected) {
      highlightClass = "outline outline-1 outline-indigo-400 outline-offset-[-1px] bg-indigo-50/25 print:!outline-none print:!bg-transparent print:!shadow-none z-10 print:z-auto relative shadow-3xs";
    }
    
    return `${baseClasses} ${highlightClass}`;
  };

  const COLUMN_NAMES: Record<number, string> = {
    [-1]: "SL",
    0: "Description",
    1: "Qty",
    2: "Unit",
    3: "Unit Price",
    4: "Amount",
  };

  const activeCellFormat: CellFormat = React.useMemo(() => {
    if (selectedCell) {
      const key = `${selectedCell.rowIndex}_${selectedCell.colIndex}`;
      return cellFormats[key] || {};
    }
    if (selectionStart) {
      const key = `${selectionStart.rowIndex}_${selectionStart.colIndex}`;
      return cellFormats[key] || {};
    }
    return {};
  }, [selectedCell, selectionStart, cellFormats]);

  const selectionSummary = React.useMemo(() => {
    if (selectionStart && selectionEnd && (selectionStart.rowIndex !== selectionEnd.rowIndex || selectionStart.colIndex !== selectionEnd.colIndex)) {
      const minRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
      const maxRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
      const minCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
      const maxCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);
      const totalCells = (maxRow - minRow + 1) * (maxCol - minCol + 1);
      return `Rows ${minRow + 1}–${maxRow + 1}, Cols ${COLUMN_NAMES[minCol] || minCol}–${COLUMN_NAMES[maxCol] || maxCol} (${totalCells} cells)`;
    }
    if (selectedCell) {
      return `Row ${selectedCell.rowIndex + 1}, ${COLUMN_NAMES[selectedCell.colIndex] || "Cell"}`;
    }
    if (selectedRowIndex >= 0) {
      return `Row ${selectedRowIndex + 1} (Entire Row)`;
    }
    return undefined;
  }, [selectedCell, selectionStart, selectionEnd, selectedRowIndex]);

  const handleApplyFormat = (formatUpdate: Partial<CellFormat>) => {
    setCellFormats((prev) => {
      const next = { ...prev };

      if (selectionStart && selectionEnd) {
        const minRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
        const maxRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
        const minCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
        const maxCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);

        for (let r = minRow; r <= maxRow; r++) {
          for (let c = minCol; c <= maxCol; c++) {
            const key = `${r}_${c}`;
            next[key] = {
              ...(next[key] || {}),
              ...formatUpdate,
            };
          }
        }
      } else if (selectedCell) {
        const key = `${selectedCell.rowIndex}_${selectedCell.colIndex}`;
        next[key] = {
          ...(next[key] || {}),
          ...formatUpdate,
        };
      } else if (selectedRowIndex >= 0) {
        [-1, 0, 1, 2, 3, 4].forEach((c) => {
          const key = `${selectedRowIndex}_${c}`;
          next[key] = {
            ...(next[key] || {}),
            ...formatUpdate,
          };
        });
      }

      return next;
    });
  };

  const handleApplyBorderPreset = (preset: string) => {
    setCellFormats((prev) => {
      const next = { ...prev };
      const minRow = selectionStart && selectionEnd ? Math.min(selectionStart.rowIndex, selectionEnd.rowIndex) : (selectedCell ? selectedCell.rowIndex : selectedRowIndex);
      const maxRow = selectionStart && selectionEnd ? Math.max(selectionStart.rowIndex, selectionEnd.rowIndex) : (selectedCell ? selectedCell.rowIndex : selectedRowIndex);
      const minCol = selectionStart && selectionEnd ? Math.min(selectionStart.colIndex, selectionEnd.colIndex) : (selectedCell ? selectedCell.colIndex : -1);
      const maxCol = selectionStart && selectionEnd ? Math.max(selectionStart.colIndex, selectionEnd.colIndex) : (selectedCell ? selectedCell.colIndex : 4);

      for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
          const key = `${r}_${c}`;
          const existing = next[key] || {};
          let borders: CellBorders = { ...(existing.borders || {}) };

          const isTop = r === minRow;
          const isBottom = r === maxRow;
          const isLeft = c === minCol;
          const isRight = c === maxCol;

          switch (preset) {
            case "all":
              borders = {
                top: "1px solid black",
                bottom: "1px solid black",
                left: "1px solid black",
                right: "1px solid black",
              };
              break;
            case "none":
              borders = {
                top: "none",
                bottom: "none",
                left: "none",
                right: "none",
              };
              break;
            case "outside":
              if (isTop) borders.top = "1px solid black";
              if (isBottom) borders.bottom = "1px solid black";
              if (isLeft) borders.left = "1px solid black";
              if (isRight) borders.right = "1px solid black";
              break;
            case "thick_outside":
              if (isTop) borders.top = "2px solid black";
              if (isBottom) borders.bottom = "2px solid black";
              if (isLeft) borders.left = "2px solid black";
              if (isRight) borders.right = "2px solid black";
              break;
            case "bottom":
              if (isBottom) borders.bottom = "1px solid black";
              break;
            case "top":
              if (isTop) borders.top = "1px solid black";
              break;
            case "left":
              if (isLeft) borders.left = "1px solid black";
              break;
            case "right":
              if (isRight) borders.right = "1px solid black";
              break;
            case "thick_bottom":
              if (isBottom) borders.bottom = "2px solid black";
              break;
            case "bottom_double":
              if (isBottom) borders.bottom = "3px double black";
              break;
            case "top_and_bottom":
              if (isTop) borders.top = "1px solid black";
              if (isBottom) borders.bottom = "1px solid black";
              break;
            case "top_and_thick_bottom":
              if (isTop) borders.top = "1px solid black";
              if (isBottom) borders.bottom = "2px solid black";
              break;
            case "top_and_double_bottom":
              if (isTop) borders.top = "1px solid black";
              if (isBottom) borders.bottom = "3px double black";
              break;
          }

          next[key] = {
            ...existing,
            borders,
          };
        }
      }

      return next;
    });
  };

  const getCellStyle = (rowIndex: number, colIndex: number): React.CSSProperties => {
    const key = `${rowIndex}_${colIndex}`;
    const fmt = cellFormats[key];
    if (!fmt) return {};

    const style: React.CSSProperties = {};

    if (fmt.fontFamily) style.fontFamily = fmt.fontFamily;
    if (fmt.fontSize) style.fontSize = `${fmt.fontSize}pt`;
    if (fmt.bold !== undefined) style.fontWeight = fmt.bold ? "bold" : "normal";
    if (fmt.italic !== undefined) style.fontStyle = fmt.italic ? "italic" : "normal";
    if (fmt.underline) {
      if (fmt.underline === "double") {
        style.textDecoration = "underline";
        style.textDecorationStyle = "double";
      } else if (fmt.underline === "single") {
        style.textDecoration = "underline";
      } else {
        style.textDecoration = "none";
      }
    }
    if (fmt.align) style.textAlign = fmt.align;
    if (fmt.valign) style.verticalAlign = fmt.valign;
    if (fmt.color) style.color = fmt.color;
    if (fmt.bgColor) style.backgroundColor = fmt.bgColor;
    if (fmt.indent) {
      style.paddingLeft = `${fmt.indent * 8}px`;
      style.textIndent = `${fmt.indent * 8}px`;
    }
    if (fmt.orientation && fmt.orientation !== "horizontal") {
      switch (fmt.orientation) {
        case "angle-up":
          style.transform = "rotate(-45deg)";
          style.display = "inline-block";
          break;
        case "angle-down":
          style.transform = "rotate(45deg)";
          style.display = "inline-block";
          break;
        case "vertical":
          style.writingMode = "vertical-rl";
          break;
        case "rotate-up":
          style.transform = "rotate(-90deg)";
          style.display = "inline-block";
          break;
        case "rotate-down":
          style.transform = "rotate(90deg)";
          style.display = "inline-block";
          break;
      }
    }

    if (fmt.borders) {
      if (fmt.borders.top) style.borderTop = fmt.borders.top;
      if (fmt.borders.bottom) style.borderBottom = fmt.borders.bottom;
      if (fmt.borders.left) style.borderLeft = fmt.borders.left;
      if (fmt.borders.right) style.borderRight = fmt.borders.right;
    }

    return style;
  };

  useEffect(() => {
    const handleGlobalShortcuts = (e: KeyboardEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;

      const target = e.target as HTMLElement;
      const isSheetInput = target?.closest?.(".sheet") || target?.hasAttribute?.("data-row");

      if (!selectedCell && !selectionStart && !isSheetInput) return;

      const key = e.key.toLowerCase();

      if (key === "b") {
        e.preventDefault();
        handleApplyFormat({ bold: !activeCellFormat.bold });
      } else if (key === "i") {
        e.preventDefault();
        handleApplyFormat({ italic: !activeCellFormat.italic });
      } else if (key === "u") {
        e.preventDefault();
        handleApplyFormat({ underline: activeCellFormat.underline === "single" ? "none" : "single" });
      } else if (key === "l" && !e.shiftKey) {
        e.preventDefault();
        handleApplyFormat({ align: "left" });
      } else if (key === "e" && !e.shiftKey) {
        e.preventDefault();
        handleApplyFormat({ align: "center" });
      } else if (key === "r" && !e.shiftKey) {
        e.preventDefault();
        handleApplyFormat({ align: "right" });
      } else if (e.shiftKey && (key === ">" || key === ".")) {
        e.preventDefault();
        const currentSize = activeCellFormat.fontSize || 8.5;
        handleApplyFormat({ fontSize: Math.min(72, currentSize + 1) });
      } else if (e.shiftKey && (key === "<" || key === ",")) {
        e.preventDefault();
        const currentSize = activeCellFormat.fontSize || 8.5;
        handleApplyFormat({ fontSize: Math.max(5, currentSize - 1) });
      }
    };

    window.addEventListener("keydown", handleGlobalShortcuts);
    return () => window.removeEventListener("keydown", handleGlobalShortcuts);
  }, [activeCellFormat, selectedCell, selectionStart, selectionEnd, selectedRowIndex]);

  const handleCellClick = (rowIndex: number, colIndex: number) => {
    setSelectedRowIndex(rowIndex);
    setSelectedCell({ rowIndex, colIndex });
    setSelectionStart({ rowIndex, colIndex });
    setSelectionEnd({ rowIndex, colIndex });
    const textarea = document.querySelector(`[data-row="${rowIndex}"][data-col="${colIndex}"]`) as HTMLTextAreaElement | null;
    if (textarea) {
      textarea.focus();
    }
  };

  const clearSpecificCell = (rowIndex: number, colIndex: number) => {
    let field: "desc" | "qty" | "unit" | "price" | null = null;
    if (colIndex === 0) field = "desc";
    else if (colIndex === 1) field = "qty";
    else if (colIndex === 2) field = "unit";
    else if (colIndex === 3) field = "price";
    
    if (field) {
      handleRowChange(rowIndex, field, "");
    }
  };

  const handleCellContextMenu = (e: React.MouseEvent, idx: number, colIdx: number) => {
    e.preventDefault();
    const clickedInsideRange = isCellSelected(idx, colIdx);
    if (!clickedInsideRange) {
      setSelectionStart({ rowIndex: idx, colIndex: colIdx });
      setSelectionEnd({ rowIndex: idx, colIndex: colIdx });
    }
    setSelectedRowIndex(idx);
    setSelectedCell({ rowIndex: idx, colIndex: colIdx });
    
    let x = e.clientX;
    let y = e.clientY;
    const menuWidth = 260;
    const menuHeight = 320;
    
    if (x + menuWidth > window.innerWidth) x = window.innerWidth - menuWidth - 10;
    if (y + menuHeight > window.innerHeight) y = window.innerHeight - menuHeight - 10;
    if (x < 0) x = 10;
    if (y < 0) y = 10;

    setContextMenu({ visible: true, x, y, rowIndex: idx, colIndex: colIdx });
  };

  const rowsTotal = rows.reduce((sum, r) => sum + r.amount, 0);
  const parsedVatPercent = parseNumericInput(vatPercent);
  const parsedTransportationFee = parseNumericInput(transportationFee);
  const vatAmount = docType === "invoice" ? (rowsTotal * parsedVatPercent) / 100 : 0;
  const grandTotal = docType === "invoice" ? (rowsTotal + vatAmount + parsedTransportationFee) : rowsTotal;
  const calculatedGrandTotal = docType === "challan" ? 0 : grandTotal;

  const handleKeyDown = (e: React.KeyboardEvent<HTMLElement>, rowIndex: number, colIndex: number) => {
    const { key } = e;
    let targetRow = rowIndex;
    let targetCol = colIndex;

    const targetEl = e.currentTarget;
    const isFormInput = targetEl instanceof HTMLInputElement || targetEl instanceof HTMLTextAreaElement;

    let cursorStart = 0;
    let cursorEnd = 0;
    let valueLength = 0;

    if (isFormInput) {
      cursorStart = targetEl.selectionStart ?? 0;
      cursorEnd = targetEl.selectionEnd ?? 0;
      valueLength = targetEl.value?.length ?? 0;
    } else {
      const sel = typeof window !== "undefined" ? window.getSelection() : null;
      if (sel && sel.rangeCount > 0) {
        const range = sel.getRangeAt(0);
        cursorStart = range.startOffset;
        cursorEnd = range.endOffset;
      }
      valueLength = (targetEl.textContent || "").length;
    }

    if (key === "ArrowUp") {
      targetRow = rowIndex - 1;
    } else if (key === "ArrowDown") {
      targetRow = rowIndex + 1;
    } else if (key === "ArrowLeft") {
      if (cursorStart === 0 && cursorEnd === 0) {
        targetCol = colIndex - 1;
      } else {
        return;
      }
    } else if (key === "ArrowRight") {
      if (cursorStart === valueLength && cursorEnd === valueLength) {
        targetCol = colIndex + 1;
      } else {
        return;
      }
    } else if (key === "Enter") {
      if (e.shiftKey) {
        return;
      }
      targetRow = rowIndex + 1;
    } else {
      return;
    }

    e.preventDefault();
    const targetElement = document.querySelector(
      `[data-row="${targetRow}"][data-col="${targetCol}"]`
    ) as HTMLElement | null;

    if (targetElement) {
      targetElement.focus();
      if (typeof (targetElement as any).select === "function") {
        (targetElement as any).select();
      }
    }
  };

  const handlePaste = (
    e: React.ClipboardEvent<HTMLElement>,
    startRowIndex: number,
    startColIndex: number
  ) => {
    const plainText = e.clipboardData.getData("text/plain") || e.clipboardData.getData("text");
    const htmlText = e.clipboardData.getData("text/html");

    if (!plainText && !htmlText) return;

    const result = parseClipboardData({ text: plainText, html: htmlText });
    const parsedGrid = result.grid;

    if (!parsedGrid || parsedGrid.length === 0) return;

    // 1. Single plain cell without internal newlines -> allow default inline insertion
    if (parsedGrid.length === 1 && parsedGrid[0].length === 1 && !parsedGrid[0][0].includes("\n")) {
      const targetEl = e.currentTarget;
      if (targetEl instanceof HTMLInputElement || targetEl instanceof HTMLTextAreaElement) {
        const parsedVal = cleanCellText(parsedGrid[0][0]);
        const start = targetEl.selectionStart ?? 0;
        const end = targetEl.selectionEnd ?? 0;
        const currentValue = targetEl.value || "";
        const newValue = currentValue.substring(0, start) + parsedVal + currentValue.substring(end);

        const fieldMap = ["desc", "qty", "unit", "price"] as const;
        const field = fieldMap[startColIndex];
        if (field) {
          e.preventDefault();
          handleRowChange(startRowIndex, field, newValue);
          setTimeout(() => {
            targetEl.focus();
            targetEl.selectionStart = targetEl.selectionEnd = start + parsedVal.length;
          }, 0);
        }
      }
      return;
    }

    // 2. Multi-cell, multi-line, or multi-column paste from Excel / Google Sheets
    e.preventDefault();

    const dataRows = result.hasHeader ? parsedGrid.slice(1) : parsedGrid;
    if (dataRows.length === 0) return;

    const hasSerialCol = result.hasSerialColumn;
    const colCount = Math.max(...dataRows.map(r => r.length));

    setRows((prevRows) => {
      const updated = [...prevRows];

      dataRows.forEach((cols, rOffset) => {
        const rIndex = startRowIndex + rOffset;
        if (rIndex >= MAX_LINES_LIMIT) return;

        if (rIndex >= updated.length) {
          updated.push({
            sl: updated.length + 1,
            desc: "",
            qty: "",
            unit: "",
            price: "",
            amount: 0,
          });
        }

        const targetRow = { ...updated[rIndex] };

        // Intelligent column routing with clean continuous text
        if (startColIndex === -1) {
          // Clicked in SL column
          if (hasSerialCol || colCount >= 4) {
            if (cols[1] !== undefined) targetRow.desc = cleanCellText(cols[1]);
            if (cols[2] !== undefined) targetRow.qty = cleanCellText(cols[2]);
            if (cols[3] !== undefined) targetRow.unit = cleanCellText(cols[3]);
            if (cols[4] !== undefined && docType !== "challan") targetRow.price = cleanCellText(cols[4]);
          } else {
            if (cols[0] !== undefined) targetRow.desc = cleanCellText(cols[0]);
            if (cols[1] !== undefined) targetRow.qty = cleanCellText(cols[1]);
            if (cols[2] !== undefined) targetRow.unit = cleanCellText(cols[2]);
            if (cols[3] !== undefined && docType !== "challan") targetRow.price = cleanCellText(cols[3]);
          }
        } else if (startColIndex === 0) {
          // Clicked in Description column (Col 0)
          if (hasSerialCol && colCount >= 4) {
            if (cols[1] !== undefined) targetRow.desc = cleanCellText(cols[1]);
            if (cols[2] !== undefined) targetRow.qty = cleanCellText(cols[2]);
            if (cols[3] !== undefined) targetRow.unit = cleanCellText(cols[3]);
            if (cols[4] !== undefined && docType !== "challan") targetRow.price = cleanCellText(cols[4]);
          } else {
            cols.forEach((cellValue, cOffset) => {
              const cIndex = startColIndex + cOffset;
              if (cIndex === 0) targetRow.desc = cleanCellText(cellValue);
              else if (cIndex === 1) targetRow.qty = cleanCellText(cellValue);
              else if (cIndex === 2) targetRow.unit = cleanCellText(cellValue);
              else if (cIndex === 3 && docType !== "challan") targetRow.price = cleanCellText(cellValue);
            });
          }
        } else {
          // Pasting into specific sub-column (Qty, Unit, or Price)
          cols.forEach((cellValue, cOffset) => {
            const cIndex = startColIndex + cOffset;
            if (cIndex === 0) targetRow.desc = cleanCellText(cellValue);
            else if (cIndex === 1) targetRow.qty = cleanCellText(cellValue);
            else if (cIndex === 2) targetRow.unit = cleanCellText(cellValue);
            else if (cIndex === 3 && docType !== "challan") targetRow.price = cleanCellText(cellValue);
          });
        }

        const cleanQty = stripHtml(String(targetRow.qty || ""));
        const cleanPrice = stripHtml(String(targetRow.price || ""));
        const q = parseNumericInput(cleanQty);
        const p = parseNumericInput(cleanPrice);
        targetRow.amount = docType === "challan" ? 0 : q * p;
        updated[rIndex] = targetRow;
      });

      showToast(`Distributed ${dataRows.length} lines from Excel`);
      return updated;
    });
  };

  const unwrapAllDescriptions = () => {
    let fixedCount = 0;
    setRows((prev) => {
      const updated = prev.map((r) => {
        const original = r.desc || "";
        const cleaned = cleanCellText(original);
        if (cleaned !== original) {
          fixedCount++;
        }
        return {
          ...r,
          desc: cleaned,
          qty: cleanCellText(r.qty || ""),
          unit: cleanCellText(r.unit || ""),
          price: cleanCellText(r.price || ""),
        };
      });
      return updated;
    });

    if (fixedCount > 0) {
      showToast(`Cleaned & unwrapped line breaks across ${fixedCount} item(s)`);
    } else {
      showToast(`All description lines are continuous and clean`);
    }
  };

  const handleDownloadExcel = async () => {
    try {
      setIsGeneratingExcel(true);
      const workbook = await generateExcelWorkbook(
        docType,
        messers,
        address,
        challanNo,
        dateVal,
        requisitionNo,
        rows,
        mergedRegions,
        invoiceNo,
        poNumber,
        parseNumericInput(vatPercent),
        parseNumericInput(transportationFee),
        cellFormats
      );
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

      const filePrefix = docType === "challan" ? "Challan" : docType === "invoice" ? "Invoice" : "Quotation";
      const identifier = docType === "challan" ? (challanNo || "NEW") : docType === "invoice" ? (invoiceNo || "NEW") : (requisitionNo || "NEW");
      const defaultFileName = `${filePrefix}_${identifier.replace(/[\/\\?%*:|"<>\s]/g, "_")}.xlsx`;

      // 1. Try modern File System Access API first (highly supported on Desktop browsers like Chrome, Edge, Opera)
      // This allows selecting directory, browsing existing files, renaming, or choosing paths dynamically.
      if ('showSaveFilePicker' in window) {
        try {
          const handle = await (window as any).showSaveFilePicker({
            suggestedName: defaultFileName,
            types: [{
              description: 'Excel Spreadsheet',
              accept: {
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
              }
            }]
          });
          const writable = await handle.createWritable();
          await writable.write(blob);
          await writable.close();
          return; // Done successfully
        } catch (err: any) {
          if (err.name === 'AbortError') {
            // User cancelled the native save picker - abort nicely without showing error/fallback
            return;
          }
          console.warn("showSaveFilePicker failed or was blocked, falling back to prompt method:", err);
        }
      }

      // 2. Fallback: Prompt the user to customize the filename, then run standard Anchor download
      const userFileName = prompt("Enter a filename to save:", defaultFileName);
      if (userFileName === null) {
        // User clicked Cancel
        return;
      }

      const finalFileName = userFileName.trim()
        ? (userFileName.toLowerCase().endsWith(".xlsx") ? userFileName.trim() : `${userFileName.trim()}.xlsx`)
        : defaultFileName;

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = finalFileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: any) {
      console.error("Excel download error:", err);
      alert("Error generating Excel: " + err.message);
    } finally {
      setIsGeneratingExcel(false);
    }
  };

  const handlePrint = () => {
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }
    setSelectedCell(null);
    setSelectionStart(null);
    setSelectionEnd(null);
    setTimeout(() => {
      window.print();
    }, 50);
  };

  const handleDownloadPDF = () => {
    const container = document.querySelector(".quotation-container");
    if (!container) return;
    
    const element = document.querySelector(".sheet") as HTMLElement | null;
    if (!element) return;
    
    setIsGeneratingPDF(true);
    container.classList.add("is-generating-pdf");
    document.body.classList.add("is-generating-pdf");

    const originalGetComputedStyle = window.getComputedStyle;
    window.getComputedStyle = function (el: Element, pseudoElt?: string | null) {
      const style = originalGetComputedStyle.call(this, el, pseudoElt);
      return new Proxy(style, {
        get(target, prop, receiver) {
          if (prop === 'getPropertyValue') {
            return function(propertyName: string) {
              const val = target.getPropertyValue(propertyName);
              return typeof val === 'string' ? replaceOklchInCss(val) : val;
            };
          }
          const val = Reflect.get(target, prop, receiver);
          if (typeof val === 'string') {
            return replaceOklchInCss(val);
          }
          if (typeof val === 'function') {
            return val.bind(target);
          }
          return val;
        }
      }) as any;
    };
    
    const sheets = Array.from(document.styleSheets);
    let concatenatedCss = "";
    const disabledSheets: { sheet: CSSStyleSheet; wasDisabled: boolean }[] = [];

    for (const sheet of sheets) {
      try {
        const rules = Array.from(sheet.cssRules || []);
        const sheetCss = rules.map(rule => rule.cssText).join("\n");
        concatenatedCss += sheetCss + "\n";
        
        disabledSheets.push({ sheet, wasDisabled: sheet.disabled });
        sheet.disabled = true;
      } catch (e) {
        console.warn("Could not read stylesheet rules (possibly cross-origin):", e);
      }
    }

    const translatedCss = replaceOklchInCss(concatenatedCss);
    const tempStyle = document.createElement("style");
    tempStyle.id = "temp-pdf-colors";
    tempStyle.textContent = translatedCss;
    document.head.appendChild(tempStyle);

    const elementsWithInlineStyle = element.querySelectorAll("[style]");
    const inlineStylesBackup = new Map<HTMLElement, string>();
    
    const rootStyle = element.getAttribute("style");
    if (rootStyle && rootStyle.includes("oklch")) {
      inlineStylesBackup.set(element, rootStyle);
      element.setAttribute("style", replaceOklchInCss(rootStyle));
    }
    
    elementsWithInlineStyle.forEach((el) => {
      const htmlEl = el as HTMLElement;
      const styleAttr = htmlEl.getAttribute("style");
      if (styleAttr && styleAttr.includes("oklch")) {
        inlineStylesBackup.set(htmlEl, styleAttr);
        htmlEl.setAttribute("style", replaceOklchInCss(styleAttr));
      }
    });

    const filePrefix = docType === "challan" ? "Challan" : docType === "invoice" ? "Invoice" : "Quotation";
    const identifier = docType === "challan" ? (challanNo || "NEW") : docType === "invoice" ? (invoiceNo || "NEW") : (requisitionNo || "NEW");
    const filename = `${filePrefix}_${identifier.replace(/[\/\\?%*:|"<>\s]/g, "_")}.pdf`;
    
    const opt = {
      margin:       10,
      filename:     filename,
      image:        { type: "jpeg" as const, quality: 0.98 },
      html2canvas:  { 
        scale: 2.5,
        useCORS: true,
        logging: false,
        scrollY: 0,
        scrollX: 0
      },
      jsPDF:        { unit: "mm", format: "a4", orientation: "portrait" as const }
    };
    
    const cleanUpAfterPdf = () => {
      window.getComputedStyle = originalGetComputedStyle;

      disabledSheets.forEach(({ sheet, wasDisabled }) => {
        sheet.disabled = wasDisabled;
      });
      
      const addedStyle = document.getElementById("temp-pdf-colors");
      if (addedStyle) {
        addedStyle.remove();
      }
      
      inlineStylesBackup.forEach((originalStyle, htmlEl) => {
        htmlEl.setAttribute("style", originalStyle);
      });
      
      container.classList.remove("is-generating-pdf");
      document.body.classList.remove("is-generating-pdf");
      setIsGeneratingPDF(false);
    };

    // @ts-ignore
    html2pdf()
      .from(element)
      .set(opt)
      .save()
      .then(() => {
        cleanUpAfterPdf();
      })
      .catch((err: any) => {
        console.error("PDF generation error:", err);
        cleanUpAfterPdf();
      });
  };

  const safeSelectedRowIndex = Math.max(0, Math.min(selectedRowIndex, rows.length - 1));
  const GRID_COLUMNS = docType === "challan" ? [-1, 0, 1, 2] : [-1, 0, 1, 2, 3, 4];

  return (
    <div className="quotation-container relative min-h-screen flex flex-col items-center bg-slate-50 py-3 sm:py-4 text-[#000] font-sans antialiased w-full">
      
      {/* Consolidated Top Toolbar Table - Sticky at Top of Whole Page */}
      <div className="sticky top-0 z-40 w-full max-w-[210mm] px-2 sm:px-0 no-print print:hidden mb-2">
        <ExcelRibbonToolbar
          activeFormat={activeCellFormat}
          onApplyFormat={handleApplyFormat}
          onApplyBorderPreset={handleApplyBorderPreset}
          selectionSummary={selectionSummary}
          canMerge={!!(selectionStart && selectionEnd && (selectionStart.rowIndex !== selectionEnd.rowIndex || selectionStart.colIndex !== selectionEnd.colIndex))}
          onToggleMerge={toggleMergeSelectedRange}
          onClearFormatting={() => {
            if (selectionStart && selectionEnd) {
              const minRow = Math.min(selectionStart.rowIndex, selectionEnd.rowIndex);
              const maxRow = Math.max(selectionStart.rowIndex, selectionEnd.rowIndex);
              const minCol = Math.min(selectionStart.colIndex, selectionEnd.colIndex);
              const maxCol = Math.max(selectionStart.colIndex, selectionEnd.colIndex);
              setCellFormats((prev) => {
                const next = { ...prev };
                for (let r = minRow; r <= maxRow; r++) {
                  for (let c = minCol; c <= maxCol; c++) {
                    delete next[`${r}_${c}`];
                  }
                }
                return next;
              });
            } else if (selectedCell) {
              setCellFormats((prev) => {
                const next = { ...prev };
                delete next[`${selectedCell.rowIndex}_${selectedCell.colIndex}`];
                return next;
              });
            } else if (selectedRowIndex >= 0) {
              setCellFormats((prev) => {
                const next = { ...prev };
                [-1, 0, 1, 2, 3, 4].forEach((c) => {
                  delete next[`${selectedRowIndex}_${c}`];
                });
                return next;
              });
            }
          }}
          docType={docType}
          onSelectDocType={(type) => {
            setDocType(type);
            setMergedRegions([]);
            setRows((prev) =>
              prev.map((r) => {
                const q = parseNumericInput(stripHtml(String(r.qty || "")));
                const p = parseNumericInput(stripHtml(String(r.price || "")));
                return {
                  ...r,
                  amount: type === "challan" ? 0 : q * p,
                };
              })
            );
          }}
          autoSaveEnabled={autoSaveEnabled}
          onToggleAutoSave={(val) => {
            setAutoSaveEnabled(val);
            localStorage.setItem("comilla_autosave_enabled", String(val));
          }}
          lastSavedTime={lastSavedTime}
          currentDocId={currentDocId}
          currentDocName={savedDocs.find((d) => d.id === currentDocId)?.name}
          onCloseCurrentDoc={resetSheetFields}
          onNewDoc={startNewDoc}
          onDuplicateDoc={currentDocId ? duplicateCurrentDoc : undefined}
          onDeleteDoc={currentDocId ? () => deleteSavedDoc(currentDocId) : undefined}
          onSaveDoc={() => saveCurrentDocToApp()}
          saveStatus={saveStatus}
          onOpenExcelModal={() => setIsExcelModalOpen(true)}
          onExportExcel={handleDownloadExcel}
          isGeneratingExcel={isGeneratingExcel}
          onPrint={handlePrint}
        />
      </div>

      {/* A4 Standard-compliant visual grid container */}
      <div className="sheet relative w-full max-w-[210mm] min-h-[297mm] bg-white p-3 sm:p-[8mm] print:p-0 shadow-xl border border-slate-200/60 rounded-xs box-border z-10 mx-auto">
        
        {/* Anti-slip Background Watermark Asset */}
        <div className="watermark-container absolute inset-0 pointer-events-none flex items-center justify-center overflow-hidden z-0 select-none">
          <img 
            src="https://i.ibb.co.com/3mNycQXx/1.png" 
            alt="Watermark background" 
            referrerPolicy="no-referrer"
            className="w-[70%] opacity-[0.045] object-contain select-none max-w-[500px]"
            style={{ printColorAdjust: "exact" }}
          />
        </div>

        {/* Outer Layout Table ensuring thead repeats company details on multi-page browser printing */}
        <table className="print-outer-layout-table w-full border-none p-0 m-0 relative z-10">
          <thead className="print:table-header-group">
            <tr>
              <td className="border-none p-0 m-0">
                {/* Top blank margin repeating on every printed page */}
                <div className="print-page-top-spacer hidden print:block h-[12mm] w-full" />
                
                <div className="business-header border-b-2 border-black pb-1.5 mb-1.5 flex flex-col sm:flex-row items-center justify-between gap-2.5 text-black text-left">
                  <div className="flex items-center gap-3">
                    <div className="logo-container h-16 w-16 sm:h-20 sm:w-20 shrink-0 rounded-full border border-slate-300 overflow-hidden bg-black flex items-center justify-center shadow-sm">
                      <img
                        src="https://i.ibb.co.com/gFBkpt8B/Chat-GPT-Image-Apr-23-2026-01-10-13-PM.png"
                        alt="Comilla Traders Logo"
                        className="w-full h-full object-cover"
                      />
                    </div>
                    <div>
                      <h1 className="text-[17pt] sm:text-[19pt] font-black tracking-tight leading-none text-black">
                        COMILLA TRADERS
                      </h1>
                      <p className="text-[8pt] font-extrabold text-slate-700 tracking-wider uppercase mt-1">
                        Ship Chandler, Marine Supplier & General Merchant
                      </p>
                      <p className="text-[7pt] font-bold text-slate-500 uppercase tracking-widest mt-0.5">
                        Mechanical & Electrical Marine Engineering Services
                      </p>
                    </div>
                  </div>

                  <div className="contact-details text-right text-[7.5pt] text-slate-800 space-y-0.5 leading-tight sm:block hidden print:block">
                    <p className="font-bold whitespace-nowrap">
                      Office: <span className="font-medium whitespace-nowrap">Jubilee Road, Chattogram, Bangladesh</span>
                    </p>
                    <p className="font-bold whitespace-nowrap">
                      Helplines: <span className="font-medium font-mono whitespace-nowrap">01819315746, 01712-900431</span>
                    </p>
                    <p className="font-bold whitespace-nowrap">
                      Official Email: <span className="font-medium whitespace-nowrap">comillatraders@gmail.com</span>
                    </p>
                    <p className="font-bold text-[7pt] tracking-widest text-indigo-700 uppercase whitespace-nowrap">
                      CHATTOGRAM &bull; BANGLADESH
                    </p>
                  </div>
                  
                  {/* Print contact information layout */}
                  <div className="text-center text-[8pt] text-slate-800 space-y-0.5 leading-tight sm:hidden print:hidden">
                    <p>Jubilee Road, Chattogram &bull; Hotlines: 01819315746</p>
                    <p>comillatraders@gmail.com</p>
                  </div>
                </div>

                {/* Repeating Document Title on multi-page browser printing */}
                <div className="doc-title text-center text-[12pt] sm:text-[13pt] font-black uppercase tracking-[8px] my-1">
                  {docType === "challan" ? "Delivery Challan" : docType === "invoice" ? "Bill / Invoice" : "Quotation"}
                </div>

                {/* Repeating Metadata Information Input Grid on multi-page browser printing */}
                <div className="meta-grid grid grid-cols-1 sm:grid-cols-2 gap-2 text-left text-[8.5pt] mb-1.5">
                  <div className="meta-box space-y-1 border border-black p-2 bg-slate-50/30 rounded-xs">
                    <div>
                      <label className="block text-[7pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Messers:</label>
                      <RichTextCell
                        value={messers}
                        onChange={(val) => setMessers(val)}
                        placeholder="Enter Client/Ship details"
                        className="w-full border-b border-dotted border-slate-400 focus:border-black font-bold text-[9pt] outline-none bg-transparent py-0.5 no-print print:hidden min-h-[22px]"
                      />
                      <div 
                        className="hidden print:block font-bold text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5 break-words whitespace-pre-wrap leading-tight"
                        dangerouslySetInnerHTML={{ __html: messers || "&nbsp;" }}
                      />
                    </div>
                    <div>
                      <label className="block text-[7pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Address:</label>
                      <RichTextCell
                        value={address}
                        onChange={(val) => setAddress(val)}
                        placeholder="Enter delivery/billing address"
                        className="w-full border-b border-dotted border-slate-400 focus:border-black text-[8.5pt] outline-none bg-transparent leading-tight py-0.5 no-print print:hidden min-h-[36px]"
                      />
                      <div 
                        className="hidden print:block text-[8.5pt] border-b border-dotted border-black min-h-[32px] py-0.5 break-words whitespace-pre-wrap leading-tight"
                        dangerouslySetInnerHTML={{ __html: address || "&nbsp;" }}
                      />
                    </div>
                  </div>

                  <div className={`meta-box grid border border-black p-2 bg-slate-50/30 rounded-xs ${
                    docType === "invoice" ? "grid-cols-3 gap-1.5" : "grid-cols-2 gap-1.5"
                  }`}>
                    {docType === "invoice" ? (
                      <>
                        <div className="meta-inner-field col-span-1">
                          <label className="block text-[7.5pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Invoice No.:</label>
                          <input 
                            type="text" 
                            value={invoiceNo}
                            onChange={(e) => setInvoiceNo(e.target.value)}
                            placeholder="Invoice number"
                            className="w-full border-b border-dotted border-slate-400 focus:border-black font-mono text-[9pt] outline-none bg-transparent py-0.5 no-print print:hidden"
                          />
                          <div className="hidden print:block font-mono text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5 break-words">
                            {invoiceNo || " "}
                          </div>
                        </div>
                        <div className="meta-inner-field col-span-1">
                          <label className="block text-[7.5pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Challan No.:</label>
                          <input 
                            type="text" 
                            value={challanNo}
                            onChange={(e) => setChallanNo(e.target.value)}
                            placeholder="Challan number"
                            className="w-full border-b border-dotted border-slate-400 focus:border-black font-mono text-[9pt] outline-none bg-transparent py-0.5 no-print print:hidden"
                          />
                          <div className="hidden print:block font-mono text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5 break-words">
                            {challanNo || " "}
                          </div>
                        </div>
                        <div className="meta-inner-field col-span-1 relative">
                          <label className="block text-[7.5pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Date:</label>
                          <div className="flex items-center gap-1 no-print print:hidden">
                            <input 
                              type="text" 
                              value={dateVal}
                              onChange={(e) => setDateVal(e.target.value)}
                              className="w-full border-b border-dotted border-slate-400 focus:border-black font-mono text-[9pt] outline-none bg-transparent py-0.5"
                            />
                            <button
                              type="button"
                              onClick={triggerDatePicker}
                              className="p-0.5 hover:bg-slate-100 rounded text-slate-600 transition-colors cursor-pointer flex items-center justify-center shrink-0"
                            >
                              <Calendar className="h-3.5 w-3.5" />
                            </button>
                          </div>
                          <div className="hidden print:block font-mono text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5">
                            {dateVal || " "}
                          </div>
                          <input
                            ref={dateRef}
                            type="date"
                            onChange={handleDatePickerChange}
                            className="absolute invisible w-0 h-0 opacity-0 pointer-events-none"
                          />
                        </div>
                        <div className="meta-inner-field col-span-1">
                          <label className="block text-[7.5pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Requisition No.:</label>
                          <input 
                            type="text" 
                            value={requisitionNo}
                            onChange={(e) => setRequisitionNo(e.target.value)}
                            placeholder="Requisition number"
                            className="w-full border-b border-dotted border-slate-400 focus:border-black font-mono text-[9pt] outline-none bg-transparent py-0.5 no-print print:hidden"
                          />
                          <div className="hidden print:block font-mono text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5 break-words">
                            {requisitionNo || " "}
                          </div>
                        </div>
                        <div className="meta-inner-field col-span-2">
                          <label className="block text-[7.5pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">PO Number:</label>
                          <input 
                            type="text" 
                            value={poNumber}
                            onChange={(e) => setPoNumber(e.target.value)}
                            placeholder="PO number"
                            className="w-full border-b border-dotted border-slate-400 focus:border-black font-mono text-[9pt] outline-none bg-transparent py-0.5 no-print print:hidden"
                          />
                          <div className="hidden print:block font-mono text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5 break-words">
                            {poNumber || " "}
                          </div>
                        </div>
                      </>
                    ) : docType === "challan" ? (
                      <>
                        <div className="meta-inner-field col-span-1">
                          <label className="block text-[7.5pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Challan No.:</label>
                          <input 
                            type="text" 
                            value={challanNo}
                            onChange={(e) => setChallanNo(e.target.value)}
                            placeholder="Challan number"
                            className="w-full border-b border-dotted border-slate-400 focus:border-black font-mono text-[9pt] outline-none bg-transparent py-0.5 no-print print:hidden"
                          />
                          <div className="hidden print:block font-mono text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5 break-words">
                            {challanNo || " "}
                          </div>
                        </div>
                        <div className="meta-inner-field col-span-1 relative">
                          <label className="block text-[7.5pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Date:</label>
                          <div className="flex items-center gap-1 no-print print:hidden">
                            <input 
                              type="text" 
                              value={dateVal}
                              onChange={(e) => setDateVal(e.target.value)}
                              className="w-full border-b border-dotted border-slate-400 focus:border-black font-mono text-[9pt] outline-none bg-transparent py-0.5"
                            />
                            <button
                              type="button"
                              onClick={triggerDatePicker}
                              className="p-0.5 hover:bg-slate-100 rounded text-slate-600 transition-colors cursor-pointer flex items-center justify-center shrink-0"
                            >
                              <Calendar className="h-3.5 w-3.5" />
                            </button>
                          </div>
                          <div className="hidden print:block font-mono text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5">
                            {dateVal || " "}
                          </div>
                          <input
                            ref={dateRef}
                            type="date"
                            onChange={handleDatePickerChange}
                            className="absolute invisible w-0 h-0 opacity-0 pointer-events-none"
                          />
                        </div>
                      </>
                    ) : (
                      <div className="meta-inner-field col-span-2 relative">
                        <label className="block text-[7.5pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Date:</label>
                        <div className="flex items-center gap-1 no-print print:hidden">
                          <input 
                            type="text" 
                            value={dateVal}
                            onChange={(e) => setDateVal(e.target.value)}
                            className="w-full border-b border-dotted border-slate-400 focus:border-black font-mono text-[9pt] outline-none bg-transparent py-0.5"
                          />
                          <button
                            type="button"
                            onClick={triggerDatePicker}
                            className="p-0.5 hover:bg-slate-100 rounded text-slate-600 transition-colors cursor-pointer flex items-center justify-center shrink-0"
                          >
                            <Calendar className="h-3.5 w-3.5" />
                          </button>
                        </div>
                        <div className="hidden print:block font-mono text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5">
                          {dateVal || " "}
                        </div>
                        <input
                          ref={dateRef}
                          type="date"
                          onChange={handleDatePickerChange}
                          className="absolute invisible w-0 h-0 opacity-0 pointer-events-none"
                        />
                      </div>
                    )}
                    
                    {docType !== "invoice" && (
                      <div className="meta-inner-field col-span-2">
                        <label className="block text-[7.5pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Requisition No.:</label>
                        <input 
                          type="text" 
                          value={requisitionNo}
                          onChange={(e) => setRequisitionNo(e.target.value)}
                          placeholder="Requisition number"
                          className="w-full border-b border-dotted border-slate-400 focus:border-black font-mono text-[9pt] outline-none bg-transparent py-0.5 no-print print:hidden"
                        />
                        <div className="hidden print:block font-mono text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5 break-words">
                          {requisitionNo || " "}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </td>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td className="border-none p-0 m-0">

                {/* Main Data Sheet Table */}
                <div className="w-full overflow-x-auto no-scrollbar -mx-4 px-4 sm:mx-0 sm:px-0 mt-2">
                  <table className="main-table w-full min-w-full border-collapse border-[1.5px] border-black table-fixed text-[9pt]">
                    <thead>
                      <tr className="bg-slate-50 text-[8pt]">
                        <th className={`${docType === 'challan' ? 'w-[7%]' : 'w-[6%]'} border border-black py-1 text-center font-bold`}>SL</th>
                        <th className={`${docType === 'challan' ? 'w-[75%]' : 'w-[56%]'} border border-black py-1 text-left px-3 font-bold`}>Description of Marine Items / Spare Parts</th>
                        <th className={`${docType === 'challan' ? 'w-[9%]' : 'w-[7%]'} border border-black py-1 text-center font-bold`}>Qty</th>
                        <th className={`${docType === 'challan' ? 'w-[9%]' : 'w-[8%]'} border border-black py-1 text-center font-bold`}>Unit</th>
                        {docType !== "challan" && (
                          <>
                            <th className="w-[11%] border border-black py-1 text-center font-bold">Price</th>
                            <th className="w-[12%] border border-black py-1 text-center font-bold">Amount</th>
                          </>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((row, idx) => (
                        <tr 
                          key={idx} 
                          className={`group hover:bg-slate-50/50 transition-colors ${
                            idx === safeSelectedRowIndex ? "bg-indigo-50/10" : ""
                          }`}
                        >
                          {GRID_COLUMNS.map((colIndex) => {
                            const { region, isAnchor } = getMergeInfo(idx, colIndex);
                            if (region && !isAnchor) return null;

                            const colSpan = region ? region.endCol - region.startCol + 1 : 1;
                            const rowSpan = region ? region.endRow - region.startRow + 1 : 1;

                            if (colIndex === -1) {
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={getCellStyle(idx, -1)}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, -1)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, -1)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, -1)}
                                  onClick={() => handleCellClick(idx, -1)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, -1)}
                                  className={getCellClassName(idx, -1, `border border-black text-center font-mono text-[8pt] align-top py-0.5 px-0.5 whitespace-nowrap transition-all cursor-pointer select-none bg-slate-50/30 text-slate-800`)}
                                >
                                  {idx + 1}
                                </td>
                              );
                            }

                            if (colIndex === 0) {
                              const cellStyle = getCellStyle(idx, 0);
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={cellStyle}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 0)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 0)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 0)}
                                  onClick={() => handleCellClick(idx, 0)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 0)}
                                  className={getCellClassName(idx, 0, `border border-black text-left px-2 text-[8.5pt] align-top py-0.5 whitespace-normal transition-all cursor-text ${region ? "bg-amber-50/10" : ""}`)}
                                >
                                  <RichTextCell
                                    value={row.desc}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 0 });
                                    }}
                                    onChange={(val) => handleRowChange(idx, "desc", val)}
                                    onKeyDown={(e) => {
                                      if (e.key === "Enter" && !e.shiftKey) {
                                        e.preventDefault();
                                        const nextArea = document.querySelector(`[data-row="${idx + 1}"][data-col="0"]`) as HTMLElement | null;
                                        if (nextArea) nextArea.focus();
                                      } else {
                                        handleKeyDown(e, idx, 0);
                                      }
                                    }}
                                    onPaste={(e) => handlePaste(e, idx, 0)}
                                    dataRow={idx}
                                    dataCol={0}
                                    style={cellStyle}
                                    className="w-full min-w-full text-left border-none outline-none bg-transparent p-0 text-slate-800 text-[8.5pt] leading-normal block overflow-hidden py-0.5 whitespace-normal break-words no-print print:hidden font-normal"
                                  />
                                  <div 
                                    style={cellStyle} 
                                    className="hidden print:block whitespace-normal break-words text-left text-slate-900 leading-normal py-0.5 text-[8.5pt] font-normal"
                                    dangerouslySetInnerHTML={{ __html: row.desc || "&nbsp;" }}
                                  />
                                </td>
                              );
                            }

                            if (colIndex === 1) {
                              const cellStyle = getCellStyle(idx, 1);
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={cellStyle}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 1)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 1)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 1)}
                                  onClick={() => handleCellClick(idx, 1)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 1)}
                                  className={getCellClassName(idx, 1, "border border-black text-center font-mono text-[8.5pt] align-top py-0.5 transition-all cursor-text")}
                                >
                                  <RichTextCell
                                    value={row.qty}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 1 });
                                    }}
                                    onChange={(val) => handleRowChange(idx, "qty", val)}
                                    onKeyDown={(e) => handleKeyDown(e, idx, 1)}
                                    onPaste={(e) => handlePaste(e, idx, 1)}
                                    dataRow={idx}
                                    dataCol={1}
                                    style={cellStyle}
                                    className="w-full text-center border-none outline-none bg-transparent px-0 font-mono text-slate-800 align-top overflow-hidden py-0.5 whitespace-normal break-normal no-print print:hidden text-[8.5pt]"
                                  />
                                  <div 
                                    style={cellStyle} 
                                    className="hidden print:block whitespace-normal break-normal text-center font-mono text-slate-900 py-0.5 text-[8.5pt]"
                                    dangerouslySetInnerHTML={{ __html: row.qty || "&nbsp;" }}
                                  />
                                </td>
                              );
                            }

                            if (colIndex === 2) {
                              const cellStyle = getCellStyle(idx, 2);
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={cellStyle}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 2)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 2)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 2)}
                                  onClick={() => handleCellClick(idx, 2)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 2)}
                                  className={getCellClassName(idx, 2, "border border-black text-center text-[8.5pt] align-top py-0.5 transition-all cursor-text")}
                                >
                                  <RichTextCell
                                    value={row.unit}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 2 });
                                    }}
                                    onChange={(val) => handleRowChange(idx, "unit", val)}
                                    onKeyDown={(e) => handleKeyDown(e, idx, 2)}
                                    onPaste={(e) => handlePaste(e, idx, 2)}
                                    dataRow={idx}
                                    dataCol={2}
                                    style={cellStyle}
                                    className="w-full text-center border-none outline-none bg-transparent px-0 text-slate-800 align-top overflow-hidden py-0.5 whitespace-normal break-normal no-print print:hidden text-[8.5pt]"
                                  />
                                  <div 
                                    style={cellStyle} 
                                    className="hidden print:block whitespace-normal break-normal text-center text-slate-900 py-0.5 text-[8.5pt]"
                                    dangerouslySetInnerHTML={{ __html: row.unit || "&nbsp;" }}
                                  />
                                </td>
                              );
                            }

                            if (colIndex === 3) {
                              const cellStyle = getCellStyle(idx, 3);
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  style={cellStyle}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 3)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 3)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 3)}
                                  onClick={() => handleCellClick(idx, 3)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 3)}
                                  className={getCellClassName(idx, 3, "border border-black text-center font-mono text-[8.5pt] align-top py-0.5 transition-all cursor-text")}
                                >
                                  <RichTextCell
                                    value={row.price}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 3 });
                                    }}
                                    onChange={(val) => handleRowChange(idx, "price", val)}
                                    onKeyDown={(e) => handleKeyDown(e, idx, 3)}
                                    onPaste={(e) => handlePaste(e, idx, 3)}
                                    dataRow={idx}
                                    dataCol={3}
                                    style={cellStyle}
                                    className="w-full text-center border-none outline-none bg-transparent px-0 font-mono text-slate-800 align-top overflow-hidden py-0.5 whitespace-normal break-normal no-print print:hidden text-[8.5pt]"
                                  />
                                  <div 
                                    style={cellStyle} 
                                    className="hidden print:block whitespace-normal break-normal text-center font-mono text-slate-900 py-0.5 text-[8.5pt]"
                                    dangerouslySetInnerHTML={{ __html: row.price || "&nbsp;" }}
                                  />
                                </td>
                              );
                            }

                            const cellStyle = getCellStyle(idx, 4);
                            return (
                              <td
                                key={colIndex}
                                colSpan={colSpan}
                                rowSpan={rowSpan}
                                style={cellStyle}
                                onMouseDown={(e) => handleCellMouseDown(e, idx, 4)}
                                onMouseEnter={() => handleCellMouseEnter(idx, 4)}
                                onMouseUp={(e) => handleCellMouseUp(e, idx, 4)}
                                onClick={() => handleCellClick(idx, 4)}
                                onContextMenu={(e) => handleCellContextMenu(e, idx, 4)}
                                className={getCellClassName(idx, 4, "border border-black text-right pr-2 font-mono text-[8.5pt] font-semibold text-slate-800 align-top py-0.5 transition-all cursor-pointer")}
                              >
                                <div style={cellStyle} className="whitespace-normal break-all leading-tight text-[8.5pt]">
                                  {row.amount !== 0 ? row.amount.toLocaleString("en-US", { minimumFractionDigits: 2 }) : "0.00"}
                                </div>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Grid Line Actions & Dynamic Row Controller */}
                <div className="no-print print:hidden my-2.5 px-3 py-2 bg-gradient-to-r from-slate-50 via-indigo-50/30 to-slate-50 rounded-xl border border-slate-200/80 shadow-xs flex flex-wrap items-center justify-between gap-3">
                  {/* Plus/Minus & Batch Line Adders */}
                  <div className="flex flex-wrap items-center gap-2">
                    {/* Add 1 Line */}
                    <button
                      type="button"
                      onClick={() => addRows(1)}
                      className="bg-indigo-600 hover:bg-indigo-700 active:scale-95 text-white font-bold text-xs h-8 px-3 rounded-lg shadow-xs transition-all cursor-pointer flex items-center gap-1.5"
                      title="Add 1 line (+)"
                    >
                      <Plus className="h-4 w-4 stroke-[2.5]" />
                      <span>+1 Line</span>
                    </button>

                    {/* Subtract 1 Line */}
                    <button
                      type="button"
                      onClick={removeRow}
                      disabled={rows.length <= 1}
                      className="bg-white hover:bg-rose-50 active:scale-95 border border-rose-200 text-rose-600 font-bold text-xs h-8 px-2.5 rounded-lg shadow-xs transition-all cursor-pointer flex items-center gap-1 disabled:opacity-40 disabled:cursor-not-allowed"
                      title="Subtract line (-)"
                    >
                      <Minus className="h-4 w-4 stroke-[2.5]" />
                      <span>-1</span>
                    </button>

                    <div className="h-5 w-[1px] bg-slate-200 mx-0.5 hidden sm:block" />

                    {/* Minimal quick add multiple lines */}
                    <form
                      onSubmit={(e) => {
                        e.preventDefault();
                        const num = parseInt(customRowCountInput, 10);
                        if (!isNaN(num) && num > 0) {
                          addRows(num);
                        }
                      }}
                      className="flex items-center gap-1 bg-white border border-slate-300 hover:border-indigo-400 focus-within:border-indigo-500 focus-within:ring-2 focus-within:ring-indigo-100 rounded-lg p-0.5 transition-all shadow-2xs"
                    >
                      <span className="text-[11px] font-semibold text-slate-500 pl-2">Add</span>
                      <input
                        type="number"
                        min="1"
                        max="1000"
                        value={customRowCountInput}
                        onChange={(e) => setCustomRowCountInput(e.target.value)}
                        placeholder="10"
                        className="w-14 h-6.5 text-center text-xs font-mono font-bold text-slate-800 border-none outline-none focus:outline-none bg-slate-50 rounded px-1"
                        title="Enter number of lines to add at once (max 1,000 total)"
                      />
                      <button
                        type="submit"
                        disabled={rows.length >= 1000}
                        className="bg-indigo-50 hover:bg-indigo-600 text-indigo-700 hover:text-white disabled:opacity-40 disabled:hover:bg-indigo-50 disabled:hover:text-indigo-700 font-bold text-xs h-6.5 px-2.5 rounded transition-all cursor-pointer flex items-center gap-1"
                        title="Add specified number of lines (limit 1,000)"
                      >
                        <Plus className="h-3 w-3 stroke-[2.5]" />
                        <span>Lines</span>
                      </button>
                    </form>

                    {/* Quick Preset Buttons */}
                    <div className="hidden md:flex items-center gap-1">
                      {[5, 10, 20, 50].map((count) => (
                        <button
                          key={count}
                          type="button"
                          disabled={rows.length >= 1000}
                          onClick={() => addRows(count)}
                          className="bg-white hover:bg-slate-100 active:scale-95 disabled:opacity-40 text-slate-600 hover:text-indigo-600 border border-slate-200 font-semibold text-[11px] h-7 px-2 rounded-md transition-all cursor-pointer"
                          title={`Quick add +${count} lines`}
                        >
                          +{count}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Status counter */}
                  <div className="flex items-center gap-1.5 text-[10px] text-slate-500 font-semibold shrink-0">
                    <span className="bg-slate-200/80 text-slate-700 px-2.5 py-1 rounded-md font-mono text-[11px]" title="Total Lines / 1,000 Max Limit">
                      Total: <strong className="text-slate-900">{rows.length}</strong> / 1000
                    </span>
                    <span className="bg-indigo-100 text-indigo-800 px-2.5 py-1 rounded-md font-mono text-[11px]">
                      Filled: <strong>{rows.filter(r => String(r.desc || "").trim() || String(r.qty || "").trim() || String(r.unit || "").trim() || String(r.price || "").trim()).length}</strong>
                    </span>
                  </div>
                </div>

                {/* Bottom closing wraps, sums, signatures */}
                <div className="closing-wrap mt-1.5">
                  {docType !== "challan" && (
                    <table className="closing-row w-full border-collapse border-2 border-black mt-1.5 bg-white text-black z-10 relative">
                      <tbody>
                        {docType === "invoice" ? (
                          <>
                            <tr className="align-stretch">
                              <td rowSpan={4} className="amount-words-container w-1/2 border-r-2 border-black p-1.5 bg-slate-50/50 text-left align-middle">
                                <span className="font-extrabold text-[6.5pt] text-slate-700 uppercase tracking-wider block mb-0.5">
                                  Amount in Words:
                                </span>
                                <span className="text-[8pt] font-mono italic text-black font-black uppercase leading-tight">
                                  {numberToWords(calculatedGrandTotal)}
                                </span>
                              </td>
                              <td className="w-1/2 p-0 border-b border-black align-stretch">
                                <div className="flex flex-row items-stretch h-full min-h-[24px] w-full">
                                  <div className="total-lbl bg-slate-50 w-[170px] shrink-0 pr-2 text-right border-r-2 border-black text-[8pt] font-bold uppercase flex items-center justify-end tracking-wider">
                                    SUBTOTAL
                                  </div>
                                  <div className="total-val flex-grow text-right pr-4 text-[9pt] font-mono font-black flex items-center justify-end px-2 py-0.5 leading-tight">
                                    {rowsTotal.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                            <tr className="align-stretch">
                              <td className="w-1/2 p-0 border-b border-black align-stretch">
                                <div className="flex flex-row items-stretch h-full min-h-[24px] w-full">
                                  <div className="total-lbl bg-slate-50 w-[170px] shrink-0 pr-2 text-right border-r-2 border-black text-[8pt] font-bold uppercase flex items-center justify-end tracking-wider">
                                    <div className="flex items-center justify-end gap-1.5 w-full pl-2">
                                      <span>VAT</span>
                                      <div className="flex items-center gap-0.5 no-print print:hidden shrink-0">
                                        <input
                                          type="text"
                                          value={vatPercent}
                                          onChange={(e) => {
                                            const val = e.target.value;
                                            if (val === "" || /^-?\d*[.,]?\d*$/.test(val)) {
                                              setVatPercent(val);
                                            }
                                          }}
                                          className="w-10 text-center border border-slate-300 rounded font-mono text-[8pt] bg-white text-slate-800 py-0.5"
                                        />
                                        <span>%</span>
                                      </div>
                                      <span className="hidden print:inline">({parsedVatPercent}%)</span>
                                    </div>
                                  </div>
                                  <div className="total-val flex-grow text-right pr-4 text-[9pt] font-mono font-semibold flex items-center justify-end px-2 py-0.5 leading-tight">
                                    {vatAmount.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                            <tr className="align-stretch">
                              <td className="w-1/2 p-0 border-b border-black align-stretch">
                                <div className="flex flex-row items-stretch h-full min-h-[24px] w-full">
                                  <div className="total-lbl bg-slate-50 w-[170px] shrink-0 pr-2 text-right border-r-2 border-black text-[8pt] font-bold uppercase flex items-center justify-end tracking-wider">
                                    <div className="flex items-center justify-end gap-1.5 w-full pl-2">
                                      <span>TRANS.</span>
                                      <div className="flex items-center no-print print:hidden shrink-0">
                                        <input
                                          type="text"
                                          value={transportationFee}
                                          onChange={(e) => {
                                            const val = e.target.value;
                                            if (val === "" || /^-?\d*[.,]?\d*$/.test(val)) {
                                              setTransportationFee(val);
                                            }
                                          }}
                                          placeholder="0"
                                          className="w-14 text-center border border-slate-300 rounded font-mono text-[8pt] bg-white text-slate-800 py-0.5"
                                        />
                                      </div>
                                    </div>
                                  </div>
                                  <div className="total-val flex-grow text-right pr-4 text-[9pt] font-mono font-semibold flex items-center justify-end px-2 py-0.5 leading-tight">
                                    {parsedTransportationFee.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                            <tr className="align-stretch">
                              <td className="w-1/2 p-0 align-stretch">
                                <div className="flex flex-row items-stretch h-full min-h-[24px] w-full">
                                  <div className="total-lbl bg-indigo-50/40 w-[170px] shrink-0 pr-2 text-right border-r-2 border-black text-[8.5pt] font-black uppercase flex items-center justify-end tracking-wider text-indigo-950">
                                    GRAND TOTAL
                                  </div>
                                  <div className="total-val flex-grow text-right pr-4 text-[10pt] font-mono font-black flex items-center justify-end px-2 py-0.5 leading-tight text-indigo-950">
                                    {grandTotal.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                  </div>
                                </div>
                              </td>
                            </tr>
                          </>
                        ) : (
                          <tr className="align-stretch">
                            <td className="amount-words-container w-1/2 border-r-2 border-black p-1.5 bg-slate-50/50 text-left align-middle">
                              <span className="font-extrabold text-[6.5pt] text-slate-700 uppercase tracking-wider block mb-0.5">
                                Amount in Words:
                              </span>
                              <span className="text-[8pt] font-mono italic text-black font-black uppercase leading-tight">
                                {numberToWords(calculatedGrandTotal)}
                              </span>
                            </td>
                            <td className="w-1/2 p-0 align-stretch">
                              <div className="flex flex-row items-stretch h-full min-h-[52.5px] w-full">
                                <div className="total-lbl bg-slate-50 w-[170px] shrink-0 pr-2 text-right border-r-2 border-black text-[8.5pt] font-bold uppercase flex items-center justify-end">
                                  TOTAL
                                </div>
                                <div className="total-val flex-grow text-right pr-4 text-[9.5pt] font-mono font-black flex items-center justify-end px-2 py-0.5 leading-tight min-h-[52.5px]">
                                  {grandTotal.toLocaleString("en-US", { minimumFractionDigits: 2 })}
                                </div>
                              </div>
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  )}

                  {/* Signatures & Stamps placing */}
                  <div className="sig-section mt-5 flex flex-row justify-between gap-6 sm:gap-10">
                    <div className="sig-box w-full sm:w-[220px] print:w-[220px] text-center flex flex-col justify-end h-[72px]">
                      <div className="sig-line border-t-[1.5px] border-black pt-1 text-[8.5pt] font-bold">
                        Receiver's Signature
                      </div>
                    </div>
                    
                    {/* Authorized stamp hidden for Challan block */}
                    {docType !== "challan" && (
                      <div className="sig-box w-full sm:w-[220px] print:w-[220px] text-center flex flex-col justify-between h-[72px] relative">
                        <div className="sig-title text-[8.5pt] font-bold text-black">For Comilla Traders</div>
                        
                        <div className="absolute inset-0 flex items-center justify-center pointer-events-none z-10 select-none pb-1">
                          <img 
                            src="https://i.ibb.co.com/jZswrtn6/image-4-removebg-preview.png"
                            alt="Comilla Traders Stamp"
                            referrerPolicy="no-referrer"
                            className="w-[90px] h-[90px] object-contain select-none"
                            style={{ printColorAdjust: "exact" }}
                          />
                        </div>

                        <div className="sig-line border-t-[1.5px] border-black pt-1 text-[8.5pt] font-bold relative z-20">
                          Authorized Signature
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Non-returnable & non-exchangeable notice */}
                  <div className="doc-footer-notice text-center mt-3 pt-1 text-[12px] leading-[18px] font-bold text-black uppercase tracking-wider">
                    ITEMS ONCE SOLD ARE NON-RETURNABLE AND NON-EXCHANGEABLE.
                  </div>
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Online Document Search, Lists, and Documentation Panel */}
      <SavedDocumentsPanel
        savedDocs={savedDocs}
        currentDocId={currentDocId}
        searchQuery={searchQuery}
        setSearchQuery={setSearchQuery}
        selectedTypeFilter={selectedTypeFilter}
        setSelectedTypeFilter={setSelectedTypeFilter}
        loadSavedDoc={loadSavedDoc}
        deleteSavedDoc={deleteSavedDoc}
        renameSavedDoc={renameSavedDoc}
      />

      {/* Cell right-click Menu context */}
      {contextMenu && contextMenu.visible && (
        <div 
          className="fixed bg-white border border-slate-200 rounded-lg shadow-xl py-1.5 w-64 z-[9999] text-xs text-slate-700"
          style={{ top: `${contextMenu.y}px`, left: `${contextMenu.x}px` }}
          onClick={(e) => e.stopPropagation()}
        >
          {contextMenu.colIndex !== undefined && contextMenu.colIndex >= 0 && contextMenu.colIndex <= 3 && (
            <>
              <button 
                onClick={() => {
                  clearSpecificCell(contextMenu.rowIndex, contextMenu.colIndex!);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold text-slate-900"
              >
                Clear Cell Content
              </button>
              <div className="my-1 border-t border-slate-100"></div>
            </>
          )}

          <button 
            onClick={() => {
              toggleMergeSelectedRange();
              setContextMenu(null);
            }}
            disabled={!hasRangeSelection() && !getMergeRegionAt(contextMenu.rowIndex, contextMenu.colIndex)}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 disabled:opacity-40 font-bold text-slate-900"
          >
            {getMergeRegionAt(contextMenu.rowIndex, contextMenu.colIndex) ? "Unmerge Cells" : "Merge Selected Cells"}
          </button>

          <div className="my-1 border-t border-slate-100"></div>

          <button 
            onClick={() => {
              setIsExcelModalOpen(true);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-emerald-50 text-emerald-800 font-bold flex items-center gap-1.5"
          >
            <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-600" />
            <span>Paste / Import from Excel...</span>
          </button>

          <button 
            onClick={() => {
              unwrapAllDescriptions();
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-indigo-50 text-indigo-800 font-bold flex items-center gap-1.5"
          >
            <WrapText className="h-3.5 w-3.5 text-indigo-600" />
            <span>Unwrap Line Breaks (Continuous Flow)</span>
          </button>

          <div className="my-1 border-t border-slate-100"></div>

          <button 
            onClick={() => {
              insertMultipleRows(contextMenu.rowIndex, 'above', 1);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold"
          >
            Insert 1 Row Above
          </button>
          <button 
            onClick={() => {
              insertMultipleRows(contextMenu.rowIndex, 'below', 1);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold"
          >
            Insert 1 Row Below
          </button>
          <button 
            onClick={() => {
              insertMultipleRows(contextMenu.rowIndex, 'below', 5);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold text-indigo-700"
          >
            Insert +5 Rows Below
          </button>

          <div className="my-1 border-t border-slate-100"></div>

          <button 
            onClick={() => {
              if (contextMenu.rowIndex > 0) {
                moveRow(contextMenu.rowIndex, 'up');
                setContextMenu(null);
              }
            }}
            disabled={contextMenu.rowIndex === 0}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 disabled:opacity-40 font-bold"
          >
            Move Row Up
          </button>
          <button 
            onClick={() => {
              if (contextMenu.rowIndex < rows.length - 1) {
                moveRow(contextMenu.rowIndex, 'down');
                setContextMenu(null);
              }
            }}
            disabled={contextMenu.rowIndex === rows.length - 1}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 disabled:opacity-40 font-bold"
          >
            Move Row Down
          </button>

          <div className="my-1 border-t border-slate-100"></div>

          <button 
            onClick={() => {
              clearSpecificRow(contextMenu.rowIndex);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold"
          >
            Clear Row Content
          </button>
          <button 
            onClick={() => {
              deleteSpecificRow(contextMenu.rowIndex);
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-rose-50 text-rose-600 font-bold border-t border-rose-50 mt-1"
          >
            Delete Row
          </button>
        </div>
      )}

      {/* Excel Smart Importer Modal */}
      <ExcelPasteModal
        isOpen={isExcelModalOpen}
        onClose={() => setIsExcelModalOpen(false)}
        onImportRows={handleImportFromExcel}
        selectedRowIndex={safeSelectedRowIndex}
        totalCurrentRows={rows.length}
        docType={docType}
      />

      {/* Floating Status Toast */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-[999999] bg-slate-900/95 backdrop-blur-xs text-white text-xs font-bold px-4 py-2.5 rounded-xl shadow-2xl border border-slate-700/80 flex items-center gap-2 animate-in fade-in slide-in-from-bottom-4 duration-250">
          <div className="bg-emerald-500 text-white rounded-full p-1">
            <CheckCheck className="h-3.5 w-3.5" />
          </div>
          <span>{toastMessage.text}</span>
        </div>
      )}
    </div>
  );
}
