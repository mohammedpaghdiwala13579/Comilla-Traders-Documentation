import React, { useState, useEffect, useRef } from "react";
// @ts-ignore
import html2pdf from "html2pdf.js";
import { Download, Printer, Calendar, Save, Trash2, Plus, Check, RefreshCw, Copy, X } from "lucide-react";
import { db } from "../lib/firebase";
import { collection, doc, setDoc, deleteDoc, onSnapshot, query, orderBy } from "firebase/firestore";
import { numberToWords } from "../utils/numberToWords";
import { parseTSV } from "../utils/tsvParser";
import { generateExcelWorkbook } from "../utils/excelGenerator";
import { QuotationRow, MergedRegion, SavedDocument } from "../types";
import SavedDocumentsPanel from "./SavedDocumentsPanel";

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

export default function QuotationBuilder() {
  const [docType, setDocType] = useState<"quotation" | "challan" | "invoice">("quotation");
  const [dateVal, setDateVal] = useState(() => {
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, "0");
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy = today.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  });
  const [messers, setMessers] = useState("");
  const [address, setAddress] = useState("");
  const [challanNo, setChallanNo] = useState("");
  const [requisitionNo, setRequisitionNo] = useState("");
  const [invoiceNo, setInvoiceNo] = useState("");
  const [poNumber, setPoNumber] = useState("");
  const [vatPercent, setVatPercent] = useState<string>("0");
  const [transportationFee, setTransportationFee] = useState<string>("0");
  const [isGeneratingExcel, setIsGeneratingExcel] = useState(false);
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);

  // In-app storage & Auto-Save states
  const [savedDocs, setSavedDocs] = useState<SavedDocument[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedTypeFilter, setSelectedTypeFilter] = useState<"all" | "quotation" | "challan" | "invoice">("all");
  const [currentDocId, setCurrentDocId] = useState<string | null>(null);
  const [autoSaveEnabled, setAutoSaveEnabled] = useState<boolean>(() => {
    if (typeof window !== "undefined") {
      const val = localStorage.getItem("comilla_autosave_enabled");
      return val === null ? true : val === "true";
    }
    return true;
  });
  const [lastSavedTime, setLastSavedTime] = useState<string | null>(null);
  const [saveStatus, setSaveStatus] = useState<"idle" | "saving" | "saved" | "error">("idle");

  // Grid rows: starts with 20 blank rows by default
  const [rows, setRows] = useState<QuotationRow[]>(() => {
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

  const [mergedRegions, setMergedRegions] = useState<MergedRegion[]>([]);
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
          mergedRegions: docMergedRegions
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
    setCurrentDocId(null);
    setLastSavedTime(null);
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
    autoSaveEnabled,
    currentDocId,
    vatPercent,
    transportationFee
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

      const q = parseFloat(String(targetRow.qty || "")) || 0;
      const p = parseFloat(String(targetRow.price || "").replace(/,/g, "")) || 0;
      targetRow.amount = q * p;

      updated[index] = targetRow;
      return updated;
    });
  };

  const addRow = () => {
    setRows((prevRows) => [
      ...prevRows,
      {
        sl: prevRows.length + 1,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      },
    ]);
  };

  const removeRow = () => {
    setRows((prevRows) => {
      if (prevRows.length <= 1) return prevRows;
      return prevRows.slice(0, -1);
    });
  };

  const insertRow = (index: number, position: 'above' | 'below') => {
    const insertAt = position === 'above' ? index : index + 1;
    setRows((prevRows) => {
      const updated = [...prevRows];
      const newRow: QuotationRow = {
        sl: 0,
        desc: "",
        qty: "",
        unit: "",
        price: "",
        amount: 0,
      };
      updated.splice(insertAt, 0, newRow);
      return updated.map((r, i) => ({
        ...r,
        sl: i + 1
      }));
    });

    setMergedRegions((prevRegions) =>
      prevRegions.map((region) => {
        if (region.startRow >= insertAt) {
          return { ...region, startRow: region.startRow + 1, endRow: region.endRow + 1 };
        }
        if (region.endRow >= insertAt) {
          return { ...region, endRow: region.endRow + 1 };
        }
        return region;
      })
    );
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
        const q = parseFloat(String(updated[r].qty || "")) || 0;
        const p = parseFloat(String(updated[r].price || "").replace(/,/g, "")) || 0;
        updated[r].amount = q * p;
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

    const isActive = selectedCell?.rowIndex === rowIndex && selectedCell?.colIndex === colIndex;

    setIsSelecting(true);
    setSelectionStart({ rowIndex, colIndex });
    setSelectionEnd({ rowIndex, colIndex });
    setSelectedCell({ rowIndex, colIndex });
    setSelectedRowIndex(rowIndex);

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
  const parsedVatPercent = parseFloat(vatPercent) || 0;
  const parsedTransportationFee = parseFloat(transportationFee) || 0;
  const vatAmount = docType === "invoice" ? (rowsTotal * parsedVatPercent) / 100 : 0;
  const grandTotal = docType === "invoice" ? (rowsTotal + vatAmount + parsedTransportationFee) : rowsTotal;
  const calculatedGrandTotal = docType === "challan" ? 0 : grandTotal;

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>, rowIndex: number, colIndex: number) => {
    const { key } = e;
    let targetRow = rowIndex;
    let targetCol = colIndex;

    const cursorStart = e.currentTarget.selectionStart;
    const cursorEnd = e.currentTarget.selectionEnd;
    const valueLength = e.currentTarget.value.length;

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
      targetRow = rowIndex + 1;
    } else {
      return;
    }

    e.preventDefault();
    const targetElement = document.querySelector(
      `[data-row="${targetRow}"][data-col="${targetCol}"]`
    ) as HTMLInputElement | HTMLTextAreaElement | null;

    if (targetElement) {
      targetElement.focus();
      if (typeof targetElement.select === "function") {
        targetElement.select();
      }
    }
  };

  const handlePaste = (
    e: React.ClipboardEvent<HTMLTextAreaElement | HTMLInputElement>,
    startRowIndex: number,
    startColIndex: number
  ) => {
    const clipboardData = e.clipboardData.getData("text");
    if (!clipboardData) return;

    if (clipboardData.includes("\t") || clipboardData.includes("\n") || clipboardData.includes("\r")) {
      e.preventDefault();
      const parsedGrid = parseTSV(clipboardData);

      if (parsedGrid.length === 0) return;

      if (parsedGrid.length === 1 && parsedGrid[0].length === 1) {
        const parsedVal = parsedGrid[0][0];
        const textarea = e.currentTarget as HTMLTextAreaElement;
        const start = textarea.selectionStart ?? 0;
        const end = textarea.selectionEnd ?? 0;
        const currentValue = textarea.value;
        const newValue = currentValue.substring(0, start) + parsedVal + currentValue.substring(end);
        
        const fieldMap = ["desc", "qty", "unit", "price"] as const;
        const field = fieldMap[startColIndex];
        handleRowChange(startRowIndex, field, newValue);
        
        setTimeout(() => {
          textarea.focus();
          textarea.selectionStart = textarea.selectionEnd = start + parsedVal.length;
        }, 0);
        return;
      }

      setRows((prevRows) => {
        const updated = [...prevRows];
        parsedGrid.forEach((cols, rOffset) => {
          const rIndex = startRowIndex + rOffset;
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
          cols.forEach((cellValue, cOffset) => {
            const cIndex = startColIndex + cOffset;
            if (cIndex === 0) targetRow.desc = cellValue;
            else if (cIndex === 1) targetRow.qty = cellValue;
            else if (cIndex === 2) targetRow.unit = cellValue;
            else if (cIndex === 3 && docType !== "challan") targetRow.price = cellValue;
          });

          const q = parseFloat(String(targetRow.qty || "")) || 0;
          const p = parseFloat(String(targetRow.price || "").replace(/,/g, "")) || 0;
          targetRow.amount = q * p;
          updated[rIndex] = targetRow;
        });
        return updated;
      });
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
        parseFloat(vatPercent) || 0,
        parseFloat(transportationFee) || 0
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
    <div className="quotation-container relative min-h-screen flex flex-col items-center bg-slate-50 py-5 overflow-x-auto text-[#000] font-sans antialiased w-full">
      
      {/* Mini App Toolbar - Consolidated Single Bar */}
      <div className="top-toolbar no-print print:hidden w-full max-w-[210mm] mb-3 flex flex-col md:flex-row justify-between items-center gap-2 px-3 sm:px-0 z-10">
        <div className="flex flex-wrap items-center gap-2 w-full md:w-auto justify-between md:justify-start">
          <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200 shadow-sm shrink-0">
            <button
              type="button"
              onClick={() => {
                setDocType("quotation");
                setMergedRegions([]);
              }}
              className={`px-3 py-1 rounded-md text-[10px] font-bold transition-all cursor-pointer ${
                docType === "quotation"
                  ? "bg-indigo-600 text-white shadow-md shadow-indigo-100"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              Quotation
            </button>
            <button
              type="button"
              onClick={() => {
                setDocType("challan");
                setMergedRegions([]);
              }}
              className={`px-3 py-1 rounded-md text-[10px] font-bold transition-all cursor-pointer ${
                docType === "challan"
                  ? "bg-indigo-600 text-white shadow-md shadow-indigo-100"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              Challan
            </button>
            <button
              type="button"
              onClick={() => {
                setDocType("invoice");
                setMergedRegions([]);
              }}
              className={`px-3 py-1 rounded-md text-[10px] font-bold transition-all cursor-pointer ${
                docType === "invoice"
                  ? "bg-indigo-600 text-white shadow-md shadow-indigo-100"
                  : "text-slate-600 hover:text-slate-900"
              }`}
            >
              Invoice
            </button>
          </div>

          <div className="flex items-center gap-1.5 bg-white px-2 py-1 rounded-lg border border-slate-200 shadow-xs shrink-0">
            <label className="relative inline-flex items-center cursor-pointer select-none">
              <input
                type="checkbox"
                checked={autoSaveEnabled}
                onChange={(e) => {
                  setAutoSaveEnabled(e.target.checked);
                  localStorage.setItem("comilla_autosave_enabled", String(e.target.checked));
                }}
                className="sr-only peer"
              />
              <div className="w-7 h-4 bg-slate-200 rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[1.5px] after:left-[1.5px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-3 after:w-3 after:transition-all peer-checked:bg-emerald-500"></div>
              <span className="ml-1.5 text-[9px] font-bold text-slate-600 uppercase tracking-wider">Auto-Save</span>
            </label>
            {lastSavedTime && (
              <span className="text-[8px] text-emerald-600 font-medium flex items-center gap-1">
                <span className="h-1.5 w-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                <span className="hidden sm:inline">{lastSavedTime}</span>
              </span>
            )}
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-1.5 w-full md:w-auto justify-end">
          <button 
            onClick={startNewDoc} 
            className="bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 font-bold text-[10px] py-1 px-2.5 rounded-md shadow-sm hover:shadow transition-all cursor-pointer flex items-center gap-1"
            title="Start a fresh blank sheet"
          >
            <Plus className="h-3 w-3" />
            <span>NEW SHEET</span>
          </button>
          
          {currentDocId && (
            <>
              <button 
                onClick={duplicateCurrentDoc} 
                className="bg-white hover:bg-indigo-50 border border-indigo-200 text-indigo-700 font-bold text-[10px] py-1 px-2.5 rounded-md shadow-sm hover:shadow transition-all cursor-pointer flex items-center gap-1"
                title="Save a duplicated copy"
              >
                <Copy className="h-3 w-3" />
                <span>DUPLICATE</span>
              </button>
              <button 
                onClick={() => deleteSavedDoc(currentDocId)} 
                className="bg-white hover:bg-rose-50 border border-rose-200 text-rose-600 hover:text-rose-700 font-bold text-[10px] py-1 px-2.5 rounded-md shadow-sm hover:shadow transition-all cursor-pointer flex items-center gap-1"
                title="Delete this sheet"
              >
                <Trash2 className="h-3 w-3 text-rose-500" />
                <span>DELETE</span>
              </button>
            </>
          )}

          <button 
            onClick={() => saveCurrentDocToApp()} 
            disabled={saveStatus === "saving"}
            className={`${
              saveStatus === "saved" 
                ? "bg-emerald-600 hover:bg-emerald-700" 
                : saveStatus === "error" 
                ? "bg-rose-600 hover:bg-rose-700" 
                : "bg-indigo-600 hover:bg-indigo-700"
            } text-white font-bold text-[10px] py-1 px-2.5 rounded-md shadow-sm hover:shadow transition-all cursor-pointer flex items-center gap-1 disabled:opacity-85`}
            title="Save to Cloud Database"
          >
            {saveStatus === "saving" ? (
              <>
                <RefreshCw className="h-3 w-3 animate-spin" />
                <span>SAVING...</span>
              </>
            ) : saveStatus === "saved" ? (
              <>
                <Check className="h-3 w-3" />
                <span>SAVED</span>
              </>
            ) : (
              <>
                <Save className="h-3 w-3" />
                <span>SAVE</span>
              </>
            )}
          </button>

          <button 
            onClick={handleDownloadExcel}
            disabled={isGeneratingExcel}
            className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-md transition-all cursor-pointer flex items-center gap-1 font-bold text-[10px] shadow-sm hover:shadow"
            title="Export Excel (.xlsx)"
          >
            <Download className="h-3 w-3" />
            <span>{isGeneratingExcel ? "GENERATING..." : "EXPORT EXCEL"}</span>
          </button>

          <button 
            onClick={handlePrint}
            className="px-2.5 py-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded-md transition-all cursor-pointer flex items-center gap-1 font-bold text-[10px] shadow-sm hover:shadow"
            title="Print or Save as PDF"
          >
            <Printer className="h-3 w-3" />
            <span>PRINT A4</span>
          </button>
        </div>
      </div>

      {/* Editing State Banner */}
      {currentDocId && (
        <div className="no-print print:hidden w-full max-w-[210mm] mb-2 px-3 sm:px-0 z-10 animate-in fade-in slide-in-from-top-2 duration-250">
          <div className="bg-indigo-50 border border-indigo-100 rounded-lg p-2.5 flex items-center justify-between text-xs text-indigo-950 shadow-sm">
            <div className="flex items-center gap-2 min-w-0">
              <span className="bg-indigo-600 text-white font-black text-[8px] px-1.5 py-0.5 rounded-sm uppercase tracking-wider shrink-0 shadow-xs">
                Editing
              </span>
              <span className="font-bold text-slate-800 truncate text-[11px]" title={savedDocs.find(d => d.id === currentDocId)?.name || "Active Sheet"}>
                {savedDocs.find(d => d.id === currentDocId)?.name || "Active Sheet"}
              </span>
            </div>
            <button
              type="button"
              onClick={resetSheetFields}
              className="text-[10px] font-bold text-indigo-600 hover:text-indigo-800 hover:bg-indigo-100/50 px-2 py-1 rounded transition-all cursor-pointer uppercase tracking-wider shrink-0"
            >
              Close
            </button>
          </div>
        </div>
      )}

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
                      <input 
                        type="text" 
                        value={messers}
                        onChange={(e) => setMessers(e.target.value)}
                        placeholder="Enter Client/Ship details"
                        className="w-full border-b border-dotted border-slate-400 focus:border-black font-bold text-[9pt] outline-none bg-transparent py-0.5 no-print print:hidden"
                      />
                      <div className="hidden print:block font-bold text-[9pt] border-b border-dotted border-black min-h-[18px] py-0.5 break-words whitespace-pre-wrap leading-tight">
                        {messers || " "}
                      </div>
                    </div>
                    <div>
                      <label className="block text-[7pt] font-extrabold text-slate-700 uppercase tracking-wider mb-0.5">Address:</label>
                      <textarea 
                        rows={2}
                        value={address}
                        onChange={(e) => setAddress(e.target.value)}
                        placeholder="Enter delivery/billing address"
                        className="w-full border-b border-dotted border-slate-400 focus:border-black text-[8.5pt] outline-none bg-transparent resize-none leading-tight py-0.5 no-print print:hidden"
                      />
                      <div className="hidden print:block text-[8.5pt] border-b border-dotted border-black min-h-[32px] py-0.5 break-words whitespace-pre-wrap leading-tight">
                        {address || " "}
                      </div>
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
                  <table className="main-table w-[650px] sm:w-full border-collapse border-[1.5px] border-black table-fixed text-[9pt]">
                    <thead>
                      <tr className="bg-slate-50 text-[8pt]">
                        <th className={`${docType === 'challan' ? 'w-[7%]' : 'w-[5%]'} border border-black py-1 text-center font-bold`}>SL</th>
                        <th className={`${docType === 'challan' ? 'w-[68%]' : 'w-[45%]'} border border-black py-1 text-left px-2 font-bold`}>Description of Marine Items / Spare Parts</th>
                        <th className={`${docType === 'challan' ? 'w-[12%]' : 'w-[8%]'} border border-black py-1 text-center font-bold`}>Qty</th>
                        <th className={`${docType === 'challan' ? 'w-[13%]' : 'w-[12%]'} border border-black py-1 text-center font-bold`}>Unit</th>
                        {docType !== "challan" && (
                          <>
                            <th className="w-[12%] border border-black py-1 text-center font-bold">Price</th>
                            <th className="w-[18%] border border-black py-1 text-center font-bold">Amount</th>
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
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, -1)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, -1)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, -1)}
                                  onClick={() => handleCellClick(idx, -1)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, -1)}
                                  className={getCellClassName(idx, -1, `border border-black text-center font-mono text-[8pt] align-top py-0.5 transition-all cursor-pointer select-none bg-slate-50/30 text-slate-800`)}
                                >
                                  {idx + 1}
                                </td>
                              );
                            }

                            if (colIndex === 0) {
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 0)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 0)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 0)}
                                  onClick={() => handleCellClick(idx, 0)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 0)}
                                  className={getCellClassName(idx, 0, `border border-black text-left px-1.5 text-[8.5pt] align-top py-0.5 break-all whitespace-normal transition-all cursor-text ${region ? "bg-amber-50/10" : ""}`)}
                                >
                                  <textarea
                                    value={row.desc}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 0 });
                                    }}
                                    onChange={(e) => {
                                      handleRowChange(idx, "desc", e.target.value);
                                      e.target.style.height = "auto";
                                      e.target.style.height = `${e.target.scrollHeight}px`;
                                    }}
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
                                    data-row={idx}
                                    data-col={0}
                                    rows={1}
                                    style={{ height: "auto", resize: "none" }}
                                    className={`w-full text-left border-none outline-none bg-transparent px-0 text-slate-800 text-[8.5pt] leading-tight block overflow-hidden py-0.5 whitespace-pre-wrap break-all no-print print:hidden`}
                                  />
                                  <div className="hidden print:block whitespace-pre-wrap break-words text-slate-900 leading-tight py-0.5 text-[8.5pt]">
                                    {row.desc || " "}
                                  </div>
                                </td>
                              );
                            }

                            if (colIndex === 1) {
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 1)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 1)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 1)}
                                  onClick={() => handleCellClick(idx, 1)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 1)}
                                  className={getCellClassName(idx, 1, "border border-black text-center font-mono text-[8.5pt] align-top py-0.5 transition-all cursor-text")}
                                >
                                  <textarea
                                    value={row.qty}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 1 });
                                    }}
                                    onChange={(e) => handleRowChange(idx, "qty", e.target.value)}
                                    onKeyDown={(e) => handleKeyDown(e, idx, 1)}
                                    onPaste={(e) => handlePaste(e, idx, 1)}
                                    data-row={idx}
                                    data-col={1}
                                    rows={1}
                                    style={{ height: "auto", resize: "none" }}
                                    className="w-full text-center border-none outline-none bg-transparent px-0 font-mono text-slate-800 align-top overflow-hidden py-0.5 whitespace-pre-wrap break-all no-print print:hidden text-[8.5pt]"
                                  />
                                  <div className="hidden print:block whitespace-pre-wrap break-words text-center font-mono text-slate-900 py-0.5 text-[8.5pt]">
                                    {row.qty || " "}
                                  </div>
                                </td>
                              );
                            }

                            if (colIndex === 2) {
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 2)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 2)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 2)}
                                  onClick={() => handleCellClick(idx, 2)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 2)}
                                  className={getCellClassName(idx, 2, "border border-black text-center text-[8.5pt] align-top py-0.5 transition-all cursor-text")}
                                >
                                  <textarea
                                    value={row.unit}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 2 });
                                    }}
                                    onChange={(e) => handleRowChange(idx, "unit", e.target.value)}
                                    onKeyDown={(e) => handleKeyDown(e, idx, 2)}
                                    onPaste={(e) => handlePaste(e, idx, 2)}
                                    data-row={idx}
                                    data-col={2}
                                    rows={1}
                                    style={{ height: "auto", resize: "none" }}
                                    className="w-full text-center border-none outline-none bg-transparent px-0 text-slate-800 align-top overflow-hidden py-0.5 whitespace-pre-wrap break-all no-print print:hidden text-[8.5pt]"
                                  />
                                  <div className="hidden print:block whitespace-pre-wrap break-words text-center text-slate-900 py-0.5 text-[8.5pt]">
                                    {row.unit || " "}
                                  </div>
                                </td>
                              );
                            }

                            if (colIndex === 3) {
                              return (
                                <td
                                  key={colIndex}
                                  colSpan={colSpan}
                                  rowSpan={rowSpan}
                                  onMouseDown={(e) => handleCellMouseDown(e, idx, 3)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 3)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 3)}
                                  onClick={() => handleCellClick(idx, 3)}
                                  onContextMenu={(e) => handleCellContextMenu(e, idx, 3)}
                                  className={getCellClassName(idx, 3, "border border-black text-center font-mono text-[8.5pt] align-top py-0.5 transition-all cursor-text")}
                                >
                                  <textarea
                                    value={row.price}
                                    onFocus={() => {
                                      setSelectedRowIndex(idx);
                                      setSelectedCell({ rowIndex: idx, colIndex: 3 });
                                    }}
                                    onChange={(e) => handleRowChange(idx, "price", e.target.value)}
                                    onKeyDown={(e) => handleKeyDown(e, idx, 3)}
                                    onPaste={(e) => handlePaste(e, idx, 3)}
                                    data-row={idx}
                                    data-col={3}
                                    rows={1}
                                    style={{ height: "auto", resize: "none" }}
                                    className="w-full text-center border-none outline-none bg-transparent px-0 font-mono text-slate-800 align-top overflow-hidden py-0.5 whitespace-pre-wrap break-all no-print print:hidden text-[8.5pt]"
                                  />
                                  <div className="hidden print:block whitespace-pre-wrap break-words text-center font-mono text-slate-900 py-0.5 text-[8.5pt]">
                                    {row.price || " "}
                                  </div>
                                </td>
                              );
                            }

                            return (
                              <td
                                key={colIndex}
                                colSpan={colSpan}
                                rowSpan={rowSpan}
                                onMouseDown={(e) => handleCellMouseDown(e, idx, 4)}
                                  onMouseEnter={() => handleCellMouseEnter(idx, 4)}
                                  onMouseUp={(e) => handleCellMouseUp(e, idx, 4)}
                                onClick={() => handleCellClick(idx, 4)}
                                onContextMenu={(e) => handleCellContextMenu(e, idx, 4)}
                                className={getCellClassName(idx, 4, "border border-black text-right pr-2 font-mono text-[8.5pt] font-semibold text-slate-800 align-top py-0.5 transition-all cursor-pointer")}
                              >
                                <div className="whitespace-normal break-all leading-tight text-[8.5pt]">
                                  {row.amount > 0 ? row.amount.toLocaleString("en-US", { minimumFractionDigits: 2 }) : "0.00"}
                                </div>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Grid Line Actions */}
                <div className="no-print print:hidden flex items-center gap-2 mt-2 mb-3">
                  <button
                    type="button"
                    onClick={addRow}
                    className="flex items-center gap-1.5 bg-indigo-50 hover:bg-indigo-100 border border-indigo-300 text-indigo-800 font-extrabold text-[10px] px-3 py-1.5 rounded-lg shadow-xs transition-all cursor-pointer uppercase tracking-wider"
                  >
                    <Plus className="h-3 w-3" />
                    Add Line
                  </button>
                  <button
                    type="button"
                    onClick={removeRow}
                    disabled={rows.length <= 1}
                    className="flex items-center gap-1.5 bg-rose-50 hover:bg-rose-100 border border-rose-300 text-rose-800 font-extrabold text-[10px] px-3 py-1.5 rounded-lg shadow-xs transition-all cursor-pointer uppercase tracking-wider disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <Trash2 className="h-3 w-3" />
                    Remove Line
                  </button>
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
                                            if (val === "" || /^\d*\.?\d*$/.test(val)) {
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
                                            if (val === "" || /^\d*\.?\d*$/.test(val)) {
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
              insertRow(contextMenu.rowIndex, 'above');
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold"
          >
            Insert Row Above
          </button>
          <button 
            onClick={() => {
              insertRow(contextMenu.rowIndex, 'below');
              setContextMenu(null);
            }}
            className="w-full text-left px-3.5 py-1.5 hover:bg-slate-100 font-bold"
          >
            Insert Row Below
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
    </div>
  );
}
