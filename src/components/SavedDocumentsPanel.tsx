import React, { useState } from "react";
import { History, Search, X, FileText, Trash2, HelpCircle, BookOpen, Layers } from "lucide-react";
import { SavedDocument } from "../types";

interface SavedDocumentsPanelProps {
  savedDocs: SavedDocument[];
  currentDocId: string | null;
  searchQuery: string;
  setSearchQuery: (val: string) => void;
  selectedTypeFilter: "all" | "quotation" | "challan" | "invoice";
  setSelectedTypeFilter: (val: "all" | "quotation" | "challan" | "invoice") => void;
  loadSavedDoc: (doc: SavedDocument) => void;
  deleteSavedDoc: (id: string, e?: React.MouseEvent) => void;
  renameSavedDoc: (id: string, e: React.MouseEvent) => void;
}

export default function SavedDocumentsPanel({
  savedDocs,
  currentDocId,
  searchQuery,
  setSearchQuery,
  selectedTypeFilter,
  setSelectedTypeFilter,
  loadSavedDoc,
  deleteSavedDoc,
  renameSavedDoc,
}: SavedDocumentsPanelProps) {
  const [showDocHelp, setShowDocHelp] = useState(false);

  // Filter documents in client
  const filteredDocs = savedDocs.filter((doc) => {
    if (selectedTypeFilter !== "all" && doc.docType !== selectedTypeFilter) {
      return false;
    }
    if (searchQuery.trim() !== "") {
      const q = searchQuery.toLowerCase().trim();
      const nameMatch = doc.name?.toLowerCase().includes(q);
      const messersMatch = doc.messers?.toLowerCase().includes(q);
      const addressMatch = doc.address?.toLowerCase().includes(q);
      const challanNoMatch = doc.challanNo?.toLowerCase().includes(q);
      const requisitionNoMatch = doc.requisitionNo?.toLowerCase().includes(q);
      const invoiceNoMatch = doc.invoiceNo?.toLowerCase().includes(q);
      const poNumberMatch = doc.poNumber?.toLowerCase().includes(q);
      const rowMatch = doc.rows?.some(r => r.desc?.toLowerCase().includes(q));
      return nameMatch || messersMatch || addressMatch || challanNoMatch || requisitionNoMatch || invoiceNoMatch || poNumberMatch || rowMatch;
    }
    return true;
  });

  return (
    <div className="saved-docs-panel no-print print:hidden w-full bg-white rounded-2xl border border-slate-200 shadow-sm p-4 sm:p-6 text-black mt-6 animate-in fade-in slide-in-from-bottom-3 duration-300">
      {/* Panel Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4 mb-4">
        <div className="flex items-start gap-3">
          <div className="p-2 bg-indigo-50 text-indigo-600 rounded-lg shrink-0">
            <History className="h-5 w-5" />
          </div>
          <div>
            <h2 className="text-base font-bold text-slate-900 tracking-tight flex items-center gap-2">
              <span>Online Cloud Storage</span>
              <span className="bg-emerald-50 border border-emerald-100 text-emerald-700 text-[10px] font-bold px-2.5 py-0.5 rounded-full uppercase tracking-wider">
                Firestore Connected ({savedDocs.length})
              </span>
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Access your saved sheets instantly from any device.
            </p>
          </div>
        </div>

        {/* Action Controls */}
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={() => setShowDocHelp(!showDocHelp)}
            className="flex items-center gap-1.5 bg-white hover:bg-slate-50 text-slate-700 font-bold text-[10px] px-3 py-1.5 rounded-lg border border-slate-200 shadow-3xs uppercase tracking-wider transition-all cursor-pointer"
          >
            <BookOpen className="h-3.5 w-3.5" />
            <span>{showDocHelp ? "Hide Docs" : "Portal Guide"}</span>
          </button>
        </div>
      </div>

      {/* Embedded Documentation Accordion */}
      {showDocHelp && (
        <div className="bg-slate-50 border border-indigo-100/60 rounded-xl p-4 mb-5 text-xs text-slate-700 space-y-3 animate-in fade-in slide-in-from-top-2 duration-200">
          <h3 className="font-extrabold text-xs text-indigo-950 uppercase tracking-wider flex items-center gap-1.5">
            <HelpCircle className="h-4 w-4 text-indigo-600" />
            System Features &amp; User Documentation
          </h3>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-1">
            <div className="bg-white p-3 rounded-lg border border-slate-100 shadow-3xs space-y-1">
              <h4 className="font-bold text-slate-900 flex items-center gap-1">
                <span className="text-indigo-600 font-mono">1.</span> Persistent Cloud Storage
              </h4>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                All changes are securely synchronized to the cloud. You can use the <strong>SAVE</strong> button to upload on demand. <strong>Auto-Save</strong> works silently in the background, updating your edits 1.5s after typing.
              </p>
            </div>

            <div className="bg-white p-3 rounded-lg border border-slate-100 shadow-3xs space-y-1">
              <h4 className="font-bold text-slate-900 flex items-center gap-1">
                <span className="text-indigo-600 font-mono">2.</span> Excel Copy &amp; Paste
              </h4>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                You can copy a table range directly from Microsoft Excel or Google Sheets, click any cell in our editor grid, and hit paste (Ctrl+V). Multiple cells, descriptions, units, and quantities will autofill instantly.
              </p>
            </div>

            <div className="bg-white p-3 rounded-lg border border-slate-100 shadow-3xs space-y-1">
              <h4 className="font-bold text-slate-900 flex items-center gap-1">
                <span className="text-indigo-600 font-mono">3.</span> Dynamic Cell Merging
              </h4>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                Select a block of cells in the grid by holding the mouse down and dragging across cells. Click <strong>Merge</strong> in the toolbar (or right-click) to combine them into a single cell. This merged layout is saved in the database!
              </p>
            </div>

            <div className="bg-white p-3 rounded-lg border border-slate-100 shadow-3xs space-y-1">
              <h4 className="font-bold text-slate-900 flex items-center gap-1">
                <span className="text-indigo-600 font-mono">4.</span> Repeating Letterhead Printing
              </h4>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                Our print style engine uses native HTML table structures to format long sheets. When printing to paper or PDF, the <strong>Comilla Traders letterhead automatically repeats</strong> at the top of every page.
              </p>
            </div>

            <div className="bg-white p-3 rounded-lg border border-slate-100 shadow-3xs space-y-1 sm:col-span-2">
              <h4 className="font-bold text-slate-900 flex items-center gap-1">
                <span className="text-indigo-600 font-mono">5.</span> Interactive Excel Downloads
              </h4>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                Our advanced spreadsheet generator creates <code>.xlsx</code> sheets with exact page-by-page visual fidelity (logos, double lines, dotted backgrounds, signature lines) but preserves <strong>fully editable formula structures</strong> (sums, multiplications) and transfers your custom merged cells directly into Excel.
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Search Bar & Type Filters */}
      <div className="space-y-3 mb-5">
        <div className="relative w-full">
          <Search className="absolute left-3.5 top-3 h-4 w-4 text-slate-400" />
          <input
            type="text"
            placeholder="Search documents by client name, address, Challan No., or description items..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full pl-10 pr-10 py-2.5 text-xs bg-slate-50 hover:bg-slate-100/70 focus:bg-white border border-slate-200 rounded-xl text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 shadow-xs transition-all"
          />
          {searchQuery && (
            <button
              type="button"
              onClick={() => setSearchQuery("")}
              className="absolute right-3 top-3 text-slate-400 hover:text-slate-600 focus:outline-none cursor-pointer"
            >
              <X className="h-4 w-4" />
            </button>
          )}
        </div>

        {savedDocs.length > 0 && (
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-50 p-2.5 rounded-xl border border-slate-100">
            <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-400 pl-1 flex items-center gap-1.5">
              <Layers className="h-3.5 w-3.5 text-slate-400" />
              Filter List:
            </span>
            <div className="flex flex-wrap items-center gap-1 bg-slate-100 p-1 rounded-lg border border-slate-200/40">
              {(["all", "quotation", "challan", "invoice"] as const).map((type) => {
                const count = type === "all"
                  ? savedDocs.length
                  : savedDocs.filter((d) => d.docType === type).length;
                return (
                  <button
                    key={type}
                    type="button"
                    onClick={() => setSelectedTypeFilter(type)}
                    className={`px-3 py-1 rounded-md text-[10px] font-extrabold uppercase tracking-wider transition-all cursor-pointer ${
                      selectedTypeFilter === type
                        ? "bg-white text-indigo-700 shadow-sm border border-slate-200/60"
                        : "text-slate-600 hover:text-slate-900"
                    }`}
                  >
                    {type === "all" ? "All Files" : type} ({count})
                  </button>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* Documents Table */}
      {(() => {
        if (savedDocs.length === 0) {
          return (
            <div className="flex flex-col items-center justify-center py-10 px-4 text-center border-2 border-dashed border-slate-200 rounded-xl bg-slate-50/50">
              <FileText className="h-10 w-10 text-slate-300 stroke-[1.5]" />
              <h3 className="font-bold text-xs text-slate-700 uppercase tracking-wider mt-3">No Saved Documents Yet</h3>
              <p className="text-xs text-slate-500 mt-1 max-w-sm">
                Saved files will appear here. Create content and click <strong className="text-indigo-600">"SAVE"</strong> to store sheets.
              </p>
            </div>
          );
        }

        if (filteredDocs.length === 0) {
          return (
            <div className="flex flex-col items-center justify-center py-10 px-4 text-center border border-dashed border-slate-200 rounded-xl bg-slate-50/50">
              <Search className="h-10 w-10 text-slate-300 stroke-[1.5]" />
              <h3 className="font-bold text-xs text-slate-700 uppercase tracking-wider mt-3">No Matching Documents</h3>
              <p className="text-xs text-slate-500 mt-1 max-w-sm">
                No documents match the search string or filter selection.
              </p>
              <button
                type="button"
                onClick={() => {
                  setSearchQuery("");
                  setSelectedTypeFilter("all");
                }}
                className="mt-4 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-[10px] py-1.5 px-3 rounded-lg border border-slate-300 uppercase tracking-wider transition-all cursor-pointer"
              >
                Clear Filters
              </button>
            </div>
          );
        }

        return (
          <div className="overflow-x-auto rounded-lg border border-slate-200">
            <table className="w-full text-left text-xs text-slate-600 border-collapse">
              <thead className="bg-slate-50 text-[10px] font-bold text-slate-500 uppercase tracking-wider border-b border-slate-200">
                <tr>
                  <th className="px-4 py-3">Document Name</th>
                  <th className="px-4 py-3 text-center w-24">Type</th>
                  <th className="px-4 py-3 text-center w-24">Date</th>
                  <th className="px-4 py-3">Last Updated</th>
                  <th className="px-4 py-3 text-right pr-6">Grand Total</th>
                  <th className="px-4 py-3 text-right pr-4 w-48">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white">
                {filteredDocs.map((doc) => {
                  const isActive = currentDocId === doc.id;
                  const docRowsTotal = doc.rows.reduce((sum: number, r: any) => sum + (Number(r.amount) || 0), 0);
                  const vat = doc.docType === "invoice" ? (docRowsTotal * (Number(doc.vatPercent) || 0)) / 100 : 0;
                  const trans = doc.docType === "invoice" ? (Number(doc.transportationFee) || 0) : 0;
                  const grand = docRowsTotal + vat + trans;
                  
                  // For Challan, grand total is shown as "—"
                  const displayTotal = doc.docType === "challan"
                    ? "—"
                    : grand.toLocaleString("en-US", { minimumFractionDigits: 2 });
                  
                  return (
                    <tr 
                      key={doc.id} 
                      onClick={() => loadSavedDoc(doc)}
                      className={`hover:bg-slate-50/80 transition-colors cursor-pointer group ${
                        isActive ? "bg-indigo-50/30 hover:bg-indigo-50/40" : ""
                      }`}
                    >
                      <td className="px-4 py-3 font-medium text-slate-800">
                        <div className="flex items-center gap-2">
                          {isActive && (
                            <span className="inline-flex items-center bg-indigo-100 text-indigo-800 text-[9px] font-bold px-1.5 py-0.2 rounded-sm tracking-wide">
                              EDITING
                            </span>
                          )}
                          <span className="truncate max-w-[220px] sm:max-w-[320px] block font-semibold text-slate-900" title={doc.name}>
                            {doc.name}
                          </span>
                        </div>
                      </td>
                      <td className="px-4 py-3 text-center">
                        <span className={`inline-flex items-center rounded-md px-2 py-0.5 text-[10px] font-extrabold tracking-wider uppercase border ${
                          doc.docType === "quotation"
                            ? "bg-blue-50 border-blue-200 text-blue-700"
                            : doc.docType === "challan"
                            ? "bg-emerald-50 border-emerald-200 text-emerald-700"
                            : "bg-indigo-50 border-indigo-200 text-indigo-700"
                        }`}>
                          {doc.docType}
                        </span>
                      </td>
                      <td className="px-4 py-3 text-center text-slate-500 font-mono">{doc.dateVal}</td>
                      <td className="px-4 py-3 text-slate-400 font-mono">
                        {new Date(doc.updatedAt).toLocaleDateString()} {new Date(doc.updatedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                      </td>
                      <td className="px-4 py-3 text-right pr-6 font-mono font-bold text-slate-900">
                        {displayTotal}
                      </td>
                      <td className="px-4 py-3 text-right pr-4 space-x-1.5 no-print" onClick={(e) => e.stopPropagation()}>
                        <button
                          type="button"
                          onClick={(e) => renameSavedDoc(doc.id, e)}
                          className="text-[10px] font-bold text-slate-500 hover:text-slate-800 hover:bg-slate-100 px-2.5 py-1 rounded transition-colors cursor-pointer"
                        >
                          Rename
                        </button>
                        <button
                          type="button"
                          onClick={(e) => deleteSavedDoc(doc.id, e)}
                          className="text-[10px] font-bold text-rose-500 hover:text-rose-700 hover:bg-rose-50 px-2.5 py-1 rounded transition-colors cursor-pointer"
                        >
                          Delete
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        );
      })()}
    </div>
  );
}
