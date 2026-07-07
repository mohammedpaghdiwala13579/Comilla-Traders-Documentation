import React from "react";
import QuotationBuilder from "./components/QuotationBuilder";

export default function App() {
  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans flex flex-col selection:bg-indigo-500 selection:text-white">
      {/* Portal Top Header banner - Non-printing */}
      <header className="no-print bg-white text-slate-800 border-b border-slate-200 py-3.5 px-6 flex flex-col sm:flex-row justify-between items-center gap-4 shadow-sm z-10">
        <div className="flex items-center gap-3">
          <div className="h-10 w-10 rounded-lg border border-slate-200 overflow-hidden bg-white flex items-center justify-center shrink-0 p-0.5 shadow-sm">
            <img
              src="https://i.ibb.co.com/gFBkpt8B/Chat-GPT-Image-Apr-23-2026-01-10-13-PM.png"
              alt="Comilla Traders Logo"
              className="w-full h-full object-cover rounded"
            />
          </div>
          <div>
            <h1 className="text-sm font-bold tracking-tight text-slate-900 flex items-center gap-1.5 uppercase font-display">
              COMILLA TRADERS
              <span className="bg-indigo-50 text-indigo-700 text-[9px] font-semibold px-1.5 py-0.5 rounded border border-indigo-200">QUOTATION TERMINAL</span>
            </h1>
            <p className="text-[10px] text-slate-500 font-medium tracking-wide">Professional Ship Chandler Quotation &amp; Delivery Challan Terminal</p>
          </div>
        </div>
      </header>

      {/* Main viewport */}
      <main className="flex-1 flex flex-col p-4 sm:p-6 max-w-7xl mx-auto w-full">
        <div className="animate-in fade-in duration-300 w-full">
          <QuotationBuilder />
        </div>
      </main>

      {/* Dynamic Professional Status Bar Footer */}
      <footer className="no-print bg-white text-slate-400 py-4 px-6 border-t border-slate-200 flex flex-col md:flex-row justify-between items-center gap-4 text-xs font-medium mt-auto">
        <div className="flex items-center gap-5 flex-wrap justify-center">
          <div className="flex items-center gap-2">
            <div className="w-1.5 h-1.5 rounded-full bg-emerald-500"></div>
            <span className="text-[10px] font-mono tracking-wide text-slate-600 uppercase font-semibold">FIREBASE_CONNECTED</span>
          </div>
          <div className="hidden sm:flex items-center gap-2 border-l border-slate-200 pl-5">
            <span className="text-[10px] text-slate-400 font-mono">DB: ai-studio-2c592343-56ab-4d40-a2ac-d15fed703e91</span>
          </div>
        </div>
        <div className="text-[10px] text-slate-400 uppercase tracking-wider text-center md:text-right font-semibold">
          &copy; {new Date().getFullYear()} Comilla Traders &bull; Ship Chandler Portal &bull; EnterprisePro Engine v4.2
        </div>
      </footer>
    </div>
  );
}
