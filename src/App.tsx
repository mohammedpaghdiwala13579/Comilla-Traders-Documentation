import React, { useState, useEffect } from "react";
import { Download, Monitor, Smartphone, X, Check, Info, Share2, PlusSquare, ArrowUpRight, Laptop } from "lucide-react";
import QuotationBuilder from "./components/QuotationBuilder";

export default function App() {
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [isInstallable, setIsInstallable] = useState(false);
  const [isStandalone, setIsStandalone] = useState(false);
  const [showBanner, setShowBanner] = useState(false);
  const [showGuidance, setShowGuidance] = useState(false);
  const [platform, setPlatform] = useState<"ios" | "android" | "mac" | "windows" | "other">("other");

  useEffect(() => {
    // 1. Check if the application is already running in standalone/installed mode
    const checkStandalone = () => {
      const isStandaloneMode = 
        window.matchMedia('(display-mode: standalone)').matches || 
        (window.navigator as any).standalone === true;
      setIsStandalone(isStandaloneMode);
      
      // Only show banner if NOT standalone and NOT dismissed in this session
      const isDismissed = sessionStorage.getItem("comilla_pwa_dismissed") === "true";
      if (!isStandaloneMode && !isDismissed) {
        setShowBanner(true);
      }
    };

    checkStandalone();

    // 2. Detect device platform for customized PWA experience
    const detectPlatform = () => {
      const ua = navigator.userAgent.toLowerCase();
      if (/iphone|ipad|ipod/.test(ua)) {
        setPlatform("ios");
      } else if (/android/.test(ua)) {
        setPlatform("android");
      } else if (/macintosh|mac os x/.test(ua)) {
        setPlatform("mac");
      } else if (/windows|win32|win64/.test(ua)) {
        setPlatform("windows");
      } else {
        setPlatform("other");
      }
    };

    detectPlatform();

    // 3. Listen for the native browser install prompt event
    const handleBeforeInstallPrompt = (e: Event) => {
      // Prevent standard browser bar from displaying
      e.preventDefault();
      // Store the event so it can be triggered with a single click
      setDeferredPrompt(e);
      setIsInstallable(true);
      
      const isDismissed = sessionStorage.getItem("comilla_pwa_dismissed") === "true";
      if (!isStandalone && !isDismissed) {
        setShowBanner(true);
      }
    };

    window.addEventListener("beforeinstallprompt", handleBeforeInstallPrompt);

    // 4. Listen for successful installation
    const handleAppInstalled = () => {
      setIsStandalone(true);
      setIsInstallable(false);
      setDeferredPrompt(null);
      setShowBanner(false);
      setShowGuidance(false);
      console.log("Comilla Traders Terminal was successfully installed!");
    };

    window.addEventListener("appinstalled", handleAppInstalled);

    return () => {
      window.removeEventListener("beforeinstallprompt", handleBeforeInstallPrompt);
      window.removeEventListener("appinstalled", handleAppInstalled);
    };
  }, [isStandalone]);

  // Single click install trigger
  const handleInstallClick = async () => {
    if (deferredPrompt) {
      // Chrome, Edge, and Android (Chrome) support direct native prompt triggering
      deferredPrompt.prompt();
      const choiceResult = await deferredPrompt.userChoice;
      if (choiceResult.outcome === 'accepted') {
        console.log('User accepted the PWA install prompt');
        setIsStandalone(true);
        setShowBanner(false);
      } else {
        console.log('User dismissed the PWA install prompt');
      }
      setDeferredPrompt(null);
    } else {
      // Fallback/guidance for devices like iOS (Safari), macOS (Safari), or browsers without prompt events
      setShowGuidance(true);
    }
  };

  const dismissBanner = () => {
    setShowBanner(false);
    sessionStorage.setItem("comilla_pwa_dismissed", "true");
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans flex flex-col selection:bg-indigo-500 selection:text-white">
      
      {/* 1. Dynamic PWA Install Banner - Top of the page */}
      {showBanner && !isStandalone && (
        <div id="pwa-install-banner" className="no-print bg-gradient-to-r from-indigo-700 via-indigo-800 to-slate-900 text-white py-2.5 px-4 sm:px-6 relative flex flex-col md:flex-row items-center justify-between gap-3 shadow-md border-b border-indigo-900/40 z-50 animate-in slide-in-from-top duration-300">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600/60 p-1.5 rounded-lg border border-indigo-500/30 flex items-center justify-center shrink-0">
              {platform === "ios" || platform === "android" ? (
                <Smartphone className="h-4 w-4 text-indigo-200 animate-pulse" />
              ) : (
                <Laptop className="h-4 w-4 text-indigo-200 animate-pulse" />
              )}
            </div>
            <div>
              <p className="text-[12px] font-semibold tracking-wide flex items-center gap-1.5">
                INSTALL APP TERMINAL
                <span className="bg-indigo-500/50 text-indigo-200 text-[9px] font-bold px-1.5 py-0.5 rounded uppercase tracking-wider">Fast &amp; Secure</span>
              </p>
              <p className="text-[10px] text-indigo-100/90 font-medium">
                Install on your {platform === "ios" ? "iPhone/iPad" : platform === "mac" ? "Macbook" : platform === "windows" ? "Windows PC" : platform === "android" ? "Android Phone" : "Device"} for standalone double-click access and offline capability.
              </p>
            </div>
          </div>
          
          <div className="flex items-center gap-2 shrink-0">
            <button
              id="btn-banner-install"
              onClick={handleInstallClick}
              className="bg-white hover:bg-slate-100 text-indigo-900 p-2 rounded-md flex items-center justify-center shadow transition-all transform hover:scale-[1.02] active:scale-[0.98] cursor-pointer shrink-0"
              title="Install Now (Single Click)"
            >
              <Download className="h-4 w-4" />
            </button>
            <button
              id="btn-banner-dismiss"
              onClick={dismissBanner}
              className="text-indigo-200 hover:text-white p-1 hover:bg-white/10 rounded transition-colors cursor-pointer"
              title="Dismiss"
            >
              <X className="h-4 w-4" />
            </button>
          </div>
        </div>
      )}

      {/* 2. Top Install Action */}
      {!isStandalone && (
        <header className="no-print bg-white/80 backdrop-blur-xs border-b border-slate-200/60 py-2 px-4 sm:px-6 flex justify-end items-center shadow-xs z-10">
          <button
            id="btn-header-install"
            onClick={handleInstallClick}
            className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 hover:text-indigo-800 border border-indigo-200 text-xs font-bold px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition-all shadow-sm cursor-pointer"
          >
            <Download className="h-3.5 w-3.5 text-indigo-600 animate-bounce" />
            <span>Install App</span>
          </button>
        </header>
      )}

      {/* 3. iOS / Safari / Manual Install Guidance Modal */}
      {showGuidance && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-in fade-in duration-200">
          <div className="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-md w-full p-6 relative animate-in zoom-in-95 duration-200">
            <button
              id="btn-modal-close"
              onClick={() => setShowGuidance(false)}
              className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 p-1 rounded-full hover:bg-slate-100 transition-colors cursor-pointer"
            >
              <X className="h-5 w-5" />
            </button>

            <div className="flex items-center gap-3 mb-4">
              <div className="h-10 w-10 rounded-full bg-indigo-50 flex items-center justify-center text-indigo-600 shrink-0">
                <Info className="h-5 w-5" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-slate-900">How to Install Application</h3>
                <p className="text-[11px] text-slate-500">Enable single-click desktop/mobile launching</p>
              </div>
            </div>

            <div className="space-y-4 my-4">
              {/* iOS Safari Guidance */}
              {platform === "ios" ? (
                <div className="bg-indigo-50/55 p-3.5 rounded-lg border border-indigo-100 text-xs space-y-2.5">
                  <p className="font-semibold text-indigo-950 flex items-center gap-1.5">
                    <Smartphone className="h-4 w-4 text-indigo-600" />
                    For iPhone &amp; iPad (Safari):
                  </p>
                  <ol className="list-decimal pl-4 space-y-2 text-slate-700">
                    <li>
                      Tap the <strong className="text-slate-900">Share</strong> button ( <Share2 className="h-3 w-3 inline-block text-indigo-600" /> icon) at the bottom or top of Safari.
                    </li>
                    <li>
                      Scroll down and select <strong className="text-slate-900">Add to Home Screen</strong> ( <PlusSquare className="h-3 w-3 inline-block text-indigo-600" /> icon).
                    </li>
                    <li>
                      Tap <strong className="text-slate-900">Add</strong> in the top-right corner.
                    </li>
                  </ol>
                  <p className="text-[10px] text-slate-500 italic mt-1">Note: This is Apple's standard requirement for web applications.</p>
                </div>
              ) : platform === "mac" ? (
                /* macOS Safari Guidance */
                <div className="bg-indigo-50/55 p-3.5 rounded-lg border border-indigo-100 text-xs space-y-2.5">
                  <p className="font-semibold text-indigo-950 flex items-center gap-1.5">
                    <Laptop className="h-4 w-4 text-indigo-600" />
                    For macOS (Safari):
                  </p>
                  <ol className="list-decimal pl-4 space-y-2 text-slate-700">
                    <li>
                      Go to the top browser menu and click <strong className="text-slate-900">File</strong>.
                    </li>
                    <li>
                      Select <strong className="text-slate-900">Add to Dock...</strong> from the dropdown menu.
                    </li>
                    <li>
                      Confirm the name and click <strong className="text-slate-900">Add</strong>.
                    </li>
                  </ol>
                  <p className="text-[10px] text-slate-500 italic mt-1">This will place a beautiful Comilla Traders icon right in your Mac Dock!</p>
                </div>
              ) : (
                /* PC / Other Browsers manual fallback instruction */
                <div className="bg-indigo-50/55 p-3.5 rounded-lg border border-indigo-100 text-xs space-y-2.5">
                  <p className="font-semibold text-indigo-950 flex items-center gap-1.5">
                    <Monitor className="h-4 w-4 text-indigo-600" />
                    Standard Browser Installation:
                  </p>
                  <ol className="list-decimal pl-4 space-y-2 text-slate-700">
                    <li>
                      Look at the right side of your browser's address bar at the top.
                    </li>
                    <li>
                      Click the <strong className="text-slate-900">Install app</strong> icon (usually a small monitor with a down arrow, or a plus symbol).
                    </li>
                    <li>
                      If not found, open the browser menu ( <strong className="text-slate-900">three dots ⋮</strong> or <strong className="text-slate-900">hamburger menu ☰</strong> ) and click <strong className="text-slate-900">Save and share</strong> ➡️ <strong className="text-slate-900">Install page...</strong>.
                    </li>
                  </ol>
                </div>
              )}

              {/* Universal Benefits */}
              <div className="bg-slate-50 p-3 rounded-lg border border-slate-200 text-[11px] text-slate-600 space-y-1.5">
                <p className="font-bold text-slate-800">Benefits of Installing:</p>
                <div className="grid grid-cols-2 gap-2 text-[10px]">
                  <div className="flex items-center gap-1">
                    <Check className="h-3 w-3 text-emerald-600" /> Offline Access Ready
                  </div>
                  <div className="flex items-center gap-1">
                    <Check className="h-3 w-3 text-emerald-600" /> Zero Browser Margins
                  </div>
                  <div className="flex items-center gap-1">
                    <Check className="h-3 w-3 text-emerald-600" /> Desktop/Dock Shortcut
                  </div>
                  <div className="flex items-center gap-1">
                    <Check className="h-3 w-3 text-emerald-600" /> Fluid App Window
                  </div>
                </div>
              </div>
            </div>

            <button
              id="btn-guidance-ok"
              onClick={() => setShowGuidance(false)}
              className="w-full bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold py-2 rounded-lg transition-colors cursor-pointer mt-4"
            >
              Got It
            </button>
          </div>
        </div>
      )}

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
