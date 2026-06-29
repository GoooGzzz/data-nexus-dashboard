"use client";

import dynamic from "next/dynamic";

// Dynamically import the main dashboard wrapper with SSR completely disabled.
// This fully isolates Leaflet, Recharts, and browser globals from Node.js during compilation.
const VisitCompliancePage = dynamic(
  () => import("./visit-compliance/VisitCompliancePage"),
  { 
    ssr: false,
    loading: () => (
      <div className="flex h-screen w-screen items-center justify-center bg-neutral-950">
        <div className="flex flex-col items-center gap-3">
          <div className="w-8 h-8 border-4 border-cyan-500 border-t-transparent rounded-full animate-spin" />
          <p className="text-sm font-medium text-neutral-400 font-mono tracking-wider">LOADING DATA NEXUS...</p>
        </div>
      </div>
    )
  }
);

export default function Home() {
  return <VisitCompliancePage />;
}