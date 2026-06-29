"use client";
import { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Upload, FileSpreadsheet, CheckCircle, AlertTriangle, XCircle,
  Users, ShieldCheck, ShieldAlert, TrendingDown, TrendingUp,
  Layers, RefreshCw, Crown, Flag,
} from "lucide-react";

// ─── TYPES ──────────────────────────────────────────────────────────────────
type VisitFileKind = "visit_raw" | "visit_followup" | "visit_summary" | "unknown";

interface LoadedFile {
  name: string;
  kind: VisitFileKind;
  size: string;
}

interface EmployeeRecord {
  code: string;
  name: string;
  department: string;
  daysReported: number;
  daysPossible: number;
  missingDates: string[];
  totalReportsFollowUp: number | null;
  inRaw: boolean;
  inFollowUp: boolean;
  inSummary: boolean;
}

interface AuditIssue {
  severity: "high" | "medium" | "low";
  message: string;
}

const fmtBytes = (b: number) => (b > 1_048_576 ? `${(b / 1_048_576).toFixed(1)} MB` : `${(b / 1024).toFixed(0)} KB`);
const pct = (a: number, b: number) => (b > 0 ? Math.round((a / b) * 100) : 0);

function detectKind(sheetNames: string[]): VisitFileKind {
  if (sheetNames.includes("Mobile")) return "visit_raw";
  if (sheetNames.includes("Submission Matrix")) return "visit_followup";
  if (sheetNames.includes("Weekly Summary")) return "visit_summary";
  return "unknown";
}

// Parse "Weekly Summary" sheet -> per-employee reported/possible days, by Department
function parseSummary(rows: unknown[][]): EmployeeRecord[] {
  const headerRow = (rows[1] || []) as unknown[]; // day-level header row (weekday names / Total / Notes)
  const dayCols: number[] = [];
  for (let c = 3; c < headerRow.length; c++) {
    const h = String(headerRow[c] ?? "");
    if (h && !h.includes("Total") && !h.toLowerCase().includes("notes")) dayCols.push(c);
  }
  const out: EmployeeRecord[] = [];
  for (let r = 2; r < rows.length; r++) {
    const row = rows[r] as unknown[];
    const name = row?.[0];
    const code = row?.[1];
    const dept = row?.[2];
    if (!name || !code) continue;
    let reported = 0;
    for (const c of dayCols) {
      const v = row[c];
      if (v !== null && v !== undefined && v !== "" && Number(v) > 0) reported++;
    }
    out.push({
      code: String(code), name: String(name), department: String(dept ?? "Unassigned"),
      daysReported: reported, daysPossible: dayCols.length, missingDates: [],
      totalReportsFollowUp: null, inRaw: false, inFollowUp: false, inSummary: true,
    });
  }
  return out;
}

// Parse "Submission Matrix" sheet -> Missing Dates + Total Reports per employee
function parseFollowUp(rows: unknown[][]): Map<string, { total: number; missing: string[] }> {
  const map = new Map<string, { total: number; missing: string[] }>();
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] as unknown[];
    const code = row?.[1];
    if (!code) continue;
    const total = Number(row?.[2]) || 0;
    const missingStr = String(row?.[3] ?? "");
    const missing = missingStr === "–" || !missingStr ? [] : missingStr.split(",").map(s => s.trim()).filter(Boolean);
    map.set(String(code), { total, missing });
  }
  return map;
}

// Parse "Mobile" raw sheet -> distinct AM-SPVR codes present (submitter codes, col index 1)
function parseRawCodes(rows: unknown[][]): Set<string> {
  const set = new Set<string>();
  for (let r = 3; r < rows.length; r++) {
    const row = rows[r] as unknown[];
    const code = row?.[1];
    if (code) set.add(String(code));
  }
  return set;
}

function buildAudit(employees: EmployeeRecord[]): AuditIssue[] {
  const issues: AuditIssue[] = [];
  const seen = new Set<string>();
  for (const e of employees) {
    if (seen.has(e.code)) issues.push({ severity: "high", message: `Duplicate employee code "${e.code}" (${e.name}) appears more than once in Summary.` });
    seen.add(e.code);
    if (!e.inRaw) issues.push({ severity: "medium", message: `${e.name} (${e.code}) has Summary/Follow-Up data but zero raw visit entries.` });
    if (!e.inFollowUp) issues.push({ severity: "medium", message: `${e.name} (${e.code}) is missing from the Follow-Up Submission Matrix.` });
    if (e.daysPossible > 0 && e.daysReported === 0) issues.push({ severity: "high", message: `${e.name} (${e.code}) has zero reported days for the entire period — possible ghost employee or total non-compliance.` });
  }
  return issues;
}

function complianceColor(p: number) {
  if (p >= 80) return "text-emerald-400";
  if (p >= 50) return "text-amber-400";
  return "text-red-400";
}

// ─── COMPONENT ──────────────────────────────────────────────────────────────
export default function VisitCompliancePage() {
  const [files, setFiles] = useState<LoadedFile[]>([]);
  const [employees, setEmployees] = useState<EmployeeRecord[]>([]);
  const [followUpMap, setFollowUpMap] = useState<Map<string, { total: number; missing: string[] }>>(new Map());
  const [rawCodes, setRawCodes] = useState<Set<string>>(new Set());
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [deptFilter, setDeptFilter] = useState("All");

  const handleFiles = useCallback(async (fileList: FileList) => {
    setIsProcessing(true);
    const newLoaded: LoadedFile[] = [];
    let newEmployees: EmployeeRecord[] | null = null;
    let newFollowUp: Map<string, { total: number; missing: string[] }> | null = null;
    let newRaw: Set<string> | null = null;

    for (const file of Array.from(fileList)) {
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const kind = detectKind(wb.SheetNames);
        newLoaded.push({ name: file.name, kind, size: fmtBytes(file.size) });

        if (kind === "visit_summary") {
          const ws = wb.Sheets["Weekly Summary"];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][];
          newEmployees = parseSummary(rows);
        } else if (kind === "visit_followup") {
          const ws = wb.Sheets["Submission Matrix"];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][];
          newFollowUp = parseFollowUp(rows);
        } else if (kind === "visit_raw") {
          const ws = wb.Sheets["Mobile"];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][];
          newRaw = parseRawCodes(rows);
        }
      } catch (err) {
        newLoaded.push({ name: file.name, kind: "unknown", size: "—" });
      }
    }

    setFiles(prev => {
      const existing = new Set(prev.map(f => f.name));
      return [...prev, ...newLoaded.filter(f => !existing.has(f.name))];
    });
    if (newFollowUp) setFollowUpMap(newFollowUp);
    if (newRaw) setRawCodes(newRaw);
    if (newEmployees) setEmployees(newEmployees);
    setIsProcessing(false);
  }, []);

  // Cross-reference once we have employees + follow-up + raw
  const enriched: EmployeeRecord[] = employees.map(e => {
    const fu = followUpMap.get(e.code);
    return {
      ...e,
      totalReportsFollowUp: fu ? fu.total : null,
      missingDates: fu ? fu.missing : [],
      inFollowUp: !!fu,
      inRaw: rawCodes.has(e.code),
    };
  });

  const audit = buildAudit(enriched);
  const departments = ["All", ...Array.from(new Set(enriched.map(e => e.department)))];
  const filtered = deptFilter === "All" ? enriched : enriched.filter(e => e.department === deptFilter);

  const deptRollup = Array.from(new Set(enriched.map(e => e.department))).map(dept => {
    const members = enriched.filter(e => e.department === dept);
    const totalReported = members.reduce((s, m) => s + m.daysReported, 0);
    const totalPossible = members.reduce((s, m) => s + m.daysPossible, 0);
    const compliance = pct(totalReported, totalPossible);
    const chronic = members.filter(m => pct(m.daysReported, m.daysPossible) < 50).length;
    return { dept, headcount: members.length, compliance, chronic };
  }).sort((a, b) => b.compliance - a.compliance);

  const overallCompliance = pct(
    enriched.reduce((s, e) => s + e.daysReported, 0),
    enriched.reduce((s, e) => s + e.daysPossible, 0)
  );
  const chronicOffenders = enriched
    .map(e => ({ ...e, compliance: pct(e.daysReported, e.daysPossible) }))
    .filter(e => e.compliance < 50)
    .sort((a, b) => a.compliance - b.compliance);

  const hasData = enriched.length > 0;

  return (
    <div className="p-6 space-y-6 max-w-full">
      {/* Header */}
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <h2 className="text-xl font-bold text-white tracking-wider">DAILY VISIT COMPLIANCE &amp; AUDIT</h2>
          <p className="text-xs text-neutral-500 mt-0.5">
            Upload Raw / Follow-Up / Summary reports for real-time KPI rollup across supervisors &amp; managers
          </p>
        </div>
        {files.length > 0 && (
          <div className="flex items-center gap-2">
            <span className="text-xs text-neutral-500 bg-neutral-800 border border-neutral-700 px-3 py-1.5 rounded-lg">
              {files.length} file{files.length > 1 ? "s" : ""} loaded
            </span>
            <button
              onClick={() => { setFiles([]); setEmployees([]); setFollowUpMap(new Map()); setRawCodes(new Set()); }}
              className="text-xs text-red-400 bg-red-500/10 border border-red-500/30 px-3 py-1.5 rounded-lg hover:bg-red-500/20 transition-colors"
            >
              Clear all
            </button>
          </div>
        )}
      </div>

      {/* Upload zone */}
      <div
        onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={(e) => { e.preventDefault(); setIsDragging(false); if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files); }}
        className={`relative border-2 border-dashed rounded-xl p-8 text-center transition-all ${
          isDragging ? "border-cyan-500 bg-cyan-500/5" : "border-neutral-700 hover:border-neutral-600 bg-neutral-900"
        }`}
      >
        <input
          type="file" accept=".xlsx,.xls" multiple
          onChange={(e) => { if (e.target.files?.length) handleFiles(e.target.files); }}
          className="absolute inset-0 opacity-0 cursor-pointer"
        />
        <Upload className={`w-8 h-8 mx-auto mb-3 ${isDragging ? "text-cyan-400" : "text-neutral-500"}`} />
        <p className="text-sm font-medium text-white">Drop Raw / Follow-Up / Summary .xlsx files here, or click to browse</p>
        <p className="text-xs text-neutral-500 mt-1">
          Auto-detected by sheet name: <span className="text-cyan-400 font-mono">Mobile</span> ·{" "}
          <span className="text-purple-400 font-mono">Submission Matrix</span> ·{" "}
          <span className="text-emerald-400 font-mono">Weekly Summary</span>
        </p>
        {isProcessing && (
          <div className="mt-3 flex items-center justify-center gap-2 text-xs text-cyan-400">
            <RefreshCw className="w-3.5 h-3.5 animate-spin" /> Parsing workbook(s)…
          </div>
        )}
      </div>

      {/* File pills */}
      {files.length > 0 && (
        <div className="flex flex-wrap gap-2">
          {files.map(f => (
            <div key={f.name} className={`flex items-center gap-2 px-4 py-2 rounded-xl border text-xs font-medium
              ${f.kind === "unknown" ? "bg-neutral-900 border-amber-500/40 text-amber-300" : "bg-neutral-900 border-neutral-700 text-neutral-300"}`}>
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span className="max-w-[200px] truncate">{f.name}</span>
              <span className="text-neutral-600">{f.size}</span>
              {f.kind === "unknown"
                ? <span className="flex items-center gap-1 text-amber-400"><AlertTriangle className="w-3 h-3" /> unrecognised</span>
                : <span className="flex items-center gap-1 text-emerald-400"><CheckCircle className="w-3 h-3" /> {f.kind.replace("visit_", "")}</span>}
            </div>
          ))}
        </div>
      )}

      {!hasData && files.length === 0 && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {[
            { icon: FileSpreadsheet, color: "cyan", title: "Raw Visit Log", desc: "AM-SPVR daily market visits per shop (Mobile sheet)" },
            { icon: ShieldCheck, color: "purple", title: "Follow-Up Matrix", desc: "Per-employee submission times, missing-date audit" },
            { icon: Layers, color: "emerald", title: "Weekly Summary", desc: "Employee × Department rollup, daily reported flag" },
          ].map((c, i) => (
            <div key={i} className={`bg-neutral-900 border rounded-xl p-5 border-${c.color}-500/30 bg-${c.color}-500/5`}>
              <c.icon className={`w-6 h-6 mb-2 text-${c.color}-400`} />
              <p className="text-white font-semibold mb-1">{c.title}</p>
              <p className="text-xs text-neutral-500 leading-relaxed">{c.desc}</p>
            </div>
          ))}
        </div>
      )}

      {hasData && (
        <>
          {/* KPI cards */}
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-4">
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs text-neutral-400">TOTAL HEADCOUNT</span>
                <Users className="w-4 h-4 text-cyan-400" />
              </div>
              <p className="text-2xl font-bold text-white font-mono">{enriched.length}</p>
            </div>
            <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-4">
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs text-neutral-400">OVERALL COMPLIANCE</span>
                {overallCompliance >= 70 ? <TrendingUp className="w-4 h-4 text-emerald-400" /> : <TrendingDown className="w-4 h-4 text-red-400" />}
              </div>
              <p className={`text-2xl font-bold font-mono ${complianceColor(overallCompliance)}`}>{overallCompliance}%</p>
            </div>
            <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-4">
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs text-neutral-400">CHRONIC NON-COMPLIANT</span>
                <ShieldAlert className="w-4 h-4 text-red-400" />
              </div>
              <p className="text-2xl font-bold text-red-400 font-mono">{chronicOffenders.length}</p>
              <p className="text-[10px] text-neutral-600 mt-0.5">&lt;50% of days reported</p>
            </div>
            <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-4">
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs text-neutral-400">AUDIT FLAGS</span>
                <Flag className="w-4 h-4 text-amber-400" />
              </div>
              <p className="text-2xl font-bold text-amber-400 font-mono">{audit.length}</p>
            </div>
          </div>

          {/* Department / Manager rollup ranking */}
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
            <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-700">
              <div className="flex items-center gap-2">
                <Crown className="w-4 h-4 text-amber-400" />
                <span className="text-xs text-neutral-300 tracking-wider font-medium">TEAM / DEPARTMENT COMPLIANCE RANKING</span>
              </div>
              <span className="text-xs text-neutral-500">{deptRollup.length} team{deptRollup.length > 1 ? "s" : ""}</span>
            </div>
            <table className="w-full text-sm">
              <thead>
                <tr className="text-xs text-neutral-500 border-b border-neutral-800">
                  <th className="text-left py-2 px-4 font-medium">RANK</th>
                  <th className="text-left py-2 px-4 font-medium">DEPARTMENT / TEAM</th>
                  <th className="text-center py-2 px-4 font-medium">HEADCOUNT</th>
                  <th className="text-center py-2 px-4 font-medium">COMPLIANCE</th>
                  <th className="text-center py-2 px-4 font-medium">CHRONIC OFFENDERS</th>
                </tr>
              </thead>
              <tbody>
                {deptRollup.map((d, i) => (
                  <tr key={d.dept} className="border-b border-neutral-800 hover:bg-neutral-800/50">
                    <td className="py-2.5 px-4 text-neutral-500 font-mono">#{i + 1}</td>
                    <td className="py-2.5 px-4 text-white font-medium">{d.dept}</td>
                    <td className="py-2.5 px-4 text-center font-mono text-neutral-300">{d.headcount}</td>
                    <td className={`py-2.5 px-4 text-center font-mono font-bold ${complianceColor(d.compliance)}`}>{d.compliance}%</td>
                    <td className="py-2.5 px-4 text-center font-mono">
                      {d.chronic > 0 ? <span className="text-red-400">{d.chronic}</span> : <span className="text-neutral-600">0</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Audit issues */}
          {audit.length > 0 && (
            <div className="bg-neutral-900 border border-amber-500/30 rounded-xl p-4 space-y-2">
              <div className="flex items-center gap-2 mb-2">
                <AlertTriangle className="w-4 h-4 text-amber-400" />
                <span className="text-xs text-amber-300 tracking-wider font-medium">DATA AUDIT — {audit.length} ISSUE{audit.length > 1 ? "S" : ""}</span>
              </div>
              <div className="space-y-1.5 max-h-48 overflow-y-auto">
                {audit.map((a, i) => (
                  <div key={i} className={`text-xs flex items-start gap-2 ${a.severity === "high" ? "text-red-300" : "text-neutral-400"}`}>
                    <span className={`mt-0.5 w-1.5 h-1.5 rounded-full shrink-0 ${a.severity === "high" ? "bg-red-400" : "bg-amber-400"}`} />
                    {a.message}
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Employee table */}
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
            <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-700 flex-wrap gap-2">
              <span className="text-xs text-neutral-300 tracking-wider font-medium">EMPLOYEE-LEVEL DETAIL</span>
              <select
                value={deptFilter} onChange={(e) => setDeptFilter(e.target.value)}
                className="bg-neutral-800 border border-neutral-600 text-xs text-neutral-300 rounded px-3 py-1.5 focus:outline-none focus:border-cyan-500"
              >
                {departments.map(d => <option key={d}>{d}</option>)}
              </select>
            </div>
            <div className="overflow-x-auto max-h-[480px] overflow-y-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-xs text-neutral-500 border-b border-neutral-800 sticky top-0 bg-neutral-900">
                    <th className="text-left py-2 px-3 font-medium">EMPLOYEE</th>
                    <th className="text-left py-2 px-3 font-medium">CODE</th>
                    <th className="text-left py-2 px-3 font-medium">DEPARTMENT</th>
                    <th className="text-center py-2 px-3 font-medium">DAYS REPORTED</th>
                    <th className="text-center py-2 px-3 font-medium">COMPLIANCE</th>
                    <th className="text-center py-2 px-3 font-medium">FOLLOW-UP</th>
                    <th className="text-center py-2 px-3 font-medium">RAW LOG</th>
                    <th className="text-left py-2 px-3 font-medium">STATUS</th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.map(e => {
                    const c = pct(e.daysReported, e.daysPossible);
                    return (
                      <tr key={e.code} className="border-b border-neutral-800 hover:bg-neutral-800/40">
                        <td className="py-2 px-3 text-white">{e.name}</td>
                        <td className="py-2 px-3 font-mono text-cyan-400">{e.code}</td>
                        <td className="py-2 px-3 text-neutral-400">{e.department}</td>
                        <td className="py-2 px-3 text-center font-mono text-neutral-300">{e.daysReported}/{e.daysPossible}</td>
                        <td className={`py-2 px-3 text-center font-mono font-bold ${complianceColor(c)}`}>{c}%</td>
                        <td className="py-2 px-3 text-center">
                          {e.inFollowUp ? <CheckCircle className="w-3.5 h-3.5 text-emerald-400 inline" /> : <XCircle className="w-3.5 h-3.5 text-red-400 inline" />}
                        </td>
                        <td className="py-2 px-3 text-center">
                          {e.inRaw ? <CheckCircle className="w-3.5 h-3.5 text-emerald-400 inline" /> : <XCircle className="w-3.5 h-3.5 text-red-400 inline" />}
                        </td>
                        <td className="py-2 px-3">
                          {c < 50
                            ? <span className="px-2 py-0.5 rounded text-[10px] font-bold uppercase bg-red-500/20 text-red-400 border border-red-500/30">At Risk</span>
                            : c < 80
                              ? <span className="px-2 py-0.5 rounded text-[10px] font-bold uppercase bg-amber-500/20 text-amber-400 border border-amber-500/30">Watch</span>
                              : <span className="px-2 py-0.5 rounded text-[10px] font-bold uppercase bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">Healthy</span>}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}
