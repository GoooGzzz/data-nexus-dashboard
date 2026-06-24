"use client";
import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Upload, FileSpreadsheet, CheckCircle, AlertTriangle,
  TrendingUp, TrendingDown, BarChart3, Users,
  Activity, Package, Table2, Layers, RefreshCw,
  Zap, Target, ArrowUpRight, ArrowDownRight, Star,
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Tooltip,
  PieChart, Pie, Cell, LineChart, Line, CartesianGrid,
  RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar, Legend,
} from "recharts";

// ─── TYPES ──────────────────────────────────────────────────────────────────
interface ParsedFile {
  name: string;
  type: "sales" | "attendance" | "unknown";
  rows: number;
  sheets: string[];
  data: Record<string, unknown>[];
  columns: string[];
  size: string;
}

interface SalesInsight {
  totalSellout: number;
  totalRevenue: number;
  topArea: string;
  topBrand: string;
  weekTrend: number;
  avgPerShop: number;
  shops: number;
  areaBreakdown: { area: string; sellout: number; pct: number }[];
  brandBreakdown: { name: string; value: number }[];
  weeklyBreakdown: { week: string; sellout: number }[];
  projectBreakdown: { project: string; sellout: number }[];
}

interface AttInsight {
  totalRecords: number;
  uniqueShops: number;
  uniqueEmps: number;
  avgDuration: number;
  shopBreakdown: { shop: string; count: number }[];
  dateBreakdown: { date: string; count: number }[];
}

const COLORS = ["#06b6d4","#8b5cf6","#f59e0b","#22c55e","#ef4444","#ec4899","#14b8a6","#f97316"];

function fmtNum(n: number) { return n.toLocaleString(); }
function fmtEGP(n: number) {
  if (n >= 1_000_000) return `EGP ${(n / 1_000_000).toFixed(1)}M`;
  if (n >= 1_000)     return `EGP ${(n / 1_000).toFixed(0)}K`;
  return `EGP ${n}`;
}
function fmtBytes(b: number) {
  return b > 1_048_576 ? `${(b / 1_048_576).toFixed(1)} MB` : `${(b / 1024).toFixed(0)} KB`;
}

function detectFileType(name: string, cols: string[]): "sales" | "attendance" | "unknown" {
  const n = name.toLowerCase();
  const colStr = cols.join(" ").toLowerCase();
  if (n.includes("attend") || colStr.includes("emp_code") || colStr.includes("start_work")) return "attendance";
  if (n.includes("sales") || n.includes("mbl") || colStr.includes("sellout")) return "sales";
  return "unknown";
}

function analyzeSales(data: Record<string, unknown>[]): SalesInsight {
  const first = data[0] || {};
  const key = (terms: string[]) => Object.keys(first).find(k => terms.some(t => k.toLowerCase().includes(t))) ?? terms[0];
  const selloutKey = key(["sellout"]);
  const priceKey   = key(["price"]);
  const areaKey    = key(["area"]);
  const brandKey   = key(["brand"]);
  const projectKey = key(["project"]);
  const shopKey    = key(["shop code","shop_code"]);
  const weekKey    = Object.keys(first).find(k => k.toLowerCase() === "w") ?? "W";

  let totalSellout = 0, totalRevenue = 0;
  const areaMap: Record<string,number> = {};
  const brandMap: Record<string,number> = {};
  const projectMap: Record<string,number> = {};
  const weekMap: Record<string,number> = {};
  const shops = new Set<string>();

  for (const row of data) {
    const s = Number(row[selloutKey]) || 0;
    const p = Number(row[priceKey]) || 0;
    totalSellout += s; totalRevenue += s * p;
    const area = String(row[areaKey] || "Unknown");
    areaMap[area] = (areaMap[area] || 0) + s;
    const brand = String(row[brandKey] || "Unknown");
    brandMap[brand] = (brandMap[brand] || 0) + s;
    const proj = String(row[projectKey] || "Unknown");
    projectMap[proj] = (projectMap[proj] || 0) + s;
    const wk = String(row[weekKey] || "");
    if (wk) weekMap[wk] = (weekMap[wk] || 0) + s;
    const sh = String(row[shopKey] || "");
    if (sh) shops.add(sh);
  }

  const areaBreakdown = Object.entries(areaMap)
    .map(([area, sellout]) => ({ area, sellout, pct: Math.round(sellout / totalSellout * 100) }))
    .sort((a,b) => b.sellout - a.sellout).slice(0,8);
  const brandBreakdown = Object.entries(brandMap)
    .map(([name, value]) => ({ name, value }))
    .sort((a,b) => b.value - a.value).slice(0,8);
  const projectBreakdown = Object.entries(projectMap)
    .map(([project, sellout]) => ({ project, sellout }))
    .sort((a,b) => b.sellout - a.sellout).slice(0,6);
  const weeklyBreakdown = Object.entries(weekMap)
    .map(([week, sellout]) => ({ week, sellout }))
    .sort((a,b) => a.week.localeCompare(b.week));
  const weeks = weeklyBreakdown;
  const weekTrend = weeks.length >= 2
    ? ((weeks[weeks.length-1].sellout - weeks[0].sellout) / weeks[0].sellout) * 100 : 0;

  return {
    totalSellout, totalRevenue, weekTrend, shops: shops.size,
    avgPerShop: shops.size ? Math.round(totalSellout / shops.size) : 0,
    topArea: areaBreakdown[0]?.area || "—",
    topBrand: brandBreakdown[0]?.name || "—",
    areaBreakdown, brandBreakdown, weeklyBreakdown, projectBreakdown,
  };
}

function analyzeAttendance(data: Record<string, unknown>[]): AttInsight {
  const first = data[0] || {};
  const key = (terms: string[]) => Object.keys(first).find(k => terms.some(t => k.toLowerCase().includes(t))) ?? terms[0];
  const shopKey = key(["shop code","shop_code"]);
  const empKey  = key(["emp_code"]);
  const durKey  = key(["duration"]);
  const dateKey = Object.keys(first).find(k => k.toLowerCase() === "date") ?? "date";

  const shopMap: Record<string,number> = {};
  const dateMap: Record<string,number> = {};
  const emps = new Set<string>();
  let totalDur = 0, durCount = 0;

  for (const row of data) {
    const shop = String(row[shopKey] || "");
    if (shop) shopMap[shop] = (shopMap[shop] || 0) + 1;
    const emp = String(row[empKey] || "");
    if (emp) emps.add(emp);
    const dur = Number(row[durKey]);
    if (!isNaN(dur) && dur > 0) { totalDur += dur; durCount++; }
    const dt = String(row[dateKey] || "");
    if (dt) { const d = dt.slice(0,10); dateMap[d] = (dateMap[d] || 0) + 1; }
  }

  return {
    totalRecords: data.length, uniqueShops: Object.keys(shopMap).length,
    uniqueEmps: emps.size, avgDuration: durCount ? Math.round(totalDur / durCount) : 0,
    shopBreakdown: Object.entries(shopMap).sort((a,b)=>b[1]-a[1]).slice(0,10)
      .map(([shop,count]) => ({ shop, count })),
    dateBreakdown: Object.entries(dateMap).sort((a,b)=>a[0].localeCompare(b[0])).slice(0,20)
      .map(([date,count]) => ({ date: date.slice(5), count })),
  };
}

// ─── STAT CARD ───────────────────────────────────────────────────────────────
function StatCard({ label, value, sub, color="cyan", icon:Icon, trend }: {
  label:string; value:string; sub?:string; color?:string;
  icon:React.ElementType; trend?:number;
}) {
  const palettes: Record<string,[string,string,string]> = {
    cyan:    ["text-cyan-400",    "border-cyan-500/30",    "bg-cyan-500/5"],
    purple:  ["text-purple-400",  "border-purple-500/30",  "bg-purple-500/5"],
    amber:   ["text-amber-400",   "border-amber-500/30",   "bg-amber-500/5"],
    emerald: ["text-emerald-400", "border-emerald-500/30", "bg-emerald-500/5"],
    red:     ["text-red-400",     "border-red-500/30",     "bg-red-500/5"],
  };
  const [tc, bc, bgc] = palettes[color] ?? palettes.cyan;
  return (
    <div className={`bg-neutral-900 border ${bc} ${bgc} rounded-xl p-5`}>
      <div className="flex items-center justify-between mb-3">
        <span className="text-xs text-neutral-500 tracking-wider font-medium uppercase">{label}</span>
        <Icon className={`w-4 h-4 ${tc}`} />
      </div>
      <p className={`text-2xl font-bold font-mono ${tc}`}>{value}</p>
      {(sub || trend !== undefined) && (
        <div className="flex items-center gap-2 mt-1">
          {sub && <p className="text-xs text-neutral-500">{sub}</p>}
          {trend !== undefined && (
            <span className={`flex items-center gap-0.5 text-xs font-medium ${trend >= 0 ? "text-emerald-400" : "text-red-400"}`}>
              {trend >= 0 ? <ArrowUpRight className="w-3 h-3" /> : <ArrowDownRight className="w-3 h-3" />}
              {Math.abs(trend).toFixed(1)}%
            </span>
          )}
        </div>
      )}
    </div>
  );
}

// ─── UPLOAD ZONE ─────────────────────────────────────────────────────────────
function UploadZone({ onFiles }: { onFiles:(f:ParsedFile[])=>void }) {
  const [dragging, setDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  const processFiles = useCallback(async (fileList: FileList) => {
    setProcessing(true); setProgress(0);
    const results: ParsedFile[] = [];
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      setProgress(Math.round(((i + 0.5) / fileList.length) * 100));
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json<Record<string,unknown>>(ws, { defval: "" });
        const columns = data.length > 0 ? Object.keys(data[0]) : [];
        results.push({ name: file.name, type: detectFileType(file.name, columns),
          rows: data.length, sheets: wb.SheetNames, data, columns, size: fmtBytes(file.size) });
      } catch {
        results.push({ name: file.name, type: "unknown", rows: 0,
          sheets: [], data: [], columns: [], size: fmtBytes(file.size) });
      }
      setProgress(Math.round(((i+1) / fileList.length) * 100));
    }
    setProcessing(false); onFiles(results);
  }, [onFiles]);

  return (
    <div
      onDragOver={e => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); if (e.dataTransfer.files.length) processFiles(e.dataTransfer.files); }}
      onClick={() => !processing && inputRef.current?.click()}
      className={`relative cursor-pointer rounded-2xl border-2 border-dashed transition-all duration-300 p-12 text-center group
        ${dragging ? "border-cyan-400 bg-cyan-500/10 scale-[1.01]" : "border-neutral-700 hover:border-cyan-500/50 bg-neutral-900/40 hover:bg-neutral-900"}`}
    >
      <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" multiple className="hidden"
        onChange={e => e.target.files && processFiles(e.target.files)} />
      {processing ? (
        <div className="space-y-4">
          <div className="w-14 h-14 mx-auto rounded-full bg-cyan-500/20 border border-cyan-500/40 flex items-center justify-center animate-spin">
            <RefreshCw className="w-6 h-6 text-cyan-400" />
          </div>
          <p className="text-cyan-400 font-medium">Parsing file…</p>
          <div className="w-56 mx-auto bg-neutral-800 rounded-full h-1.5 overflow-hidden">
            <div className="h-full bg-gradient-to-r from-cyan-500 to-purple-500 transition-all duration-300 rounded-full" style={{ width:`${progress}%` }} />
          </div>
          <p className="text-neutral-500 text-sm">{progress}%</p>
        </div>
      ) : (
        <div className="space-y-4">
          <div className={`w-16 h-16 mx-auto rounded-2xl flex items-center justify-center border transition-all duration-300
            ${dragging ? "bg-cyan-500/30 border-cyan-400 scale-110" : "bg-neutral-800 border-neutral-700 group-hover:bg-cyan-500/10 group-hover:border-cyan-500/40"}`}>
            <Upload className={`w-7 h-7 transition-colors ${dragging ? "text-cyan-300" : "text-neutral-400 group-hover:text-cyan-400"}`} />
          </div>
          <div>
            <p className="text-white font-semibold text-lg">Drop Excel files here</p>
            <p className="text-neutral-500 text-sm mt-1">or click to browse — .xlsx, .xls, .csv</p>
          </div>
          <div className="flex items-center justify-center gap-6 text-xs text-neutral-600">
            <span className="flex items-center gap-1"><FileSpreadsheet className="w-3.5 h-3.5" />Sales reports</span>
            <span className="flex items-center gap-1"><Users className="w-3.5 h-3.5" />Attendance logs</span>
            <span className="flex items-center gap-1"><Layers className="w-3.5 h-3.5" />Multi-sheet</span>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── SALES TABS ───────────────────────────────────────────────────────────────
function SalesAnalysis({ insight, file }: { insight: SalesInsight; file: ParsedFile }) {
  const [tab, setTab] = useState("overview");
  const tabs = [
    { id:"overview", label:"Overview",    icon:BarChart3 },
    { id:"area",     label:"By Area",     icon:Target },
    { id:"brand",    label:"By Brand",    icon:Package },
    { id:"trend",    label:"Weekly",      icon:TrendingUp },
    { id:"cross",    label:"Cross-Check", icon:Layers },
    { id:"table",    label:"Raw Data",    icon:Table2 },
  ];

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-1 p-1 bg-neutral-900 border border-neutral-700 rounded-xl">
        {tabs.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-medium transition-all
              ${tab===t.id ? "bg-cyan-500/20 text-cyan-400 border border-cyan-500/40 shadow-lg shadow-cyan-500/10" : "text-neutral-500 hover:text-neutral-300"}`}>
            <t.icon className="w-3.5 h-3.5" />{t.label}
          </button>
        ))}
      </div>

      {tab === "overview" && (
        <div className="space-y-4">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <StatCard label="Total Sellout" value={fmtNum(insight.totalSellout)} sub="units" color="cyan" icon={Package} />
            <StatCard label="Revenue Est." value={fmtEGP(insight.totalRevenue)} color="emerald" icon={TrendingUp} />
            <StatCard label="Active Shops" value={fmtNum(insight.shops)} color="purple" icon={Target} />
            <StatCard label="Week Trend" value={`${insight.weekTrend>0?"+":""}${insight.weekTrend.toFixed(1)}%`}
              color={insight.weekTrend>=0?"emerald":"red"} icon={Activity} trend={insight.weekTrend} />
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5">
              <p className="text-xs text-neutral-500 tracking-wider mb-4 font-medium">SALES BY AREA</p>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={insight.areaBreakdown}>
                  <XAxis dataKey="area" tick={{ fill:"#737373", fontSize:11 }} />
                  <YAxis tick={{ fill:"#737373", fontSize:11 }} />
                  <Tooltip contentStyle={{ background:"#1a1a1a", border:"1px solid #333", borderRadius:8 }} />
                  <Bar dataKey="sellout" radius={[4,4,0,0]}>
                    {insight.areaBreakdown.map((_,i) => <Cell key={i} fill={COLORS[i%COLORS.length]} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5">
              <p className="text-xs text-neutral-500 tracking-wider mb-4 font-medium">BRAND SHARE</p>
              <ResponsiveContainer width="100%" height={200}>
                <PieChart>
                  <Pie data={insight.brandBreakdown} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={75} paddingAngle={3}>
                    {insight.brandBreakdown.map((_,i) => <Cell key={i} fill={COLORS[i%COLORS.length]} />)}
                  </Pie>
                  <Tooltip contentStyle={{ background:"#1a1a1a", border:"1px solid #333", borderRadius:8 }} />
                  <Legend wrapperStyle={{ fontSize:11, color:"#737373" }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      )}

      {tab === "area" && (
        <div className="space-y-4">
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
            <div className="px-5 py-3 border-b border-neutral-700">
              <p className="text-xs text-neutral-400 tracking-wider font-medium">AREA PERFORMANCE BREAKDOWN</p>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-xs text-neutral-500 border-b border-neutral-800">
                    {["RANK","AREA","SELLOUT","SHARE","BAR"].map(h=>(
                      <th key={h} className="text-left py-3 px-4 font-medium">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {insight.areaBreakdown.map((a,i) => (
                    <tr key={a.area} className="border-b border-neutral-800 hover:bg-neutral-800/50 transition-colors">
                      <td className="py-3 px-4 text-neutral-500 font-mono text-xs">#{i+1}</td>
                      <td className="py-3 px-4 text-white font-medium">{a.area}</td>
                      <td className="py-3 px-4 text-cyan-400 font-mono">{fmtNum(a.sellout)}</td>
                      <td className="py-3 px-4 text-neutral-400 font-mono">{a.pct}%</td>
                      <td className="py-3 px-4">
                        <div className="w-32 bg-neutral-800 rounded-full h-2">
                          <div className="h-2 rounded-full" style={{ width:`${a.pct}%`, background:COLORS[i%COLORS.length] }} />
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5">
            <p className="text-xs text-neutral-500 tracking-wider mb-4 font-medium">AREA RADAR</p>
            <ResponsiveContainer width="100%" height={260}>
              <RadarChart data={insight.areaBreakdown.slice(0,6)}>
                <PolarGrid stroke="#333" />
                <PolarAngleAxis dataKey="area" tick={{ fill:"#737373", fontSize:11 }} />
                <PolarRadiusAxis tick={{ fill:"#737373", fontSize:9 }} />
                <Radar dataKey="sellout" stroke="#06b6d4" fill="#06b6d4" fillOpacity={0.25} />
              </RadarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {tab === "brand" && (
        <div className="space-y-4">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            {insight.brandBreakdown.slice(0,4).map((b,i) => (
              <div key={b.name} className="bg-neutral-900 border border-neutral-700 rounded-xl p-4">
                <div className="flex items-center gap-2 mb-2">
                  <div className="w-2 h-2 rounded-full" style={{ background:COLORS[i] }} />
                  <span className="text-xs text-neutral-400 font-medium">{b.name}</span>
                </div>
                <p className="text-xl font-bold font-mono text-white">{fmtNum(b.value)}</p>
                <p className="text-xs text-neutral-600 mt-0.5">{Math.round(b.value/insight.totalSellout*100)}% share</p>
              </div>
            ))}
          </div>
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5">
            <p className="text-xs text-neutral-500 tracking-wider mb-4 font-medium">BRAND RANKING</p>
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={insight.brandBreakdown} layout="vertical">
                <XAxis type="number" tick={{ fill:"#737373", fontSize:11 }} />
                <YAxis type="category" dataKey="name" tick={{ fill:"#a3a3a3", fontSize:12 }} width={70} />
                <Tooltip contentStyle={{ background:"#1a1a1a", border:"1px solid #333", borderRadius:8 }} />
                <Bar dataKey="value" radius={[0,4,4,0]}>
                  {insight.brandBreakdown.map((_,i) => <Cell key={i} fill={COLORS[i%COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {tab === "trend" && (
        <div className="space-y-4">
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5">
            <p className="text-xs text-neutral-500 tracking-wider mb-4 font-medium">WEEKLY SELLOUT TREND</p>
            <ResponsiveContainer width="100%" height={260}>
              <LineChart data={insight.weeklyBreakdown}>
                <CartesianGrid strokeDasharray="3 3" stroke="#262626" />
                <XAxis dataKey="week" tick={{ fill:"#737373", fontSize:11 }} />
                <YAxis tick={{ fill:"#737373", fontSize:11 }} />
                <Tooltip contentStyle={{ background:"#1a1a1a", border:"1px solid #333", borderRadius:8 }} />
                <Line type="monotone" dataKey="sellout" stroke="#06b6d4" strokeWidth={2.5}
                  dot={{ fill:"#06b6d4", r:4 }} activeDot={{ r:6, fill:"#22d3ee" }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5">
            <p className="text-xs text-neutral-500 tracking-wider mb-4 font-medium">PROJECT PERFORMANCE</p>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={insight.projectBreakdown}>
                <XAxis dataKey="project" tick={{ fill:"#737373", fontSize:11 }} />
                <YAxis tick={{ fill:"#737373", fontSize:11 }} />
                <Tooltip contentStyle={{ background:"#1a1a1a", border:"1px solid #333", borderRadius:8 }} />
                <Bar dataKey="sellout" radius={[4,4,0,0]}>
                  {insight.projectBreakdown.map((_,i) => <Cell key={i} fill={COLORS[i%COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {tab === "cross" && (
        <div className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="bg-neutral-900 border border-amber-500/30 rounded-xl p-5">
              <div className="flex items-center gap-2 mb-4">
                <AlertTriangle className="w-4 h-4 text-amber-400" />
                <p className="text-xs text-amber-400 tracking-wider font-medium">ANOMALIES TO INVESTIGATE</p>
              </div>
              <div className="space-y-3">
                {[
                  { label:"Zero-sellout shops (have attendance)", desc:"Staff present but 0 sales — ghost visit or data gap" },
                  { label:"High sellout / no attendance",        desc:"Sales with no staff logged — potential fraud signal" },
                  { label:"Area decline > 10% WoW",             desc:"Week-over-week drop per area — needs field check" },
                  { label:"Single-employee shops",              desc:"Only 1 employee logged — unverified sales risk" },
                ].map((item,i) => (
                  <div key={i} className="flex gap-3 p-3 bg-neutral-800/50 rounded-lg border border-neutral-700">
                    <div className="w-5 h-5 shrink-0 rounded-full bg-amber-500/20 border border-amber-500/40 flex items-center justify-center mt-0.5">
                      <span className="text-amber-400 text-xs font-bold">{i+1}</span>
                    </div>
                    <div>
                      <p className="text-sm text-white font-medium">{item.label}</p>
                      <p className="text-xs text-neutral-500 mt-0.5">{item.desc}</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
            <div className="bg-neutral-900 border border-cyan-500/30 rounded-xl p-5">
              <div className="flex items-center gap-2 mb-4">
                <Star className="w-4 h-4 text-cyan-400" />
                <p className="text-xs text-cyan-400 tracking-wider font-medium">INVESTIGATION IDEAS</p>
              </div>
              <div className="space-y-3">
                {[
                  { label:"Join on Shop Code",              desc:"Match sales + attendance by shop code to detect ghost-sales" },
                  { label:"Time-window match",              desc:"Validate sale timestamps fall within employee shift" },
                  { label:"Employee-to-sales ratio",        desc:"Low → under-staffed; high → possible padding" },
                  { label:"Weekly attendance vs sales corr",desc:"Pearson correlation — gaps reveal operational issues" },
                  { label:"Brand concentration risk",       desc:"Over-reliance on Samsung? Analyze substitution effect" },
                ].map((item,i) => (
                  <div key={i} className="flex gap-3 p-3 bg-neutral-800/50 rounded-lg border border-neutral-700">
                    <div className="w-5 h-5 shrink-0 rounded-full bg-cyan-500/20 border border-cyan-500/40 flex items-center justify-center mt-0.5">
                      <Zap className="w-3 h-3 text-cyan-400" />
                    </div>
                    <div>
                      <p className="text-sm text-white font-medium">{item.label}</p>
                      <p className="text-xs text-neutral-500 mt-0.5">{item.desc}</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5">
            <p className="text-xs text-neutral-500 tracking-wider mb-3 font-medium">FILE QUALITY METRICS</p>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {[
                { label:"Shops Covered",   value: fmtNum(insight.shops),          note:"unique codes" },
                { label:"Avg Sellout/Shop",value: fmtNum(insight.avgPerShop),      note:"units" },
                { label:"Top Area",        value: insight.topArea,                 note:"by sellout" },
                { label:"Trend",           value: insight.weekTrend >= 0 ? "↑ Growth" : "↓ Decline", note:`${insight.weekTrend.toFixed(1)}%` },
              ].map((m,i) => (
                <div key={i} className="p-3 bg-neutral-800 rounded-lg border border-neutral-700">
                  <p className="text-xs text-neutral-500">{m.label}</p>
                  <p className="text-base font-bold font-mono text-white mt-1">{m.value}</p>
                  <p className="text-xs text-neutral-600">{m.note}</p>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {tab === "table" && (
        <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
          <div className="px-5 py-3 border-b border-neutral-700 flex items-center justify-between">
            <p className="text-xs text-neutral-400 tracking-wider font-medium">RAW DATA — first 50 rows</p>
            <span className="text-xs text-neutral-600">{fmtNum(file.rows)} total</span>
          </div>
          <div className="overflow-auto max-h-96">
            <table className="w-full text-xs">
              <thead className="sticky top-0 bg-neutral-800">
                <tr>{file.columns.slice(0,10).map(c => (
                  <th key={c} className="text-left py-2 px-3 text-neutral-400 font-medium whitespace-nowrap border-b border-neutral-700">{c}</th>
                ))}</tr>
              </thead>
              <tbody>
                {file.data.slice(0,50).map((row,i) => (
                  <tr key={i} className="border-b border-neutral-800 hover:bg-neutral-800/50">
                    {file.columns.slice(0,10).map(c => (
                      <td key={c} className="py-2 px-3 text-neutral-300 whitespace-nowrap max-w-[140px] truncate">{String(row[c] ?? "")}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── ATTENDANCE TABS ──────────────────────────────────────────────────────────
function AttendanceAnalysis({ insight, file }: { insight: AttInsight; file: ParsedFile }) {
  const [tab, setTab] = useState("overview");
  const tabs = [
    { id:"overview", label:"Overview",      icon:BarChart3 },
    { id:"daily",    label:"Daily Pattern", icon:Activity },
    { id:"shops",    label:"Top Shops",     icon:Target },
    { id:"table",    label:"Raw Data",      icon:Table2 },
  ];

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-1 p-1 bg-neutral-900 border border-neutral-700 rounded-xl">
        {tabs.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-medium transition-all
              ${tab===t.id ? "bg-purple-500/20 text-purple-400 border border-purple-500/40" : "text-neutral-500 hover:text-neutral-300"}`}>
            <t.icon className="w-3.5 h-3.5" />{t.label}
          </button>
        ))}
      </div>

      {tab === "overview" && (
        <div className="space-y-4">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <StatCard label="Total Records"    value={fmtNum(insight.totalRecords)} color="purple" icon={Layers} />
            <StatCard label="Unique Shops"     value={fmtNum(insight.uniqueShops)}  color="cyan"   icon={Target} />
            <StatCard label="Unique Employees" value={fmtNum(insight.uniqueEmps)}   color="amber"  icon={Users} />
            <StatCard label="Avg Duration"     value={`${insight.avgDuration}min`} color="emerald" icon={Activity} />
          </div>
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5">
            <p className="text-xs text-neutral-500 tracking-wider mb-4 font-medium">DAILY ATTENDANCE VOLUME</p>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={insight.dateBreakdown}>
                <XAxis dataKey="date" tick={{ fill:"#737373", fontSize:10 }} />
                <YAxis tick={{ fill:"#737373", fontSize:11 }} />
                <Tooltip contentStyle={{ background:"#1a1a1a", border:"1px solid #333", borderRadius:8 }} />
                <Bar dataKey="count" fill="#8b5cf6" radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {tab === "daily" && (
        <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5">
          <p className="text-xs text-neutral-500 tracking-wider mb-4 font-medium">DAILY CHECK-INS TREND</p>
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={insight.dateBreakdown}>
              <CartesianGrid strokeDasharray="3 3" stroke="#262626" />
              <XAxis dataKey="date" tick={{ fill:"#737373", fontSize:10 }} />
              <YAxis tick={{ fill:"#737373", fontSize:11 }} />
              <Tooltip contentStyle={{ background:"#1a1a1a", border:"1px solid #333", borderRadius:8 }} />
              <Line type="monotone" dataKey="count" stroke="#8b5cf6" strokeWidth={2.5} dot={{ fill:"#8b5cf6", r:3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}

      {tab === "shops" && (
        <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
          <div className="px-5 py-3 border-b border-neutral-700">
            <p className="text-xs text-neutral-400 tracking-wider font-medium">TOP 10 SHOPS BY CHECK-INS</p>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-xs text-neutral-500 border-b border-neutral-800">
                  {["RANK","SHOP CODE","CHECK-INS","BAR"].map(h=>(
                    <th key={h} className="text-left py-3 px-4 font-medium">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {insight.shopBreakdown.map((s,i) => (
                  <tr key={s.shop} className="border-b border-neutral-800 hover:bg-neutral-800/50">
                    <td className="py-3 px-4 text-neutral-500 text-xs font-mono">#{i+1}</td>
                    <td className="py-3 px-4 text-cyan-400 font-mono text-xs">{s.shop}</td>
                    <td className="py-3 px-4 text-white font-mono">{fmtNum(s.count)}</td>
                    <td className="py-3 px-4">
                      <div className="w-32 bg-neutral-800 rounded-full h-2">
                        <div className="h-2 rounded-full bg-purple-500"
                          style={{ width:`${Math.round(s.count / insight.shopBreakdown[0].count * 100)}%` }} />
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab === "table" && (
        <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
          <div className="px-5 py-3 border-b border-neutral-700 flex items-center justify-between">
            <p className="text-xs text-neutral-400 tracking-wider font-medium">RAW DATA — first 50 rows</p>
            <span className="text-xs text-neutral-600">{fmtNum(file.rows)} total</span>
          </div>
          <div className="overflow-auto max-h-96">
            <table className="w-full text-xs">
              <thead className="sticky top-0 bg-neutral-800">
                <tr>{file.columns.slice(0,8).map(c => (
                  <th key={c} className="text-left py-2 px-3 text-neutral-400 font-medium whitespace-nowrap border-b border-neutral-700">{c}</th>
                ))}</tr>
              </thead>
              <tbody>
                {file.data.slice(0,50).map((row,i) => (
                  <tr key={i} className="border-b border-neutral-800 hover:bg-neutral-800/50">
                    {file.columns.slice(0,8).map(c => (
                      <td key={c} className="py-2 px-3 text-neutral-300 whitespace-nowrap max-w-[140px] truncate">{String(row[c] ?? "")}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
export default function ImportReportsPage() {
  const [files, setFiles] = useState<ParsedFile[]>([]);
  const [activeFile, setActiveFile] = useState<string | null>(null);

  const handleFiles = useCallback((newFiles: ParsedFile[]) => {
    setFiles(prev => {
      const existingNames = new Set(prev.map(f => f.name));
      const merged = [...prev, ...newFiles.filter(f => !existingNames.has(f.name))];
      return merged;
    });
    if (newFiles.length > 0) setActiveFile(newFiles[0].name);
  }, []);

  const current = files.find(f => f.name === activeFile);
  const salesInsight  = current?.type === "sales"      ? analyzeSales(current.data)      : null;
  const attInsight    = current?.type === "attendance"  ? analyzeAttendance(current.data) : null;
  const salesFiles    = files.filter(f => f.type === "sales");
  const attFiles      = files.filter(f => f.type === "attendance");

  return (
    <div className="p-6 space-y-6 max-w-full">
      {/* Header */}
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <h2 className="text-xl font-bold text-white tracking-wider">IMPORT REPORTS</h2>
          <p className="text-xs text-neutral-500 mt-0.5">Upload Excel files for instant analysis, cross-checking & anomaly detection</p>
        </div>
        {files.length > 0 && (
          <div className="flex items-center gap-2">
            <span className="text-xs text-neutral-500 bg-neutral-800 border border-neutral-700 px-3 py-1.5 rounded-lg">
              {files.length} file{files.length>1?"s":""} loaded
            </span>
            <button onClick={() => { setFiles([]); setActiveFile(null); }}
              className="text-xs text-red-400 bg-red-500/10 border border-red-500/30 px-3 py-1.5 rounded-lg hover:bg-red-500/20 transition-colors">
              Clear all
            </button>
          </div>
        )}
      </div>

      <UploadZone onFiles={handleFiles} />

      {files.length > 0 && (
        <div className="space-y-4">
          {/* File pills */}
          <div className="flex flex-wrap gap-2">
            {files.map(f => (
              <button key={f.name} onClick={() => setActiveFile(f.name)}
                className={`flex items-center gap-2 px-4 py-2 rounded-xl border text-xs font-medium transition-all
                  ${activeFile===f.name
                    ? f.type==="sales"?"bg-cyan-500/20 border-cyan-500/50 text-cyan-300"
                      : f.type==="attendance"?"bg-purple-500/20 border-purple-500/50 text-purple-300"
                        :"bg-neutral-700 border-neutral-500 text-white"
                    : "bg-neutral-900 border-neutral-700 text-neutral-400 hover:border-neutral-500"}`}>
                <FileSpreadsheet className="w-3.5 h-3.5" />
                <span className="max-w-[180px] truncate">{f.name}</span>
                <span className={`px-1.5 py-0.5 rounded text-[10px]
                  ${f.type==="sales"?"bg-cyan-500/20 text-cyan-400"
                    :f.type==="attendance"?"bg-purple-500/20 text-purple-400"
                      :"bg-neutral-700 text-neutral-400"}`}>
                  {f.type==="unknown"?"?":f.type}
                </span>
                <span className="text-neutral-600">{fmtNum(f.rows)}r</span>
              </button>
            ))}
          </div>

          {/* Current file bar */}
          {current && (
            <div className="flex flex-wrap items-center gap-3 p-4 bg-neutral-900 border border-neutral-700 rounded-xl">
              <div className={`w-8 h-8 rounded-lg flex items-center justify-center border
                ${current.type==="sales"?"bg-cyan-500/20 border-cyan-500/30":"bg-purple-500/20 border-purple-500/30"}`}>
                <FileSpreadsheet className={`w-4 h-4 ${current.type==="sales"?"text-cyan-400":"text-purple-400"}`} />
              </div>
              <div>
                <p className="text-sm text-white font-medium">{current.name}</p>
                <p className="text-xs text-neutral-500">{fmtNum(current.rows)} rows · {current.columns.length} cols · {current.size} · {current.sheets.length} sheet{current.sheets.length>1?"s":""}</p>
              </div>
              {current.type !== "unknown" && (
                <span className="ml-auto flex items-center gap-1 text-xs text-emerald-400 bg-emerald-500/10 border border-emerald-500/30 px-2.5 py-1 rounded-lg">
                  <CheckCircle className="w-3 h-3" />
                  Auto-detected: {current.type}
                </span>
              )}
            </div>
          )}

          {current && salesInsight && <SalesAnalysis insight={salesInsight} file={current} />}
          {current && attInsight   && <AttendanceAnalysis insight={attInsight} file={current} />}
          {current && current.type === "unknown" && (
            <div className="bg-neutral-900 border border-amber-500/30 rounded-xl p-8 text-center">
              <AlertTriangle className="w-10 h-10 text-amber-400 mx-auto mb-3" />
              <p className="text-amber-400 font-medium mb-1">File type not recognised</p>
              <p className="text-sm text-neutral-400">Detected columns: {current.columns.slice(0,6).join(", ")}{current.columns.length>6?"…":""}</p>
              <p className="text-xs text-neutral-600 mt-2">Expected: "Sellout", "emp_code", "shop code", "duration"…</p>
            </div>
          )}

          {salesFiles.length > 0 && attFiles.length > 0 && (
            <div className="bg-gradient-to-br from-neutral-900 to-neutral-800/40 border border-emerald-500/30 rounded-xl p-5 space-y-3">
              <div className="flex items-center gap-2">
                <Layers className="w-4 h-4 text-emerald-400" />
                <p className="text-xs text-emerald-400 tracking-wider font-medium">MULTI-FILE CROSS-CHECK READY</p>
              </div>
              <p className="text-sm text-neutral-300">Both a <span className="text-cyan-400">sales file</span> and an <span className="text-purple-400">attendance file</span> are loaded. Ask Nexus AI to run a join analysis or cross-file anomaly check.</p>
              <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
                {[
                  { label:"Sales rows",     value: fmtNum(salesFiles.reduce((s,f)=>s+f.rows,0)), color:"text-cyan-400" },
                  { label:"Attendance rows",value: fmtNum(attFiles.reduce((s,f)=>s+f.rows,0)),   color:"text-purple-400" },
                  { label:"Join key",       value: "Shop Code",                                   color:"text-emerald-400" },
                ].map((m,i) => (
                  <div key={i} className="p-3 bg-neutral-800/60 rounded-lg border border-neutral-700">
                    <p className="text-xs text-neutral-500">{m.label}</p>
                    <p className={`text-base font-bold font-mono mt-1 ${m.color}`}>{m.value}</p>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* Empty state cards */}
      {files.length === 0 && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {[
            { icon:FileSpreadsheet, color:"cyan",    title:"Sales Reports",  desc:"MBL weekly sellout, brand × area breakdown, W-coded weeks", badge:".xlsx" },
            { icon:Users,           color:"purple",  title:"Attendance Logs",desc:"Employee check-in/out, duration, shop-level headcount",      badge:".xlsx" },
            { icon:Layers,          color:"emerald", title:"Cross-Analysis", desc:"Load both files to detect ghost-sales, staffing gaps, fraud", badge:"multi" },
          ].map((c,i) => {
            const palettes: Record<string,[string,string,string]> = {
              cyan:   ["text-cyan-400",   "border-cyan-500/30",   "bg-cyan-500/5"],
              purple: ["text-purple-400", "border-purple-500/30", "bg-purple-500/5"],
              emerald:["text-emerald-400","border-emerald-500/30","bg-emerald-500/5"],
            };
            const [tc,bc,bgc] = palettes[c.color];
            return (
              <div key={i} className={`bg-neutral-900 border ${bc} ${bgc} rounded-xl p-5`}>
                <div className={`w-10 h-10 rounded-xl ${bgc} border ${bc} flex items-center justify-center mb-3`}>
                  <c.icon className={`w-5 h-5 ${tc}`} />
                </div>
                <p className="text-white font-semibold mb-1">{c.title}</p>
                <p className="text-xs text-neutral-500 leading-relaxed">{c.desc}</p>
                <span className={`inline-block mt-3 text-xs px-2 py-0.5 rounded ${bgc} border ${bc} ${tc} font-mono`}>{c.badge}</span>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
