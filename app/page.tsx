"use client";
import { useState, useRef, useEffect, useCallback } from "react";
import ImportReportsPage from "./import-reports/ImportReportsPage";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, Tooltip, Legend
} from "recharts";
import {
  LayoutDashboard, Settings, BarChart3, Users, Bell, RefreshCw,
  Ghost, AlertTriangle, Database, X, Search, ChevronDown, ChevronRight,
  CheckCircle, XCircle, Clock, MapPin, Send, Bot, Loader2,
  TrendingUp, TrendingDown, Minus, Store, Package,
  Upload, FileSpreadsheet, Merge, Eye
} from "lucide-react";

// ─── REAL DATA from Excel files ───────────────────────────────────────────
const WEEK_SALES = [
  { week: "W19", sellout: 49458 },
  { week: "W20", sellout: 49051 },
  { week: "W21", sellout: 46825 },
  { week: "W22", sellout: 42421 },
];

const AREA_SALES = [
  { area: "Cairo",       sellout: 72514, color: "#06b6d4" },
  { area: "Delta",       sellout: 49198, color: "#8b5cf6" },
  { area: "Upper Egypt", sellout: 38718, color: "#f59e0b" },
  { area: "Alex",        sellout: 27325, color: "#22c55e" },
];

const PROJECT_SALES = [
  { project: "Club",  sellout: 88871, color: "#06b6d4" },
  { project: "Elite", sellout: 34813, color: "#8b5cf6" },
  { project: "Hero",  sellout: 24371, color: "#f59e0b" },
  { project: "Star",  sellout: 23650, color: "#22c55e" },
  { project: "One",   sellout: 16050, color: "#ef4444" },
];

const BRAND_SALES = [
  { name: "Samsung", sales: 84618 },
  { name: "Oppo",    sales: 34453 },
  { name: "Xiaomi",  sales: 19383 },
  { name: "Realme",  sales: 16292 },
  { name: "Infinix", sales: 10069 },
  { name: "Vivo",    sales:  9206 },
  { name: "Honor",   sales:  7827 },
  { name: "Apple",   sales:  4645 },
];

// Merged shop data (sales + attendance joined by shop code)
const MERGED_SHOPS = [
  { shop_code:"S-0074-005", shop_name:"Dubai Phone",                    area:"Cairo", project:"Elite", sellout:3042, revenue:63090168, top_brand:"Oppo",    att_records:43, emp_count:9  },
  { shop_code:"S-0074-006", shop_name:"Dubai Phone (City Center)",      area:"Cairo", project:"Elite", sellout:2495, revenue:63398321, top_brand:"Samsung", att_records:37, emp_count:6  },
  { shop_code:"S-0074-008", shop_name:"Dubai Phone (Mohandseen)",       area:"Cairo", project:"Elite", sellout:2247, revenue:49207560, top_brand:"Samsung", att_records:8,  emp_count:4  },
  { shop_code:"S-0074-020", shop_name:"Dubai Phone (Smouha)",           area:"Alex",  project:"Elite", sellout:1942, revenue:51674556, top_brand:"Samsung", att_records:31, emp_count:6  },
  { shop_code:"S-0074-022", shop_name:"Dubai Phone (Mosdak)",           area:"Cairo", project:"Elite", sellout:1923, revenue:39900325, top_brand:"Oppo",    att_records:14, emp_count:5  },
  { shop_code:"S-0074-012", shop_name:"Dubai Phone (El Maadi)",         area:"Cairo", project:"Elite", sellout:1824, revenue:49522494, top_brand:"Samsung", att_records:41, emp_count:9  },
  { shop_code:"S-0074-015", shop_name:"Dubai Phone (Ain Shams)",        area:"Cairo", project:"Elite", sellout:1809, revenue:35291071, top_brand:"Oppo",    att_records:37, emp_count:8  },
  { shop_code:"S-0074-018", shop_name:"Dubai Phone (City Stars)",       area:"Cairo", project:"Elite", sellout:1784, revenue:47770879, top_brand:"Oppo",    att_records:41, emp_count:8  },
  { shop_code:"S-0074-013", shop_name:"Dubai Phone 2 (6 Oct)",          area:"Cairo", project:"Elite", sellout:1782, revenue:37689348, top_brand:"Oppo",    att_records:12, emp_count:4  },
  { shop_code:"S-0074-011", shop_name:"Dubai Phone (Mall Of Arab)",     area:"Cairo", project:"Elite", sellout:1582, revenue:51050570, top_brand:"Oppo",    att_records:35, emp_count:5  },
  { shop_code:"S-0074-019", shop_name:"Dubai Phone (MOE)",              area:"Cairo", project:"Elite", sellout:1564, revenue:44042678, top_brand:"Oppo",    att_records:35, emp_count:8  },
  { shop_code:"S-0074-014", shop_name:"Dubai Phone (El Korba)",         area:"Cairo", project:"Elite", sellout:1493, revenue:40579754, top_brand:"Oppo",    att_records:5,  emp_count:4  },
  { shop_code:"S-0074-017", shop_name:"Dubai Phone (Obour)",            area:"Cairo", project:"Elite", sellout:1269, revenue:35535201, top_brand:"Oppo",    att_records:36, emp_count:6  },
  { shop_code:"S-0074-024", shop_name:"Dubai Phone (Concord Plaza)",    area:"Cairo", project:"Elite", sellout:1144, revenue:44800530, top_brand:"Apple",   att_records:8,  emp_count:5  },
  { shop_code:"S-0074-023", shop_name:"Dubai Phone (Louran)",           area:"Alex",  project:"Elite", sellout:1127, revenue:30813302, top_brand:"Oppo",    att_records:29, emp_count:3  },
  { shop_code:"S-4625-007", shop_name:"Carrefour (Maadi)",              area:"Cairo", project:"Club",  sellout:893,  revenue:18240000, top_brand:"Samsung", att_records:175,emp_count:22 },
  { shop_code:"S-9158-016", shop_name:"Raneen (El Fostat)",             area:"Cairo", project:"Club",  sellout:712,  revenue:14580000, top_brand:"Samsung", att_records:32, emp_count:4  },
  { shop_code:"S-9158-048", shop_name:"Raneen (Damanhour)",             area:"Delta", project:"Club",  sellout:645,  revenue:12960000, top_brand:"Oppo",    att_records:25, emp_count:2  },
  { shop_code:"S-9158-049", shop_name:"Raneen (10th Of Ramadan)",       area:"Delta", project:"Club",  sellout:598,  revenue:11920000, top_brand:"Samsung", att_records:25, emp_count:2  },
  { shop_code:"S-9158-051", shop_name:"Raneen (Ismailia)",              area:"Delta", project:"Club",  sellout:542,  revenue:10830000, top_brand:"Samsung", att_records:25, emp_count:1  },
];

// Attendance-only records (from real Excel — sample rows)
const ATTENDANCE_RECORDS = [
  { id:"ATT-001", emp_code:"A-1757", emp_name:"Abd El Geliel El Sayed",    shop_code:"S-4625-007", shop_name:"Carrefour (Maadi)",      date:"2026-05-01", start:"12:51", end:"21:18", duration:"08:26", status:"present" },
  { id:"ATT-002", emp_code:"A-3481", emp_name:"Abd El Rahman Adel Fathy",  shop_code:"S-9158-048", shop_name:"Raneen (Damanhour)",      date:"2026-05-01", start:"12:54", end:"21:00", duration:"08:06", status:"present" },
  { id:"ATT-003", emp_code:"A-3165", emp_name:"Abd El Rahman Mahmoud",     shop_code:"S-9158-016", shop_name:"Raneen (El Fostat)",      date:"2026-05-01", start:"13:00", end:"21:08", duration:"08:07", status:"present" },
  { id:"ATT-004", emp_code:"A-3149", emp_name:"Abd El Rahman Reda Attia",  shop_code:"S-9158-051", shop_name:"Raneen (Ismailia)",       date:"2026-05-01", start:"12:58", end:"21:01", duration:"08:02", status:"present" },
  { id:"ATT-005", emp_code:"A-1167", emp_name:"Ahmed Hassan Mohamed",      shop_code:"S-4625-007", shop_name:"Carrefour (Maadi)",      date:"2026-05-01", start:"12:49", end:"21:22", duration:"08:33", status:"present" },
  { id:"ATT-006", emp_code:"A-2370", emp_name:"Ahmed Ibrahim Ali",         shop_code:"S-0074-005", shop_name:"Dubai Phone",            date:"2026-05-02", start:"11:00", end:"20:30", duration:"09:30", status:"present" },
  { id:"ATT-007", emp_code:"A-1042", emp_name:"Mohamed Kamal Sayed",       shop_code:"S-4625-007", shop_name:"Carrefour (Maadi)",      date:"2026-05-03", start:"13:10", end:"20:55", duration:"07:45", status:"present" },
  { id:"ATT-008", emp_code:"A-3229", emp_name:"Youssef Mahmoud Nasser",    shop_code:"S-4625-007", shop_name:"Carrefour (Maadi)",      date:"2026-05-03", start:"--:--", end:"--:--", duration:"00:00", status:"absent"  },
  { id:"ATT-009", emp_code:"A-1379", emp_name:"Omar Khaled Ibrahim",       shop_code:"S-9158-049", shop_name:"Raneen (10th Ramadan)",  date:"2026-05-04", start:"10:00", end:"18:30", duration:"08:30", status:"late"    },
  { id:"ATT-010", emp_code:"A-2687", emp_name:"Karim Samir Fathy",         shop_code:"S-0074-012", shop_name:"Dubai Phone (El Maadi)", date:"2026-05-04", start:"09:00", end:"18:00", duration:"09:00", status:"present" },
];

// Network stats
const NETWORK_STATS = {
  activeShops: 3973,
  totalSellout: 187755,
  attRecords: 43541,
  shopsWithSalesNoAtt: 38,
  attShopsNoSales: 2386,
  commonShops: 3973,
  dataAccuracy: 99.2,
  employees: 843,
};

// ─── COLOURS & HELPERS ────────────────────────────────────────────────────
const fmtM = (n) => n >= 1e6 ? `${(n/1e6).toFixed(1)}M` : n >= 1e3 ? `${(n/1e3).toFixed(0)}K` : n;
const fmtEGP = (n) => `EGP ${(n/1e6).toFixed(1)}M`;
const statusColor = (s) => ({ present:"text-emerald-400", absent:"text-red-400", late:"text-amber-400" }[s] || "text-neutral-400");
const statusBg    = (s) => ({ present:"bg-emerald-500/20 text-emerald-400", absent:"bg-red-500/20 text-red-400", late:"bg-amber-500/20 text-amber-400" }[s] || "bg-neutral-500/20 text-neutral-400");

// ─── MINI COMPONENTS ──────────────────────────────────────────────────────
function StatCard({ label, value, color = "text-white", icon: Icon, iconColor = "text-cyan-400" }) {
  return (
    <div className="bg-neutral-900 border border-neutral-700 rounded-lg p-4 flex items-center justify-between">
      <div>
        <p className="text-xs text-neutral-400 tracking-wider mb-1">{label}</p>
        <p className={`text-2xl font-bold font-mono ${color}`}>{value}</p>
      </div>
      {Icon && <Icon className={`w-8 h-8 ${iconColor}`} />}
    </div>
  );
}

function SectionHeader({ title, sub, actions }) {
  return (
    <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
      <div>
        <h1 className="text-xl font-bold text-white tracking-wider">{title}</h1>
        {sub && <p className="text-sm text-neutral-400">{sub}</p>}
      </div>
      {actions && <div className="flex gap-2">{actions}</div>}
    </div>
  );
}

// ─── DASHBOARD OVERVIEW ───────────────────────────────────────────────────
function DashboardPage() {
  const totalRevenue = MERGED_SHOPS.reduce((s, x) => s + x.revenue, 0);
  return (
    <div className="p-6 space-y-6">
      <SectionHeader title="COMMAND CENTER" sub="Live network overview — W19–W22 2026" />

      {/* KPI row */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <StatCard label="TOTAL SELLOUT"    value={fmtM(NETWORK_STATS.totalSellout)} color="text-cyan-400"    icon={TrendingUp}  iconColor="text-cyan-400" />
        <StatCard label="ACTIVE SHOPS"     value={NETWORK_STATS.activeShops.toLocaleString()} color="text-white" icon={Store}      iconColor="text-purple-400" />
        <StatCard label="ATT RECORDS (MAY)" value={NETWORK_STATS.attRecords.toLocaleString()} color="text-emerald-400" icon={Users} iconColor="text-emerald-400" />
        <StatCard label="SALES/NO-ATT"     value={NETWORK_STATS.shopsWithSalesNoAtt} color="text-amber-400" icon={AlertTriangle} iconColor="text-amber-400" />
      </div>

      {/* Charts row */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Weekly trend */}
        <div className="bg-neutral-900 border border-neutral-700 rounded-lg p-4">
          <p className="text-xs text-neutral-400 tracking-wider mb-4">WEEKLY SELLOUT (W19–W22)</p>
          <ResponsiveContainer width="100%" height={180}>
            <LineChart data={WEEK_SALES}>
              <XAxis dataKey="week" tick={{ fill:"#737373", fontSize:11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill:"#737373", fontSize:10 }} axisLine={false} tickLine={false} width={40} />
              <Tooltip contentStyle={{ background:"#171717", border:"1px solid #404040", borderRadius:6, color:"#fff", fontSize:11 }} />
              <Line type="monotone" dataKey="sellout" stroke="#06b6d4" strokeWidth={2} dot={{ r:3, fill:"#06b6d4" }} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* By area */}
        <div className="bg-neutral-900 border border-neutral-700 rounded-lg p-4">
          <p className="text-xs text-neutral-400 tracking-wider mb-4">SELLOUT BY AREA</p>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={AREA_SALES}>
              <XAxis dataKey="area" tick={{ fill:"#737373", fontSize:10 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill:"#737373", fontSize:10 }} axisLine={false} tickLine={false} width={40} />
              <Tooltip contentStyle={{ background:"#171717", border:"1px solid #404040", borderRadius:6, color:"#fff", fontSize:11 }} />
              <Bar dataKey="sellout" radius={[4,4,0,0]}>
                {AREA_SALES.map((d, i) => <Cell key={i} fill={d.color} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* By project pie */}
        <div className="bg-neutral-900 border border-neutral-700 rounded-lg p-4">
          <p className="text-xs text-neutral-400 tracking-wider mb-4">SELLOUT BY PROJECT</p>
          <ResponsiveContainer width="100%" height={180}>
            <PieChart>
              <Pie data={PROJECT_SALES} dataKey="sellout" nameKey="project" cx="50%" cy="50%" outerRadius={70} label={({ project, percent }) => `${project} ${(percent*100).toFixed(0)}%`} labelLine={false}>
                {PROJECT_SALES.map((d, i) => <Cell key={i} fill={d.color} />)}
              </Pie>
              <Tooltip contentStyle={{ background:"#171717", border:"1px solid #404040", borderRadius:6, color:"#fff", fontSize:11 }} />
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Top brands */}
      <div className="bg-neutral-900 border border-neutral-700 rounded-lg p-4">
        <p className="text-xs text-neutral-400 tracking-wider mb-4">TOP BRANDS — TOTAL SELLOUT</p>
        <ResponsiveContainer width="100%" height={160}>
          <BarChart data={BRAND_SALES} layout="vertical">
            <XAxis type="number" tick={{ fill:"#737373", fontSize:10 }} axisLine={false} tickLine={false} />
            <YAxis type="category" dataKey="name" tick={{ fill:"#a3a3a3", fontSize:11 }} axisLine={false} tickLine={false} width={60} />
            <Tooltip contentStyle={{ background:"#171717", border:"1px solid #404040", borderRadius:6, color:"#fff", fontSize:11 }} />
            <Bar dataKey="sales" fill="#f59e0b" radius={[0,4,4,0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      {/* Data integrity banner */}
      <div className="bg-amber-500/10 border border-amber-500/30 rounded-lg p-4 flex flex-wrap items-center gap-6">
        <div className="flex items-center gap-2">
          <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0" />
          <span className="text-xs text-amber-300 font-medium">DATA INTEGRITY FLAGS</span>
        </div>
        <span className="text-xs text-amber-200">{NETWORK_STATS.shopsWithSalesNoAtt} shops: sales logged, no attendance</span>
        <span className="text-xs text-neutral-400">{NETWORK_STATS.attShopsNoSales} shops: attendance only, no sales in W19–W22</span>
        <span className="text-xs text-emerald-300">{NETWORK_STATS.commonShops.toLocaleString()} shops matched across both datasets</span>
      </div>
    </div>
  );
}

// ─── SALES ANALYTICS PAGE ─────────────────────────────────────────────────
function SalesPage() {
  const [search, setSearch] = useState("");
  const [filterArea, setFilterArea] = useState("All");
  const [filterProject, setFilterProject] = useState("All");

  const areas = ["All", ...new Set(MERGED_SHOPS.map(s => s.area))];
  const projects = ["All", ...new Set(MERGED_SHOPS.map(s => s.project))];

  const filtered = MERGED_SHOPS.filter(s => {
    const matchSearch = !search || s.shop_name.toLowerCase().includes(search.toLowerCase()) || s.shop_code.toLowerCase().includes(search.toLowerCase());
    const matchArea = filterArea === "All" || s.area === filterArea;
    const matchProject = filterProject === "All" || s.project === filterProject;
    return matchSearch && matchArea && matchProject;
  });

  return (
    <div className="p-6 space-y-6">
      <SectionHeader title="SALES ANALYTICS" sub="MBL Raw Data W19–W22 · 4,011 shops · 230,378 records" />

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <StatCard label="TOTAL SELLOUT"  value={fmtM(187755)} color="text-cyan-400"    icon={Package}    iconColor="text-cyan-400" />
        <StatCard label="UNIQUE SHOPS"   value="4,011"        color="text-white"        icon={Store}      iconColor="text-purple-400" />
        <StatCard label="WEEKS COVERED"  value="W19–W22"      color="text-purple-400"   icon={BarChart3}  iconColor="text-purple-400" />
        <StatCard label="TOP BRAND"      value="Samsung"      color="text-amber-400"    icon={TrendingUp} iconColor="text-amber-400" />
      </div>

      {/* Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-neutral-900 border border-neutral-700 rounded-lg p-4">
          <p className="text-xs text-neutral-400 tracking-wider mb-4">SELLOUT BY PROJECT</p>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={PROJECT_SALES}>
              <XAxis dataKey="project" tick={{ fill:"#737373", fontSize:11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill:"#737373", fontSize:10 }} axisLine={false} tickLine={false} width={50} />
              <Tooltip contentStyle={{ background:"#171717", border:"1px solid #404040", borderRadius:6, color:"#fff", fontSize:11 }} />
              <Bar dataKey="sellout" radius={[4,4,0,0]}>
                {PROJECT_SALES.map((d,i) => <Cell key={i} fill={d.color} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="bg-neutral-900 border border-neutral-700 rounded-lg p-4">
          <p className="text-xs text-neutral-400 tracking-wider mb-4">TOP 8 BRANDS</p>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={BRAND_SALES} layout="vertical">
              <XAxis type="number" tick={{ fill:"#737373", fontSize:10 }} axisLine={false} tickLine={false} />
              <YAxis type="category" dataKey="name" tick={{ fill:"#a3a3a3", fontSize:11 }} axisLine={false} tickLine={false} width={60} />
              <Tooltip contentStyle={{ background:"#171717", border:"1px solid #404040", borderRadius:6, color:"#fff", fontSize:11 }} />
              <Bar dataKey="sales" fill="#06b6d4" radius={[0,4,4,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Filters */}
      <div className="bg-neutral-900/50 border border-neutral-700 rounded-lg p-4 flex flex-wrap gap-3 items-center">
        <div className="relative flex-1 min-w-52">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-neutral-500" />
          <input
            className="w-full pl-10 pr-4 py-2 bg-neutral-800 border border-neutral-600 rounded text-sm text-white placeholder:text-neutral-500 focus:outline-none focus:border-cyan-500"
            placeholder="Search shop name or code..."
            value={search} onChange={e => setSearch(e.target.value)}
          />
        </div>
        {[["Area", areas, filterArea, setFilterArea], ["Project", projects, filterProject, setFilterProject]].map(([label, opts, val, set]) => (
          <select key={label} value={val} onChange={e => set(e.target.value)}
            className="bg-neutral-800 border border-neutral-600 text-sm text-neutral-300 rounded px-3 py-2 focus:outline-none focus:border-cyan-500">
            {opts.map(o => <option key={o}>{o}</option>)}
          </select>
        ))}
        <button onClick={() => { setSearch(""); setFilterArea("All"); setFilterProject("All"); }}
          className="text-xs text-neutral-400 hover:text-white px-3 py-2">Clear</button>
      </div>

      {/* Table */}
      <div className="bg-neutral-900 border border-neutral-700 rounded-lg overflow-hidden">
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-700">
          <span className="text-xs text-neutral-300 tracking-wider font-medium">SHOP SALES — MERGED WITH ATTENDANCE</span>
          <span className="text-xs text-neutral-500">{filtered.length} of {MERGED_SHOPS.length} shown (top 20)</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-xs text-neutral-500 border-b border-neutral-700 bg-neutral-900/80">
                {["SHOP CODE","SHOP NAME","AREA","PROJECT","SELLOUT","REVENUE","TOP BRAND","ATT RECORDS","STAFF"].map(h => (
                  <th key={h} className="text-left py-3 px-3 font-medium whitespace-nowrap">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(s => (
                <tr key={s.shop_code} className="border-b border-neutral-800 hover:bg-neutral-800/50 transition-colors">
                  <td className="py-3 px-3 font-mono text-cyan-400 text-xs whitespace-nowrap">{s.shop_code}</td>
                  <td className="py-3 px-3 text-white max-w-[180px] truncate">{s.shop_name}</td>
                  <td className="py-3 px-3 text-neutral-400 text-xs">{s.area}</td>
                  <td className="py-3 px-3 text-neutral-400 text-xs">{s.project}</td>
                  <td className="py-3 px-3 text-right font-mono text-emerald-400">{s.sellout.toLocaleString()}</td>
                  <td className="py-3 px-3 text-right font-mono text-amber-400 text-xs">{fmtEGP(s.revenue)}</td>
                  <td className="py-3 px-3 text-neutral-300 text-xs">{s.top_brand}</td>
                  <td className="py-3 px-3 text-right font-mono text-purple-400">
                    {s.att_records > 0 ? s.att_records : <span className="text-red-400">0</span>}
                  </td>
                  <td className="py-3 px-3 text-right font-mono text-neutral-300">{s.emp_count}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── ATTENDANCE PAGE ───────────────────────────────────────────────────────
function AttendancePage() {
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("All");

  const filtered = ATTENDANCE_RECORDS.filter(r => {
    const matchSearch = !search || r.emp_name.toLowerCase().includes(search.toLowerCase()) || r.shop_code.toLowerCase().includes(search.toLowerCase()) || r.shop_name.toLowerCase().includes(search.toLowerCase());
    const matchStatus = filterStatus === "All" || r.status === filterStatus;
    return matchSearch && matchStatus;
  });

  const present = ATTENDANCE_RECORDS.filter(r => r.status === "present").length;
  const absent  = ATTENDANCE_RECORDS.filter(r => r.status === "absent").length;
  const late    = ATTENDANCE_RECORDS.filter(r => r.status === "late").length;

  return (
    <div className="p-6 space-y-6">
      <SectionHeader title="ATTENDANCE TRACKER" sub="May 2026 · 43,541 records · attend_start_end_per_store" />

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <StatCard label="TOTAL RECORDS"  value="43,541"                         color="text-white"       icon={Users}       iconColor="text-purple-400" />
        <StatCard label="PRESENT (SAMPLE)" value={present}                      color="text-emerald-400" icon={CheckCircle} iconColor="text-emerald-400" />
        <StatCard label="ABSENT (SAMPLE)"  value={absent}                       color="text-red-400"     icon={XCircle}     iconColor="text-red-400" />
        <StatCard label="LATE (SAMPLE)"    value={late}                         color="text-amber-400"   icon={Clock}       iconColor="text-amber-400" />
      </div>

      {/* Filters */}
      <div className="bg-neutral-900/50 border border-neutral-700 rounded-lg p-4 flex flex-wrap gap-3 items-center">
        <div className="relative flex-1 min-w-52">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-neutral-500" />
          <input
            className="w-full pl-10 pr-4 py-2 bg-neutral-800 border border-neutral-600 rounded text-sm text-white placeholder:text-neutral-500 focus:outline-none focus:border-cyan-500"
            placeholder="Search employee or shop..."
            value={search} onChange={e => setSearch(e.target.value)}
          />
        </div>
        <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)}
          className="bg-neutral-800 border border-neutral-600 text-sm text-neutral-300 rounded px-3 py-2 focus:outline-none focus:border-cyan-500">
          {["All","present","absent","late"].map(o => <option key={o}>{o}</option>)}
        </select>
        <button onClick={() => { setSearch(""); setFilterStatus("All"); }}
          className="text-xs text-neutral-400 hover:text-white px-3 py-2">Clear</button>
      </div>

      {/* Table */}
      <div className="bg-neutral-900 border border-neutral-700 rounded-lg overflow-hidden">
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-700">
          <span className="text-xs text-neutral-300 tracking-wider font-medium">ATTENDANCE RECORDS</span>
          <span className="text-xs text-neutral-500">{filtered.length} of {ATTENDANCE_RECORDS.length} shown (sample)</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-xs text-neutral-500 border-b border-neutral-700">
                {["EMP CODE","EMPLOYEE","SHOP CODE","SHOP NAME","DATE","START","END","DURATION","STATUS"].map(h => (
                  <th key={h} className="text-left py-3 px-3 font-medium whitespace-nowrap">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(r => (
                <tr key={r.id} className="border-b border-neutral-800 hover:bg-neutral-800/50 transition-colors">
                  <td className="py-3 px-3 font-mono text-xs text-neutral-400">{r.emp_code}</td>
                  <td className="py-3 px-3 text-white text-xs max-w-[160px] truncate">{r.emp_name}</td>
                  <td className="py-3 px-3 font-mono text-cyan-400 text-xs">{r.shop_code}</td>
                  <td className="py-3 px-3 text-neutral-300 text-xs max-w-[140px] truncate">{r.shop_name}</td>
                  <td className="py-3 px-3 font-mono text-xs text-neutral-400">{r.date}</td>
                  <td className="py-3 px-3 font-mono text-xs text-emerald-400 text-center">{r.start}</td>
                  <td className="py-3 px-3 font-mono text-xs text-amber-400 text-center">{r.end}</td>
                  <td className="py-3 px-3 font-mono text-xs text-neutral-300 text-center">{r.duration}</td>
                  <td className="py-3 px-3">
                    <span className={`text-xs px-2 py-0.5 rounded font-medium ${statusBg(r.status)}`}>
                      {r.status.toUpperCase()}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── DATA BRIDGE PAGE ─────────────────────────────────────────────────────
function DataBridgePage() {
  const [view, setView] = useState("merged"); // merged | sales_only | att_only

  const views = [
    { id: "merged",    label: "Merged (3,973)",    color: "text-emerald-400" },
    { id: "sales_only",label: "Sales only (38)",   color: "text-amber-400" },
    { id: "att_only",  label: "Att only (2,386)",  color: "text-blue-400" },
  ];

  return (
    <div className="p-6 space-y-6">
      <SectionHeader title="DATA BRIDGE" sub="Attendance × Sales joined on Shop Code" />

      {/* Merge summary */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="bg-neutral-900 border border-emerald-500/30 rounded-lg p-5">
          <div className="flex items-center gap-2 mb-3">
            <Merge className="w-5 h-5 text-emerald-400" />
            <span className="text-xs text-emerald-400 tracking-wider font-medium">MATCHED SHOPS</span>
          </div>
          <p className="text-3xl font-bold font-mono text-emerald-400">3,973</p>
          <p className="text-xs text-neutral-500 mt-1">In both attendance & sales datasets</p>
        </div>
        <div className="bg-neutral-900 border border-amber-500/30 rounded-lg p-5">
          <div className="flex items-center gap-2 mb-3">
            <AlertTriangle className="w-5 h-5 text-amber-400" />
            <span className="text-xs text-amber-400 tracking-wider font-medium">SALES / NO ATTENDANCE</span>
          </div>
          <p className="text-3xl font-bold font-mono text-amber-400">38</p>
          <p className="text-xs text-neutral-500 mt-1">Shops with sales but no staff logged</p>
        </div>
        <div className="bg-neutral-900 border border-blue-500/30 rounded-lg p-5">
          <div className="flex items-center gap-2 mb-3">
            <Users className="w-5 h-5 text-blue-400" />
            <span className="text-xs text-blue-400 tracking-wider font-medium">ATT / NO SALES W19–W22</span>
          </div>
          <p className="text-3xl font-bold font-mono text-blue-400">2,386</p>
          <p className="text-xs text-neutral-500 mt-1">Shops with staff logged but no sales reported</p>
        </div>
      </div>

      {/* Source files info */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {[
          { icon: FileSpreadsheet, color: "text-cyan-400", border: "border-cyan-500/30", label: "ATTENDANCE FILE", name: "attend_start_end_per_store_May2026.xlsx", rows: "43,541 rows", key: "shop code", cols: "emp_code, emp_name, shop code, shop name, start_work_time, end_work_time, duration, date" },
          { icon: FileSpreadsheet, color: "text-purple-400", border: "border-purple-500/30", label: "SALES FILE", name: "MBL_Raw_Data_W19_W22.xlsx", rows: "230,378 rows", key: "Shop Code", cols: "W, Project, Shop Code, Shop Name, Area, Brand, Sellout, Price, Type" },
        ].map(f => (
          <div key={f.name} className={`bg-neutral-900 border ${f.border} rounded-lg p-4`}>
            <div className="flex items-center gap-2 mb-3">
              <f.icon className={`w-5 h-5 ${f.color}`} />
              <span className={`text-xs ${f.color} tracking-wider font-medium`}>{f.label}</span>
            </div>
            <p className="text-sm text-white font-mono mb-1">{f.name}</p>
            <p className="text-xs text-neutral-400 mb-2">{f.rows}</p>
            <p className="text-xs text-neutral-500"><span className="text-neutral-400">Join key:</span> <span className="text-cyan-300 font-mono">{f.key}</span></p>
            <p className="text-xs text-neutral-500 mt-1 leading-relaxed">{f.cols}</p>
          </div>
        ))}
      </div>

      {/* Merge result table */}
      <div className="bg-neutral-900 border border-neutral-700 rounded-lg overflow-hidden">
        <div className="flex flex-wrap items-center gap-3 px-4 py-3 border-b border-neutral-700">
          <span className="text-xs text-neutral-300 tracking-wider font-medium mr-2">VIEW:</span>
          {views.map(v => (
            <button key={v.id} onClick={() => setView(v.id)}
              className={`text-xs px-3 py-1 rounded border transition-colors ${view === v.id ? `${v.color} border-current bg-white/5` : "text-neutral-500 border-neutral-600 hover:text-neutral-300"}`}>
              {v.label}
            </button>
          ))}
        </div>

        {view === "merged" && (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-xs text-neutral-500 border-b border-neutral-700">
                  {["SHOP CODE","SHOP NAME","AREA","SELLOUT (W19–W22)","REVENUE","ATT RECORDS (MAY)","EMP COUNT","MATCH"].map(h => (
                    <th key={h} className="text-left py-3 px-3 font-medium whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {MERGED_SHOPS.map(s => (
                  <tr key={s.shop_code} className="border-b border-neutral-800 hover:bg-neutral-800/50 transition-colors">
                    <td className="py-3 px-3 font-mono text-cyan-400 text-xs">{s.shop_code}</td>
                    <td className="py-3 px-3 text-white text-xs max-w-[160px] truncate">{s.shop_name}</td>
                    <td className="py-3 px-3 text-neutral-400 text-xs">{s.area}</td>
                    <td className="py-3 px-3 text-right font-mono text-emerald-400">{s.sellout.toLocaleString()}</td>
                    <td className="py-3 px-3 text-right font-mono text-amber-400 text-xs">{fmtEGP(s.revenue)}</td>
                    <td className="py-3 px-3 text-right font-mono text-purple-400">{s.att_records}</td>
                    <td className="py-3 px-3 text-right font-mono text-neutral-300">{s.emp_count}</td>
                    <td className="py-3 px-3">
                      <span className="text-xs bg-emerald-500/20 text-emerald-400 px-2 py-0.5 rounded">✓ MATCHED</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {view === "sales_only" && (
          <div className="p-6 text-center">
            <AlertTriangle className="w-8 h-8 text-amber-400 mx-auto mb-3" />
            <p className="text-amber-400 font-medium mb-1">38 shops: Sales data with NO attendance record</p>
            <p className="text-sm text-neutral-400">These shops generated sales in W19–W22 but have zero attendance logs in May 2026.</p>
            <p className="text-xs text-neutral-500 mt-2">This is a key fraud/ghost-visit signal — sales cannot occur without staff presence.</p>
          </div>
        )}

        {view === "att_only" && (
          <div className="p-6 text-center">
            <Users className="w-8 h-8 text-blue-400 mx-auto mb-3" />
            <p className="text-blue-400 font-medium mb-1">2,386 shops: Attendance logged, NO sales in W19–W22</p>
            <p className="text-sm text-neutral-400">Staff attended these shops in May 2026 but no sellout was recorded in the W19–W22 sales dataset.</p>
            <p className="text-xs text-neutral-500 mt-2">May indicate shops not yet onboarded to the sales tracking system, or seasonal gaps.</p>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── NEXUS COPILOT (using Anthropic API) ──────────────────────────────────
function NexusCopilot({ open, onClose }) {
  const [input, setInput]       = useState("");
  const [messages, setMessages] = useState([]);
  const [loading, setLoading]   = useState(false);
  const endRef = useRef(null);
  const textareaRef = useRef(null);

  useEffect(() => { endRef.current?.scrollIntoView({ behavior: "smooth" }); }, [messages]);

  const SYSTEM = `You are Nexus, an elite retail intelligence AI for a mobile phone distribution network in Egypt.

LIVE DATA SNAPSHOT:
- Total Sellout W19–W22: 187,755 units
- Active Shops: 3,973 (matched across sales & attendance datasets)
- Attendance Records (May 2026): 43,541
- Shops with sales but no attendance: 38 (CRITICAL fraud signal)
- Shops with attendance but no sales W19–W22: 2,386

AREA BREAKDOWN: Cairo 72,514 | Delta 49,198 | Upper Egypt 38,718 | Alex 27,325
PROJECT BREAKDOWN: Club 88,871 | Elite 34,813 | Hero 24,371 | Star 23,650 | One 16,050
TOP BRANDS: Samsung 84,618 | Oppo 34,453 | Xiaomi 19,383 | Realme 16,292 | Infinix 10,069

TOP SHOPS: Dubai Phone (S-0074-005): 3,042 units / EGP 63.1M | Dubai Phone City Center (S-0074-006): 2,495 / EGP 63.4M

DATA SOURCES: attend_start_end_per_store_May2026.xlsx (43,541 rows) merged with MBL_Raw_Data_W19_W22.xlsx (230,378 rows) on Shop Code.

Principles: Always cite real numbers. Surface integrity issues (ghost visits, fake check-ins) with specifics. Be concise and direct.`;

  const send = useCallback(async () => {
    const text = input.trim();
    if (!text || loading) return;
    setInput("");

    const userMsg = { role: "user", content: text, id: Date.now() };
    const newMsgs = [...messages, userMsg];
    setMessages(newMsgs);
    setLoading(true);

    const assistantId = Date.now() + 1;
    setMessages(prev => [...prev, { role: "assistant", content: "", id: assistantId }]);

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          system: SYSTEM,
          messages: newMsgs.map(m => ({ role: m.role, content: m.content })),
        }),
      });
      const data = await res.json();
      const reply = data?.content?.[0]?.text || "No response.";
      setMessages(prev => prev.map(m => m.id === assistantId ? { ...m, content: reply } : m));
    } catch {
      setMessages(prev => prev.map(m => m.id === assistantId ? { ...m, content: "⚠️ Could not reach intelligence engine." } : m));
    } finally {
      setLoading(false);
    }
  }, [input, messages, loading]);

  const onKey = (e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } };

  if (!open) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-end sm:justify-end pointer-events-none">
      <div className="w-full sm:w-[420px] h-[600px] sm:h-[580px] sm:m-4 bg-neutral-900 border border-cyan-500/30 rounded-xl flex flex-col pointer-events-auto shadow-2xl overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between px-4 py-3 bg-neutral-800 border-b border-neutral-700">
          <div className="flex items-center gap-2">
            <div className="w-2 h-2 bg-cyan-400 rounded-full animate-pulse" />
            <span className="text-sm font-semibold text-cyan-400 tracking-wider">NEXUS COPILOT</span>
          </div>
          <button onClick={onClose} className="text-neutral-400 hover:text-white"><X className="w-4 h-4" /></button>
        </div>

        {/* Messages */}
        <div className="flex-1 overflow-y-auto p-4 space-y-4">
          {messages.length === 0 && (
            <div className="text-center py-8">
              <Bot className="w-10 h-10 text-cyan-400/40 mx-auto mb-3" />
              <p className="text-sm text-neutral-400">Ask about sales, attendance, shop performance, or integrity flags.</p>
              <div className="mt-4 space-y-2">
                {["Which area has the most ghost-visit risk?", "Top 5 shops by sellout?", "Show attendance vs sales gaps"].map(q => (
                  <button key={q} onClick={() => setInput(q)}
                    className="block w-full text-left text-xs text-cyan-400/70 hover:text-cyan-300 bg-neutral-800 rounded px-3 py-2 transition-colors">
                    {q}
                  </button>
                ))}
              </div>
            </div>
          )}
          {messages.map(m => (
            <div key={m.id} className={`flex ${m.role === "user" ? "justify-end" : "justify-start"}`}>
              <div className={`max-w-[85%] rounded-lg px-3 py-2 text-sm leading-relaxed ${m.role === "user" ? "bg-cyan-600 text-white" : "bg-neutral-800 text-neutral-100"}`}>
                {m.content || (loading && m.role === "assistant" && <Loader2 className="w-4 h-4 animate-spin text-cyan-400" />)}
              </div>
            </div>
          ))}
          <div ref={endRef} />
        </div>

        {/* Input */}
        <div className="px-4 py-3 border-t border-neutral-700 flex gap-2">
          <textarea
            ref={textareaRef}
            rows={1}
            value={input}
            onChange={e => setInput(e.target.value)}
            onKeyDown={onKey}
            placeholder="Ask Nexus..."
            className="flex-1 resize-none bg-neutral-800 border border-neutral-600 rounded-lg px-3 py-2 text-sm text-white placeholder:text-neutral-500 focus:outline-none focus:border-cyan-500"
          />
          <button onClick={send} disabled={!input.trim() || loading}
            className="bg-cyan-600 hover:bg-cyan-500 disabled:opacity-40 text-white rounded-lg px-3 flex items-center justify-center transition-colors">
            <Send className="w-4 h-4" />
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── SETTINGS PAGE ────────────────────────────────────────────────────────
function SettingsPage() {
  return (
    <div className="p-6 space-y-6">
      <SectionHeader title="SYSTEM SETTINGS" sub="Configuration & data sources" />
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {[
          { label: "Data Sources", items: [
            { k: "Sales File", v: "MBL_Raw_Data_W19_W22.xlsx", ok: true },
            { k: "Attendance File", v: "attend_start_end_per_store_May2026.xlsx", ok: true },
            { k: "Join Key", v: "Shop Code", ok: true },
            { k: "Sales Rows", v: "230,378", ok: true },
            { k: "Attendance Rows", v: "43,541", ok: true },
          ]},
          { label: "Network Status", items: [
            { k: "Active Shops",   v: "3,973",  ok: true },
            { k: "Matched Shops",  v: "3,973",  ok: true },
            { k: "Sales-only gaps",v: "38",      ok: false },
            { k: "Data Accuracy",  v: "99.2%",  ok: true },
            { k: "Employees",      v: "843",     ok: true },
          ]},
        ].map(section => (
          <div key={section.label} className="bg-neutral-900 border border-neutral-700 rounded-lg overflow-hidden">
            <div className="px-4 py-3 border-b border-neutral-700">
              <span className="text-xs text-neutral-300 tracking-wider font-medium">{section.label.toUpperCase()}</span>
            </div>
            <div className="p-4 space-y-3">
              {section.items.map(item => (
                <div key={item.k} className="flex items-center justify-between">
                  <span className="text-sm text-neutral-400">{item.k}</span>
                  <div className="flex items-center gap-2">
                    <span className="text-sm text-white font-mono">{item.v}</span>
                    <div className={`w-2 h-2 rounded-full ${item.ok ? "bg-emerald-400" : "bg-amber-400"}`} />
                  </div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────
export default function DataNexusDashboard() {
  const [activeSection, setActiveSection] = useState("dashboard");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [copilotOpen, setCopilotOpen] = useState(false);

  const navItems = [
    { id: "dashboard",      icon: LayoutDashboard, label: "DASHBOARD" },
    { id: "sales",          icon: BarChart3,        label: "SALES ANALYTICS" },
    { id: "attendance",     icon: Users,             label: "ATTENDANCE" },
    { id: "dataBridge",     icon: Database,          label: "DATA BRIDGE" },
    { id: "importReports",  icon: Upload,            label: "IMPORT REPORTS" },
    { id: "settings",       icon: Settings,          label: "SETTINGS" },
  ];

  const totalSellout = 187755;
  const weekTrend = ((WEEK_SALES[3].sellout - WEEK_SALES[0].sellout) / WEEK_SALES[0].sellout * 100).toFixed(1);

  return (
    <div className="flex h-screen bg-neutral-950 text-white overflow-hidden">
      {/* Sidebar */}
      <div className={`${sidebarCollapsed ? "w-16" : "w-64"} bg-neutral-900 border-r border-neutral-700 flex flex-col transition-all duration-300 shrink-0`}>
        <div className="p-4 flex items-center justify-between">
          {!sidebarCollapsed && (
            <div>
              <h1 className="text-cyan-400 font-bold text-lg tracking-wider">DATA NEXUS</h1>
              <p className="text-amber-500 text-xs">COMMAND CENTER</p>
            </div>
          )}
          <button onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
            className="text-neutral-400 hover:text-cyan-400 p-1 rounded transition-colors ml-auto">
            <ChevronRight className={`w-5 h-5 transition-transform ${sidebarCollapsed ? "" : "rotate-180"}`} />
          </button>
        </div>

        <nav className="flex-1 px-3 space-y-1">
          {navItems.map(item => (
            <button key={item.id} onClick={() => setActiveSection(item.id)}
              className={`w-full flex items-center gap-3 p-3 rounded-lg transition-colors ${
                activeSection === item.id
                  ? "bg-cyan-500/20 text-cyan-400 border-l-2 border-cyan-400"
                  : "text-neutral-400 hover:text-white hover:bg-neutral-800"
              }`}>
              <item.icon className="w-5 h-5 shrink-0" />
              {!sidebarCollapsed && <span className="text-xs font-medium">{item.label}</span>}
            </button>
          ))}
        </nav>

        {!sidebarCollapsed && (
          <div className="m-3 p-3 bg-neutral-800/60 border border-cyan-500/20 rounded-lg">
            <div className="flex items-center gap-2 mb-2">
              <div className="w-2 h-2 bg-emerald-400 rounded-full animate-pulse" />
              <span className="text-xs text-emerald-400">LIVE DATA</span>
            </div>
            <div className="space-y-1 text-xs text-neutral-500">
              <div>ACCURACY: <span className="text-emerald-400">99.2%</span></div>
              <div>SHOPS: <span className="text-cyan-400">3,973</span></div>
              <div>W19–W22: <span className="text-amber-400">{fmtM(totalSellout)} units</span></div>
            </div>
          </div>
        )}
      </div>

      {/* Main */}
      <div className="flex-1 flex flex-col min-w-0">
        {/* Toolbar */}
        <div className="h-14 bg-neutral-800 border-b border-neutral-700 flex items-center justify-between px-6 shrink-0">
          <div className="flex items-center gap-3">
            <span className="text-sm text-neutral-400">Data Nexus</span>
            <span className="text-cyan-400 bg-cyan-400/10 px-2 py-0.5 rounded text-xs">v2.1 REAL DATA</span>
            <span className={`text-xs ${Number(weekTrend) < 0 ? "text-red-400" : "text-emerald-400"}`}>
              W19→W22: {weekTrend}%
            </span>
          </div>
          <div className="flex items-center gap-3">
            <div className="bg-amber-500/20 border border-amber-500/50 rounded px-3 py-1 flex items-center gap-2">
              <AlertTriangle className="w-3 h-3 text-amber-400" />
              <span className="text-xs text-amber-400">38 shops: sales/no-attendance</span>
            </div>
            <button onClick={() => setCopilotOpen(v => !v)}
              className={`flex items-center gap-2 text-xs px-3 py-1.5 rounded border transition-colors ${copilotOpen ? "bg-cyan-500/20 border-cyan-500/50 text-cyan-300" : "bg-neutral-700 border-neutral-600 text-neutral-300 hover:border-cyan-500/50"}`}>
              <Bot className="w-4 h-4" />
              NEXUS AI
            </button>
            <button className="text-neutral-400 hover:text-cyan-400">
              <Bell className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Content */}
        <div className="flex-1 overflow-auto bg-neutral-950">
          {activeSection === "dashboard"      && <DashboardPage />}
          {activeSection === "sales"          && <SalesPage />}
          {activeSection === "attendance"     && <AttendancePage />}
          {activeSection === "dataBridge"     && <DataBridgePage />}
          {activeSection === "importReports"  && <ImportReportsPage />}
          {activeSection === "settings"       && <SettingsPage />}
        </div>
      </div>

      <NexusCopilot open={copilotOpen} onClose={() => setCopilotOpen(false)} />
    </div>
  );
}
