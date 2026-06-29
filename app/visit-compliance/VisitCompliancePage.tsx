"use client";
import { useState, useCallback, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Upload, FileSpreadsheet, CheckCircle, AlertTriangle, XCircle,
  Users, ShieldCheck, ShieldAlert, TrendingDown, TrendingUp,
  Layers, RefreshCw, Crown, Flag, Copy, FileWarning, Eye, Gauge,
  Zap, Clock, Target, Activity, Ghost, MapPin, Download, X,
  FileDown, MessageSquare, ChevronRight, Mail, Printer, FastForward, Radar, Search, BarChart3, UserCheck, CheckSquare, Square, Send, Navigation, Monitor, Calendar
} from "lucide-react";
import {
  AreaChart, Area, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid, ReferenceLine,
  LabelList, LineChart, Line
} from "recharts";
import { MapContainer, TileLayer, CircleMarker, Tooltip as LeafletTooltip } from 'react-leaflet';
import 'leaflet/dist/leaflet.css';
import html2canvas from 'html2canvas';

// ─── GLOBAL STYLES FOR ANIMATIONS ───────────────────────────────────────────
const pulseStyle = `
@keyframes pulse-led {
  0% { box-shadow: 0 0 0 0 rgba(248, 113, 113, 0.7); }
  70% { box-shadow: 0 0 0 10px rgba(248, 113, 113, 0); }
  100% { box-shadow: 0 0 0 0 rgba(248, 113, 113, 0); }
}
.pulse-led { animation: pulse-led 2s infinite; }
`;

// ─── TYPES ──────────────────────────────────────────────────────────────────
type VisitFileKind = "visit_raw" | "visit_followup" | "visit_summary" | "unknown";
interface LoadedFile { name: string; kind: VisitFileKind; size: string; }
interface EmployeeRecord { code: string; name: string; department: string; daysReported: number; daysPossible: number; }
interface VisitRow { date: string; spvrCode: string; spvrName: string; shopCode: string; shopName: string; area: string; governorate: string; samsungShortage: string; compShortage: string; selloutMovement: string; brand: string; movement: string; comment: string; action1: string; accountFeedback: string; action2: string; lat: number; lon: number; }
interface IntegrityProfile { code: string; name: string; totalVisits: number; uniqueShops: number; commentedVisits: number; blankComments: number; lowInfoComments: number; templatedRepeats: number; topRepeatedComment: { text: string; count: number } | null; uniquenessRatio: number; unsupportedClaims: number; integrityScore: number; avgCommentLen: number; lateNightCount: number; singleShopLoop: boolean; ghostScore: number; suspiciousFlags: string[]; teleportationFlags: string[]; rushHourFlags: string[]; commentTopics: string[]; avgSentiment: number; mutatedCopyPastes: number; dailyTrend: { date: string; count: number }[]; }

const LOW_INFO_SET = new Set(["ok","fine","good","done","no issue","nothing","n/a","na","-","none","good.","ok."]);
const fmtBytes = (b: number) => b > 1_048_576 ? `${(b/1_048_576).toFixed(1)} MB` : `${(b/1024).toFixed(0)} KB`;
const pct = (a: number, b: number) => b > 0 ? Math.round((a/b)*100) : 0;

// ─── HELPERS: NLP & FRAUD ───────────────────────────────────────────────────
function hashStrToCoord(str: string): [number, number] { let h1 = 0, h2 = 0; for (let i = 0; i < str.length; i++) { h1 = ((h1 << 5) - h1 + str.charCodeAt(i)) | 0; if (i > 0) h2 = ((h2 << 5) - h2 + str.charCodeAt(i)) | 0; } return [24.0 + (Math.abs(h1) % 100) / 15.0, 28.0 + (Math.abs(h2) % 100) / 10.0]; }
function calculateDistance(lat1: number, lon1: number, lat2: number, lon2: number): number { const R = 6371; const dLat = (lat2 - lat1) * Math.PI / 180; const dLon = (lon2 - lon1) * Math.PI / 180; const a = Math.sin(dLat/2) * Math.sin(dLat/2) + Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * Math.sin(dLon/2) * Math.sin(dLon/2); return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)); }
function levenshtein(a: string, b: string): number { const tmp=[]; let i, j, prev; for (i = 0; i <= a.length; i++) tmp[i] = [i]; for (j = 0; j <= b.length; j++) tmp[0][j] = j; for (i = 1; i <= a.length; i++) { for (j = 1; j <= b.length; j++) { prev = tmp[i - 1][j - 1]; if (a[i - 1] !== b[j - 1]) prev = Math.min(prev, tmp[i][j - 1], tmp[i - 1][j]) + 1; tmp[i][j] = prev; } } return tmp[a.length][b.length]; }
const posWords = new Set(["good","great","excellent","positive","satisfied","happy","clean","organized","active","sold","helped"]);
const negWords = new Set(["bad","poor","terrible","negative","angry","dirty","messy","out of stock","shortage","rude","issue","problem","missing"]);
function analyzeSentiment(text: string): number { const words = text.toLowerCase().split(/\s+/); if (words.length === 0) return 0; let score = 0; words.forEach(w => { if(posWords.has(w)) score++; if(negWords.has(w)) score--; }); return Math.max(-1, Math.min(1, score / words.length)); }
function extractTopics(text: string): string[] { const t = text.toLowerCase(); const topics: string[] = []; if (t.includes("out of stock") || t.includes("shortage") || t.includes("oos")) topics.push("Shortage"); if (t.includes("display") || t.includes("planogram") || t.includes("shelf")) topics.push("Display"); if (t.includes("competitor") || t.includes("promo") || t.includes("offer")) topics.push("Competitor"); if (t.includes("training") || t.includes("staff") || t.includes("employee")) topics.push("Staff"); if (t.includes("sales") || t.includes("sellout") || t.includes("revenue")) topics.push("Sales"); if (topics.length === 0 && text.length > 10) topics.push("General Ops"); return topics; }

// ─── PARSING LOGIC ──────────────────────────────────────────────────────────
function detectKind(sheetNames: string[]): VisitFileKind { if (sheetNames.includes("Mobile")) return "visit_raw"; if (sheetNames.includes("Submission Matrix")) return "visit_followup"; if (sheetNames.includes("Weekly Summary")) return "visit_summary"; return "unknown"; }
function parseSummary(rows: unknown[][]): EmployeeRecord[] { const headerRow = (rows[1]||[]) as unknown[]; const dayCols: number[] = []; for (let c=3;c<headerRow.length;c++){ const h=String(headerRow[c]??""); if(h&&!h.includes("Total")&&!h.toLowerCase().includes("notes")) dayCols.push(c); } const out: EmployeeRecord[] = []; for(let r=2;r<rows.length;r++){ const row=rows[r] as unknown[]; const name=row?.[0],code=row?.[1],dept=row?.[2]; if(!name||!code) continue; let reported=0; for(const c of dayCols){ const v=row[c]; if(v!==null&&v!==undefined&&v!==""&&Number(v)>0) reported++; } out.push({code:String(code),name:String(name),department:String(dept??"Unassigned"),daysReported:reported,daysPossible:dayCols.length}); } return out; }
function parseFollowUp(rows: unknown[][]): Map<string,{total:number;missing:string[]}> { const map = new Map<string,{total:number;missing:string[]}>(); for(let r=1;r<rows.length;r++){ const row=rows[r] as unknown[]; const code=row?.[1]; if(!code) continue; const total=Number(row?.[2])||0; const missingStr=String(row?.[3]??""); const missing=missingStr==="–"||!missingStr?[]:missingStr.split(",").map(s=>s.trim()).filter(Boolean); map.set(String(code),{total,missing}); } return map; }
function parseRawVisits(rows: unknown[][]): VisitRow[] { const out: VisitRow[] = []; for(let r=3;r<rows.length;r++){ const row=rows[r] as unknown[]; if(!row?.[1]) continue; const areaStr = String(row[5]??"Unknown"); const shopStr = String(row[4]??row[3]??""); let lat = 0, lon = 0; const coordMatch = areaStr.match(/(-?\d+\.\d+),\s*(-?\d+\.\d+)/); if (coordMatch) { lat = parseFloat(coordMatch[1]); lon = parseFloat(coordMatch[2]); } else { [lat, lon] = hashStrToCoord(areaStr + shopStr); } out.push({ date:String(row[0]??""),spvrCode:String(row[1]??""),spvrName:String(row[2]??""),shopCode:String(row[3]??""),shopName:shopStr,area:areaStr,governorate:String(row[6]???""), samsungShortage:String(row[7]??""),compShortage:String(row[8]??""),selloutMovement:String(row[9]??""),brand:String(row[10]???""), movement:String(row[11]??""),comment:String(row[12]??""),action1:String(row[13]??""),accountFeedback:String(row[14]??""),action2:String(row[15]???""), lat, lon }); } return out; }
function getHour(dateStr: string): number|null { const m = dateStr.match(/[T ](\d{2}):(\d{2})/); return m ? parseInt(m[1],10) : null; }

// ─── ENHANCED INTEGRITY BUILDER ─────────────────────────────────────────────
function buildIntegrityProfiles(visits: VisitRow[], employees: EmployeeRecord[]): IntegrityProfile[] {
  const byEmp = new Map<string,VisitRow[]>(); for(const v of visits){ if(!byEmp.has(v.spvrCode)) byEmp.set(v.spvrCode,[]); byEmp.get(v.spvrCode)!.push(v); }
  const nameByCode = new Map(employees.map(e=>[e.code,e.name])); const profiles: IntegrityProfile[] = [];
  for(const [code,rows] of byEmp.entries()){
    const sortedRows = [...rows].sort((a,b) => a.date.localeCompare(b.date)); const shops = new Set(rows.map(r=>r.shopCode));
    let blank=0,lowInfo=0,unsupported=0,lateNight=0,totalLen=0; const commentCounts = new Map<string,number>(); const allComments: string[] = []; const allTopicsSet = new Set<string>();
    let totalSentiment = 0, commentsWithSentiment = 0; const dailyMap = new Map<string, number>(); let teleportationFlags: string[] = []; let rushHourFlags: string[] = [];
    for(const r of rows) { const d = r.date.slice(0,10); dailyMap.set(d, (dailyMap.get(d)||0) + 1); } const dailyTrend = Array.from(dailyMap.entries()).map(([date, count]) => ({date:date.slice(5), count})).sort((a,b)=>a.date.localeCompare(b.date));
    for(let i=0; i<sortedRows.length; i++){
      const curr = sortedRows[i]; const currTime = new Date(curr.date).getTime();
      if(i < sortedRows.length - 1){ const next = sortedRows[i+1]; const nextTime = new Date(next.date).getTime(); const timeDiffHrs = Math.abs(nextTime - currTime) / (1000 * 3600); if(timeDiffHrs > 0 && timeDiffHrs < 12) { const distKm = calculateDistance(curr.lat, curr.lon, next.lat, next.lon); if(distKm / timeDiffHrs > 150){ teleportationFlags.push(`Traveled ${Math.round(distKm)}km in ${Math.round(timeDiffHrs*60)}m`); } } }
      const hour = getHour(curr.date); if(hour !== null && hour >= 8 && hour <= 18) { const visitsThisHour = sortedRows.filter(r => getHour(r.date) === hour && r.date.slice(0,10) === curr.date.slice(0,10)).length; if(visitsThisHour > 5 && !rushHourFlags.some(f => f.includes(curr.date.slice(0,10)))) rushHourFlags.push(`${visitsThisHour} visits in 1h on ${curr.date.slice(0,10)}`); }
      const c=curr.comment.toLowerCase(); totalLen+=curr.comment.length; if(!c) blank++; else { if(LOW_INFO_SET.has(c)||c.length<=3) lowInfo++; else { commentCounts.set(c,(commentCounts.get(c)||0)+1); allComments.push(c); const sentiment = analyzeSentiment(curr.comment); if(sentiment !== 0) { totalSentiment += sentiment; commentsWithSentiment++; } extractTopics(curr.comment).forEach(t => allTopicsSet.add(t)); } }
      const shortage=curr.samsungShortage&&!["n","no","none",""].includes(curr.samsungShortage.toLowerCase()); if((shortage||curr.selloutMovement.length>0)&&!c) unsupported++; if(hour!==null&&(hour>=23||hour<=4)) lateNight++;
    }
    let mutatedCopyPastes = 0; for(let i=0; i<allComments.length; i++){ for(let j=i+1; j<allComments.length; j++){ if(allComments[i] === allComments[j]) continue; const maxLen = Math.max(allComments[i].length, allComments[j].length); if(maxLen === 0) continue; if(1 - (levenshtein(allComments[i], allComments[j]) / maxLen) > 0.85) mutatedCopyPastes++; } }
    const commented=rows.length-blank; const uniqueComments=commentCounts.size+lowInfo; let templatedRepeats=0; let top: {text:string;count:number}|null=null;
    for(const [text,count] of commentCounts.entries()){ if(count>1) templatedRepeats+=count-1; if(!top||count>top.count) top={text,count}; }
    const uniquenessRatio=commented>0?pct(uniqueComments,commented):100; const avgCommentLen=rows.length>0?Math.round(totalLen/rows.length):0; const singleShopLoop=shops.size===1&&rows.length>4; const avgSentiment = commentsWithSentiment > 0 ? totalSentiment / commentsWithSentiment : 0;
    const suspiciousFlags: string[] = []; const blankPct=pct(blank,rows.length); if(blankPct>30) suspiciousFlags.push(`${blank} blank comments`); if(templatedRepeats>=3) suspiciousFlags.push(`Copy-pasted comments`); if(mutatedCopyPastes > 2) suspiciousFlags.push(`Mutated copy-pastes`); if(pct(lateNight,rows.length)>50&&rows.length>=4) suspiciousFlags.push(`Late night batch filing`); if(unsupported>0) suspiciousFlags.push(`Unsupported claims`); if(avgCommentLen<25&&rows.length>=5) suspiciousFlags.push(`Minimal effort comments`); if(singleShopLoop) suspiciousFlags.push(`Single shop loop`); if(teleportationFlags.length > 0) suspiciousFlags.push(`Impossible travel`); if(rushHourFlags.length > 0) suspiciousFlags.push(`Rush-hour filing`);
    const ghostScore=Math.min(100,Math.round(blankPct*0.30+Math.min(100,templatedRepeats*8)*0.20+pct(lateNight,rows.length)*0.25+Math.min(100,unsupported*15)*0.15+(singleShopLoop?20:0)*0.10));
    const penalty=templatedRepeats*4+lowInfo*1.5+unsupported*3+blank*1+mutatedCopyPastes*5+teleportationFlags.length*15+rushHourFlags.length*8; const integrityScore=Math.max(0,Math.min(100,Math.round(100-penalty)));
    profiles.push({ code,name:nameByCode.get(code)||code, totalVisits:rows.length,uniqueShops:shops.size, commentedVisits:commented,blankComments:blank,lowInfoComments:lowInfo, templatedRepeats,topRepeatedComment:top&&top.count>1?top:null, uniquenessRatio,unsupportedClaims:unsupported,integrityScore, avgCommentLen,lateNightCount:lateNight,singleShopLoop, ghostScore,suspiciousFlags, teleportationFlags, rushHourFlags, commentTopics: Array.from(allTopicsSet), avgSentiment, mutatedCopyPastes, dailyTrend });
  }
  return profiles.sort((a,b)=>a.integrityScore-b.integrityScore);
}

// ─── UI HELPERS ──────────────────────────────────────────────────────────────
const compColor = (p:number) => p>=80?"text-emerald-400":p>=50?"text-amber-400":"text-red-400";
const intgColor = (p:number) => p>=80?"text-cyan-400":p>=50?"text-amber-400":"text-red-400";
const ghostBadge = (s:number) => s>=60?{label:"HIGH RISK",bg:"bg-red-500/20",text:"text-red-400",border:"border-red-500/30"}:s>=30?{label:"SUSPICIOUS",bg:"bg-amber-500/20",text:"text-amber-400",border:"border-amber-500/30"}:{label:"CLEAN",bg:"bg-emerald-500/20",text:"text-emerald-400",border:"border-emerald-500/30"};
function Defs() { return (<defs><linearGradient id="areaCyan" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#22d3ee" stopOpacity={0.55}/><stop offset="100%" stopColor="#22d3ee" stopOpacity={0}/></linearGradient><linearGradient id="hbarCyan" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stopColor="#0e7490" stopOpacity={0.7}/><stop offset="100%" stopColor="#22d3ee" stopOpacity={1}/></linearGradient><linearGradient id="hbarPurple" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stopColor="#5b21b6" stopOpacity={0.7}/><stop offset="100%" stopColor="#a78bfa" stopOpacity={1}/></linearGradient><linearGradient id="hbarAmber" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stopColor="#92400e" stopOpacity={0.7}/><stop offset="100%" stopColor="#fbbf24" stopOpacity={1}/></linearGradient><filter id="shadow" height="140%" width="140%"><feDropShadow dx="0" dy="3" stdDeviation="4" floodColor="#000" floodOpacity={0.55}/></filter><filter id="glowCyan"><feDropShadow dx="0" dy="0" stdDeviation="4" floodColor="#22d3ee" floodOpacity={0.55}/></filter></defs>); }
const GlowDot = (props: any) => { const {cx,cy} = props; if(!cx||!cy) return null; return (<g><circle cx={cx} cy={cy} r={7} fill="#22d3ee" opacity={0.15}/><circle cx={cx} cy={cy} r={3.5} fill="#22d3ee" filter="url(#glowCyan)"/></g>); };
const GradBar = (props: any) => { const {x,y,width,height,index} = props; const grads = ["url(#hbarCyan)","url(#hbarPurple)","url(#hbarAmber)"]; const fill = grads[index%grads.length]; return (<g filter="url(#shadow)"><rect x={x} y={y+1} width={width} height={height-2} rx={5} ry={5} fill={fill}/><rect x={x} y={y+1} width={width*0.38} height={(height-2)*0.4} rx={5} ry={5} fill="white" opacity={0.09}/></g>); };
function GhostMeter({score}:{score:number}) { const filled=Math.round(score/10); const color=score>=60?"bg-red-400":score>=30?"bg-amber-400":"bg-emerald-400"; return (<div className="flex items-center gap-0.5">{Array.from({length:10}).map((_,i)=>(<div key={i} className={`w-1.5 h-3.5 rounded-sm ${i<filled?color:"bg-neutral-700"}`}/>))}</div>); }
function FlagTooltip({profile}:{profile:IntegrityProfile}) { const [open,setOpen]=useState(false); const n=profile.suspiciousFlags.length; if(n===0) return <span className="text-neutral-600 text-xs">—</span>; return (<div className="relative inline-block"><button onMouseEnter={()=>setOpen(true)} onMouseLeave={()=>setOpen(false)} className="flex items-center gap-1 text-xs font-medium text-amber-400"><Zap className="w-3 h-3"/>{n} flag{n>1?"s":""}</button>{open&&(<div className="absolute z-50 bottom-full left-0 mb-2 w-80 bg-neutral-800 border border-amber-500/40 rounded-xl p-3 shadow-2xl shadow-black/60"><p className="text-[10px] text-amber-300 font-bold mb-2">⚠ SUSPICIOUS SIGNALS</p><div className="space-y-1.5">{profile.suspiciousFlags.map((f,i)=>(<div key={i} className="flex items-start gap-2 text-xs text-neutral-200"><span className="w-1.5 h-1.5 rounded-full bg-amber-400 mt-1 shrink-0"/>{f}</div>))}</div></div>)}</div>); }
function WhyCell({e,ip}:{e:any;ip:IntegrityProfile|undefined}) { const reasons: string[] = []; if(e.compliance<50) reasons.push(`Only ${e.compliance}% submission rate`); if(ip){ if(ip.blankComments>0) reasons.push(`${ip.blankComments} blank comments`); if(ip.templatedRepeats>=3) reasons.push("Copy-pastes repeatedly"); if(ip.teleportationFlags.length>0) reasons.push("Impossible travel"); } if(reasons.length===0) return <span className="text-neutral-600 text-xs">—</span>; return (<ul className="space-y-0.5">{reasons.map((r,i)=>(<li key={i} className="text-[10px] text-neutral-300 flex items-start gap-1"><span className="text-red-400 shrink-0">•</span>{r}</li>))}</ul>); }
function SparklineChart({ data, color }: { data: {count:number}[], color: string }) { if (!data || data.length < 2) return <span className="text-neutral-700 text-[10px]">—</span>; return (<ResponsiveContainer width={50} height={20}><LineChart data={data} margin={{ top: 2, right: 2, bottom: 2, left: 2 }}><Line type="monotone" dataKey="count" stroke={color} strokeWidth={1.5} dot={false} /></LineChart></ResponsiveContainer>); }

// ─── 3D SYSTEM SIDE PANEL ───────────────────────────────────────────────────
function SystemSidePanel({ visits, employees, files }: { visits: VisitRow[], employees: EmployeeRecord[], files: LoadedFile[] }) {
  const [time, setTime] = useState(new Date());
  useEffect(() => { const timer = setInterval(() => setTime(new Date()), 1000); return () => clearInterval(timer); }, []);
  
  return (
    <div className="fixed left-0 top-0 h-screen w-72 bg-neutral-950 border-r border-neutral-800 z-50 flex flex-col shadow-2xl overflow-y-auto"
      style={{ background: "linear-gradient(180deg, #0a0a12 0%, #111119 100%)" }}>
      <div className="p-4 border-b border-neutral-800 bg-neutral-900/50">
        <div className="flex items-center gap-2 mb-2"><div className="w-1 h-6 rounded-full bg-cyan-500"/><h2 className="text-sm font-bold text-white tracking-widest">SMARTSENSE</h2></div>
        <p className="text-[10px] text-neutral-500 ml-3">MX TEAM GRP - Command Center</p>
      </div>

      <div className="p-4 border-b border-neutral-800">
        <div className="bg-neutral-900 rounded-lg p-3 border border-cyan-900/30 shadow-[inset_0_0_15px_rgba(0,0,0,0.8)]">
          <div className="flex items-center justify-between mb-1"><Monitor className="w-3 h-3 text-cyan-400"/><span className="text-[10px] text-cyan-400 font-mono">SYSTEM TIME</span></div>
          <p className="text-2xl font-bold text-white font-mono tracking-wider">{time.toLocaleTimeString('en-US', {hour12: false})}</p>
          <p className="text-[10px] text-neutral-400 font-mono">{time.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
        </div>
      </div>

      <div className="p-4 border-b border-neutral-800 space-y-3">
        <h3 className="text-[10px] text-neutral-500 tracking-widest font-bold">DATA STATS</h3>
        <div className="flex justify-between text-xs"><span className="text-neutral-400">Files Imported</span><span className="text-white font-bold">{files.length}</span></div>
        <div className="flex justify-between text-xs"><span className="text-neutral-400">Total Staff</span><span className="text-white font-bold">{employees.length}</span></div>
        <div className="flex justify-between text-xs"><span className="text-neutral-400">Total Visits</span><span className="text-white font-bold">{visits.length}</span></div>
        {visits.length > 0 && <div className="flex justify-between text-xs"><span className="text-neutral-400">Date Range</span><span className="text-cyan-400 font-mono text-[10px]">{visits[0]?.date.slice(0,10)} - {visits[visits.length-1]?.date.slice(0,10)}</span></div>}
      </div>

      <div className="p-4 space-y-3">
        <h3 className="text-[10px] text-neutral-500 tracking-widest font-bold">PAGE GUIDE</h3>
        {[
          {icon: Gauge, label: "KPI Dashboard", desc: "High level metrics"},
          {icon: Activity, label: "Volume Trends", desc: "Daily visit charts"},
          {icon: UserCheck, label: "Team Tracker", desc: "Individual performance"},
          {icon: MapPin, label: "Egypt Map", desc: "Spatial intelligence"},
          {icon: Ghost, label: "Ghost Desk", desc: "Fraud detection"},
          {icon: FileDown, label: "CEO Report", desc: "Export professional PDF"},
        ].map(g => (
          <div key={g.label} className="flex items-start gap-3 p-2 rounded-lg hover:bg-neutral-800/50 transition-colors cursor-pointer group">
            <g.icon className="w-4 h-4 text-neutral-600 group-hover:text-cyan-400 transition-colors mt-0.5"/>
            <div><p className="text-xs font-medium text-neutral-300 group-hover:text-white">{g.label}</p><p className="text-[10px] text-neutral-600">{g.desc}</p></div>
          </div>
        ))}
      </div>
      
      <div className="mt-auto p-4 border-t border-neutral-800">
        <div className="flex items-center gap-2 text-[10px] text-neutral-600"><div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"/>System Online</div>
      </div>
    </div>
  );
}

// ─── EMPLOYEE SCORECARD MODAL (ORGANIZED EVIDENCE) ──────────────────────────
function EmployeeScorecardModal({ employee, profile, visits, onClose }: { employee: any; profile: IntegrityProfile | undefined; visits: VisitRow[]; onClose: () => void }) {
  const integrity = profile?.integrityScore ?? 100; const trust = Math.round(employee.compliance * 0.5 + integrity * 0.5);
  const badge = ghostBadge(profile?.ghostScore ?? 0); const complianceColor = employee.compliance >= 80 ? "#34d399" : employee.compliance >= 50 ? "#fbbf24" : "#f87171";
  const integrityColor = integrity >= 80 ? "#22d3ee" : integrity >= 50 ? "#fbbf24" : "#f87171"; const trustColor = trust >= 80 ? "#a78bfa" : trust >= 50 ? "#fbbf24" : "#f87171";
  const pipRef = useRef<HTMLDivElement>(null);
  
  // Organized Evidence
  const goodEvidence = visits.filter(v => analyzeSentiment(v.comment) > 0.2 && v.comment.length > 20);
  const badEvidence = visits.filter(v => v.comment.trim() === "" || LOW_INFO_SET.has(v.comment.toLowerCase().trim()) || analyzeSentiment(v.comment) < -0.2);

  const handleGeneratePIP = async () => { if (!pipRef.current) return; const canvas = await html2canvas(pipRef.current, { backgroundColor: '#111' }); const url = canvas.toDataURL('image/png'); const a = document.createElement('a'); a.href = url; a.download = `PIP-${employee.name}.png`; a.click(); };
  useEffect(() => { const handler = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); }; document.addEventListener("keydown", handler); return () => document.removeEventListener("keydown", handler); }, [onClose]);

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4" onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="absolute inset-0 bg-black/70 backdrop-blur-sm" />
      <div className="relative z-10 bg-neutral-900 border border-neutral-700 rounded-2xl shadow-2xl shadow-black/80 w-full max-w-4xl max-h-[90vh] overflow-y-auto">
        <div className="sticky top-0 z-20 bg-neutral-900/95 backdrop-blur-md border-b border-neutral-700 px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="relative w-16 h-16 rounded-xl overflow-hidden border-2 border-cyan-500/30 bg-neutral-800 flex items-center justify-center">
              <svg width="100%" height="100%" viewBox="0 0 100 100"><rect width="100" height="100" fill="#1a1a1a"/><path d="M30 100 L30 70 Q50 50 70 70 L70 100 Z" fill="#111111"/><circle cx="50" cy="35" r="20" fill="#333"/><text x="50" y="88" textAnchor="middle" fill="white" fontSize="12" fontWeight="bold" fontFamily="Arial">SmartSense</text></svg>
            </div>
            <div><h3 className="text-white font-bold text-lg">{employee.name}</h3><p className="text-cyan-400 font-mono text-xs">{employee.code}</p><p className="text-[10px] font-bold text-neutral-400 tracking-widest mt-1">SMARTSENSE-LTD - MX TEAM GRP</p></div>
          </div>
          <div className="flex items-center gap-2"><button onClick={handleGeneratePIP} className="flex items-center gap-1.5 text-xs font-medium text-red-300 bg-red-500/10 border border-red-500/30 px-3 py-1.5 rounded-lg hover:bg-red-500/20"><Printer className="w-3.5 h-3.5"/> Generate PIP</button><button onClick={onClose} className="w-8 h-8 rounded-lg bg-neutral-800 border border-neutral-600 flex items-center justify-center text-neutral-400 hover:text-white"><X className="w-4 h-4" /></button></div>
        </div>
        <div className="p-6 space-y-6">
          <div ref={pipRef} className="absolute left-[-9999px] top-[-9999px] w-[800px] p-8 bg-neutral-900 text-white space-y-4"><h1 className="text-2xl font-bold border-b border-neutral-700 pb-2">Performance Improvement Plan - SMARTSENSE</h1><div className="grid grid-cols-2 gap-4 text-sm"><p><strong>Employee:</strong> {employee.name}</p><p><strong>Department:</strong> {employee.department}</p><p><strong>Compliance:</strong> {employee.compliance}%</p><p><strong>Integrity:</strong> {integrity}%</p></div><div><h3 className="font-bold text-red-400 mb-1">Key Infractions:</h3><ul className="list-disc pl-5 text-neutral-300">{profile?.suspiciousFlags.slice(0,3).map(f => <li key={f}>{f}</li>)}</ul></div></div>

          <div className="flex flex-wrap items-center justify-center gap-8">
            {[ { val: employee.compliance, label: "Compliance", color: complianceColor }, { val: integrity, label: "Integrity", color: integrityColor }, { val: trust, label: "Trust", color: trustColor } ].map(g => (
              <div key={g.label} className="flex flex-col items-center gap-2">
                <div style={{ width: 120, height: 120, borderRadius: "50%", background: `conic-gradient(from 135deg, ${g.color} ${(g.val/100)*360}deg, rgba(30,30,42,0.9) ${(g.val/100)*360}deg)`, boxShadow: `0 10px 30px rgba(0,0,0,0.7), 0 0 15px ${g.color}22`, padding: 10 }}>
                  <div style={{ width: "100%", height: "100%", borderRadius: "50%", background: "radial-gradient(circle at 42% 38%, #1e1e30 0%, #0c0c14 80%)", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column" }}>
                    <span style={{ fontSize: 20, fontWeight: "bold", fontFamily: "monospace", color: g.color }}>{g.val}%</span>
                    <span style={{ fontSize: 8, color: "#737373", textTransform: "uppercase", letterSpacing: 2, marginTop: 2 }}>{g.label}</span>
                  </div>
                </div>
              </div>
            ))}
          </div>

          {/* ORGANIZED EVIDENCE LEDGER */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 max-h-[40vh] overflow-y-auto pr-1">
            <div className="bg-emerald-950/20 border border-emerald-900/40 rounded-xl p-4">
              <h4 className="text-xs font-bold text-emerald-400 mb-3 flex items-center gap-2"><CheckCircle className="w-4 h-4"/> Positive Evidence</h4>
              {goodEvidence.length === 0 ? <p className="text-xs text-neutral-600">No strong positive evidence found.</p> : (
                <div className="space-y-2">{goodEvidence.slice(0,20).map((v,i)=>(
                  <div key={i} className="bg-neutral-900/50 rounded p-2 text-[11px] border border-emerald-900/20">
                    <div className="flex justify-between text-neutral-500 mb-1"><span>{v.date.slice(0,10)}</span><span className="text-emerald-400">+Sentiment</span></div>
                    <p className="text-neutral-300 italic">"{v.comment}"</p>
                  </div>
                ))}</div>
              )}
            </div>
            <div className="bg-red-950/20 border border-red-900/40 rounded-xl p-4">
              <h4 className="text-xs font-bold text-red-400 mb-3 flex items-center gap-2"><AlertTriangle className="w-4 h-4"/> Negative / Blank Evidence</h4>
              {badEvidence.length === 0 ? <p className="text-xs text-neutral-600">No negative evidence found.</p> : (
                <div className="space-y-2">{badEvidence.slice(0,20).map((v,i)=>(
                  <div key={i} className="bg-neutral-900/50 rounded p-2 text-[11px] border border-red-900/20">
                    <div className="flex justify-between text-neutral-500 mb-1"><span>{v.date.slice(0,10)}</span><span className="text-red-400">{v.comment.trim() === "" ? "BLANK" : "LOW INFO"}</span></div>
                    <p className="text-neutral-300 italic">{v.comment || "(No Comment Provided)"}</p>
                  </div>
                ))}</div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── VISIT MAP WIDGET (EGYPT FOCUS & LED PULSE) ─────────────────────────────
function VisitMapWidget({ visits, integrityByCode }: { visits: VisitRow[]; integrityByCode: Map<string, IntegrityProfile> }) {
  const [timeSlider, setTimeSlider] = useState(24);
  const [isClient, setIsClient] = useState(false);
  useEffect(() => { setIsClient(true); }, []);
  const filteredByTime = useMemo(() => { if (timeSlider === 24) return visits; return visits.filter(v => { const hr = getHour(v.date); return hr !== null && hr <= timeSlider; }); }, [visits, timeSlider]);

  if (!isClient) return <div className="h-[400px] bg-neutral-900 animate-pulse rounded-xl" />;
  return (
    <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
      <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-700"><div className="flex items-center gap-2"><MapPin className="w-4 h-4 text-cyan-400" /><span className="text-xs text-neutral-300 tracking-wider font-semibold">EGYPT SPATIAL INTELLIGENCE</span></div><div className="flex items-center gap-4"><span className="text-[10px] text-neutral-500">Route: {timeSlider === 24 ? 'Full Day' : `Until ${timeSlider}:00`}</span><input type="range" min="8" max="24" value={timeSlider} onChange={e => setTimeSlider(parseInt(e.target.value))} className="w-32 h-1 bg-neutral-700 rounded-lg appearance-none cursor-pointer accent-cyan-500" /></div></div>
      <style>{pulseStyle}</style>
      <div className="relative h-[500px] w-full bg-neutral-950 z-0">
        <MapContainer center={[26.8206, 30.8025]} zoom={6} maxBounds={[[22.0, 24.0], [31.5, 36.0]]} maxZoom={18} minZoom={6} style={{ height: '100%', width: '100%' }} className="z-0">
          <TileLayer url="https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png" />
          {filteredByTime.map((v, idx) => {
            const ip = integrityByCode.get(v.spvrCode); const isGhost = ip && ip.ghostScore >= 60;
            return (
              <CircleMarker key={idx} center={[v.lat, v.lon]} radius={isGhost ? 8 : 4} pathOptions={{ color: isGhost ? '#f87171' : '#22d3ee', weight: isGhost ? 2 : 1, fillColor: isGhost ? 'rgba(248,113,113,0.6)' : 'rgba(34,211,238,0.3)', fillOpacity: 0.9, className: isGhost ? 'pulse-led' : '' }}>
                <LeafletTooltip direction="top" offset={[0, -5]}><div className="text-xs bg-neutral-900 p-1 rounded shadow-lg border border-neutral-700 font-mono"><span className="text-cyan-400 font-bold">[{v.spvrCode}]</span> <span className="text-white">{v.shopName}</span></div></LeafletTooltip>
              </CircleMarker>
            );
          })}
        </MapContainer>
      </div>
    </div>
  );
}

// ─── EMPLOYEE PERFORMANCE TRACKER PANEL ──────────────────────────────────────
function EmployeePerformanceTracker({ enriched, integrityByCode, fieldTrust, onSelectEmployee }: any) {
  const [selectedCode, setSelectedCode] = useState<string>(enriched[0]?.code || "");
  const emp = enriched.find((e: any) => e.code === selectedCode); const ip = integrityByCode.get(selectedCode); const ft = fieldTrust.find((e: any) => e.code === selectedCode);
  const teamMembers = enriched.filter((e: any) => e.department === emp?.department); const avgTeamCompliance = teamMembers.length > 0 ? Math.round(teamMembers.reduce((s: number, m: any) => s + m.compliance, 0) / teamMembers.length) : 0;
  const qualityGrade = ft ? (ft.trust >= 80 ? 'A' : ft.trust >= 50 ? 'C' : 'F') : 'N/A';
  const comparisonData = [ { name: 'You', compliance: emp?.compliance || 0, integrity: ip?.integrityScore || 0 }, { name: 'Team Avg', compliance: avgTeamCompliance, integrity: teamMembers.length > 0 ? Math.round(teamMembers.reduce((s: number, m: any) => s + (integrityByCode.get(m.code)?.integrityScore ?? 100), 0) / teamMembers.length) : 0 } ];

  return (
    <div className="bg-neutral-900 border border-cyan-500/20 rounded-xl overflow-hidden">
      <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-700"><div className="flex items-center gap-2"><UserCheck className="w-4 h-4 text-cyan-400" /><span className="text-xs text-neutral-300 tracking-wider font-semibold">EMPLOYEE TRACKER</span></div><select value={selectedCode} onChange={(e) => { setSelectedCode(e.target.value); onSelectEmployee(e.target.value); }} className="bg-neutral-800 border border-neutral-600 text-xs text-neutral-300 rounded px-3 py-1.5">{enriched.map((e: any) => <option key={e.code} value={e.code}>{e.name}</option>)}</select></div>
      {emp && ip && ft && (<div className="p-5 grid grid-cols-1 lg:grid-cols-3 gap-6"><div className="space-y-3"><h3 className="text-sm font-bold text-white">{emp.name}</h3><table className="w-full text-xs border-collapse"><tbody><tr className="border-b border-neutral-800"><td className="py-2 text-neutral-400">Team Avg %</td><td className="py-2 text-right font-mono font-bold text-white">{avgTeamCompliance}%</td></tr><tr className="border-b border-neutral-800"><td className="py-2 text-neutral-400">Total Visits</td><td className="py-2 text-right font-mono font-bold text-white">{ip.totalVisits}</td></tr><tr className="border-b border-neutral-800"><td className="py-2 text-neutral-400">Quality Grade</td><td className={`py-2 text-right font-mono font-bold ${qualityGrade === 'A' ? 'text-emerald-400' : qualityGrade === 'C' ? 'text-amber-400' : 'text-red-400'}`}>{qualityGrade}</td></tr><tr><td className="py-2 text-neutral-400">Duplicates</td><td className="py-2 text-right font-mono font-bold text-amber-400">{ip.templatedRepeats + ip.mutatedCopyPastes}</td></tr></tbody></table></div><div><p className="text-[10px] text-neutral-500 tracking-wider font-bold mb-2">VISIT TREND</p><div className="h-48"><ResponsiveContainer width="100%" height="100%"><AreaChart data={ip.dailyTrend}><defs><Defs/></defs><XAxis dataKey="date" tick={{fill:"#525252",fontSize:9}} axisLine={false} tickLine={false} /><YAxis tick={{fill:"#525252",fontSize:9}} axisLine={false} tickLine={false} width={26} /><Tooltip contentStyle={{background:"#111",border:"1px solid #2a2a2a",borderRadius:8,color:"#fff"}} /><Area type="monotone" dataKey="count" stroke="#22d3ee" fill="url(#areaCyan)" /></AreaChart></ResponsiveContainer></div></div><div><p className="text-[10px] text-neutral-500 tracking-wider font-bold mb-2">VS TEAM</p><div className="h-48"><ResponsiveContainer width="100%" height="100%"><BarChart data={comparisonData}><XAxis dataKey="name" tick={{fill:"#e5e5e5",fontSize:10}} axisLine={false} tickLine={false} /><YAxis domain={[0, 100]} tick={{fill:"#525252",fontSize:9}} axisLine={false} tickLine={false} width={26}/><Tooltip contentStyle={{background:"#111",border:"1px solid #2a2a2a",borderRadius:8,color:"#fff"}} /><Bar dataKey="compliance" fill="#34d399" radius={[4,4,0,0]} /><Bar dataKey="integrity" fill="#22d3ee" radius={[4,4,0,0]} /></BarChart></ResponsiveContainer></div></div></div>)}
    </div>
  );
}

function CommandBar({ isOpen, onClose, employees, onSelectEmployee, onExport }:{ isOpen:boolean, onClose:()=>void, employees:any[], onSelectEmployee:(code:string)=>void, onExport:()=>void }) {
  const [query, setQuery] = useState(""); if (!isOpen) return null;
  const filtered = employees.filter(e => e.name.toLowerCase().includes(query.toLowerCase()) || e.code.toLowerCase().includes(query.toLowerCase()));
  return (<div className="fixed inset-0 z-[200] flex items-start justify-center pt-[20vh]" onClick={onClose}><div className="absolute inset-0 bg-black/60 backdrop-blur-sm" /><div className="relative z-10 w-full max-w-lg bg-neutral-900 border border-neutral-600 rounded-2xl shadow-2xl overflow-hidden" onClick={e => e.stopPropagation()}><div className="flex items-center gap-3 px-4 border-b border-neutral-700"><Search className="w-5 h-5 text-neutral-400" /><input autoFocus placeholder="Search or type 'export'..." value={query} onChange={e => setQuery(e.target.value)} className="w-full bg-transparent py-4 text-white text-sm outline-none" /><kbd className="text-[10px] text-neutral-500 border border-neutral-600 rounded px-1.5 py-0.5">ESC</kbd></div><div className="max-h-80 overflow-y-auto p-2">{query.toLowerCase() === 'export' && <button onClick={onExport} className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-purple-300 hover:bg-neutral-800"><Download className="w-4 h-4" /> CEO Report</button>}{filtered.map(e => (<button key={e.code} onClick={() => { onSelectEmployee(e.code); onClose(); }} className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-white hover:bg-neutral-800"><span>{e.name}</span><span className="text-neutral-500 font-mono text-xs ml-auto">{e.code}</span></button>))}</div></div></div>);
}

// ─── PROFESSIONAL CEO EXPORT REPORT ─────────────────────────────────────────
function exportCEOReport(enriched: any[], integrityProfiles: IntegrityProfile[], visits: VisitRow[], audit: any[]) {
  const timestamp = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const totalStaff = enriched.length; const avgComp = Math.round(enriched.reduce((s,e)=>s+e.compliance,0)/Math.max(1,totalStaff)); const highRisk = integrityProfiles.filter(p => p.ghostScore >= 60).length;
  const sortedProfiles = [...integrityProfiles].sort((a,b) => a.ghostScore - b.ghostScore);
  const bestPerformers = sortedProfiles.slice(0, 3).reverse(); const weakestPerformers = sortedProfiles.slice(0, 3);

  const reportHTML = `<!DOCTYPE html><html><head><title>SmartSense CEO Report</title><style>
    body{font-family:'Segoe UI',sans-serif;background:#f8f9fa;color:#333;margin:0;padding:40px}
    .header{display:flex;justify-content:space-between;align-items:center;border-bottom:4px solid #000;padding-bottom:20px;margin-bottom:30px}
    .header h1{margin:0;font-size:28px;letter-spacing:1px}.header p{margin:5px 0 0;color:#555;font-size:14px}
    .kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:20px;margin-bottom:40px}
    .kpi-box{background:#fff;padding:20px;border-radius:8px;box-shadow:0 10px 20px rgba(0,0,0,0.08);text-align:center;border-top:4px solid #22d3ee;transform:perspective(1000px) rotateX(2deg)}
    .kpi-box.risk{border-top-color:#f87171}.kpi-box h3{margin:0;color:#888;font-size:12px;text-transform:uppercase;letter-spacing:1px}.kpi-box .value{font-size:36px;font-weight:bold;margin:10px 0;color:#111}
    table{width:100%;border-collapse:collapse;background:#fff;box-shadow:0 10px 20px rgba(0,0,0,0.05);border-radius:8px;overflow:hidden;margin-bottom:30px}
    th{background:#1e293b;color:#fff;padding:12px 15px;text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:1px}
    td{padding:12px 15px;border-bottom:1px solid #eee;font-size:13px;color:#555} tr:hover td{background:#f1f5f9}
    .badge{padding:4px 8px;border-radius:12px;font-size:10px;font-weight:bold}.badge-high{background:#fee2e2;color:#b91c1c}.badge-clean{background:#d1fae5;color:#047857}
    .metrics{background:#fff;padding:20px;border-radius:8px;margin-bottom:30px;border-left:5px solid #a78bfa}
    .team-list{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:30px}
    .team-box{background:#fff;padding:20px;border-radius:8px;box-shadow:0 4px 6px rgba(0,0,0,0.05)}
    footer{text-align:center;color:#888;font-size:12px;margin-top:40px;border-top:1px solid #ddd;padding-top:20px}
  </style></head><body>
    <div class="header"><div><h1>SMARTSENSE-LTD</h1><p>MX TEAM GRP - CEO Integrity Dashboard</p></div><div style="text-align:right"><p style="font-weight:bold;font-size:16px">${timestamp}</p><p>Confidential Executive Report</p></div></div>
    
    <div class="kpi-grid"><div class="kpi-box"><h3>Total Staff</h3><div class="value">${totalStaff}</div></div><div class="kpi-box"><h3>Avg Compliance</h3><div class="value">${avgComp}%</div></div><div class="kpi-box risk"><h3>High Risk</h3><div class="value">${highRisk}</div></div><div class="kpi-box"><h3>Flags</h3><div class="value">${audit.length}</div></div></div>

    <div class="metrics"><h3 style="margin-top:0">📊 Metrics Definition</h3><ul style="margin:0;padding-left:20px;color:#555;font-size:14px"><li><b>Compliance:</b> Percentage of expected days the employee submitted a report.</li><li><b>Integrity:</b> Score based on comment quality, lack of copy-pasting, and lack of unsupported claims.</li><li><b>Ghost Score:</b> Risk of data fabrication (0-100). Based on late-night filing, impossible travel, and repeated blanks.</li><li><b>Trust:</b> Composite of Compliance (50%) + Integrity (50%).</li></ul></div>

    <div class="team-list">
      <div class="team-box"><h3 style="color:#047857;margin-top:0">🏆 Best Performers</h3>${bestPerformers.map(p => `<p style="margin:5px 0"><b>${p.name}</b> - Trust: ${Math.round((100-p.ghostScore)*0.5 + p.integrityScore*0.5)}% <span class="badge badge-clean">CLEAN</span></p>`).join('')}</div>
      <div class="team-box"><h3 style="color:#b91c1c;margin-top:0">⚠️ Weakest Performers</h3>${weakestPerformers.map(p => `<p style="margin:5px 0"><b>${p.name}</b> - Ghost: ${p.ghostScore} <span class="badge badge-high">RISK</span> <span style="color:#888;font-size:12px">${p.suspiciousFlags[0]||''}</span></p>`).join('')}</div>
    </div>

    <h2 style="font-size:18px;margin-bottom:15px">Ghost Activity Suspects</h2><table><thead><tr><th>Employee</th><th>Code</th><th>Ghost Score</th><th>Risk Level</th></tr></thead><tbody>${integrityProfiles.filter(p=>p.ghostScore>=30).slice(0,15).map(p => `<tr><td style="font-weight:bold">${p.name}</td><td>${p.code}</td><td style="font-weight:bold">${p.ghostScore}</td><td><span class="badge ${p.ghostScore>=60?'badge-high':'badge-clean'}">${p.ghostScore>=60?'HIGH RISK':'SUSPICIOUS'}</span></td></tr>`).join('')}</tbody></table>

    <footer>Generated by SmartSense Field Reporting System © 2026</footer></body></html>`;
  const win = window.open('', '_blank'); if (win) { win.document.write(reportHTML); win.document.close(); }
}

// ─── MAIN COMPONENT ──────────────────────────────────────────────────────────
export default function VisitCompliancePage() {
  const [files,setFiles]=useState<LoadedFile[]>([]);
  const [employees,setEmployees]=useState<EmployeeRecord[]>([]);
  const [followUpMap,setFollowUpMap]=useState<Map<string,{total:number;missing:string[]}>>(new Map());
  const [rawCodes,setRawCodes]=useState<Set<string>>(new Set());
  const [visits,setVisits]=useState<VisitRow[]>([]);
  const [isDragging,setIsDragging]=useState(false);
  const [isProcessing,setIsProcessing]=useState(false);
  const [deptFilter,setDeptFilter]=useState("All");
  const [lowestTab,setLowestTab]=useState<"trust"|"ghost">("trust");
  const [selectedEmployeeCode, setSelectedEmployeeCode] = useState<string | null>(null);
  const [statusFilter, setStatusFilter] = useState<string | null>(null);
  const [commandBarOpen, setCommandBarOpen] = useState(false);
  const [selectedForAction, setSelectedForAction] = useState<Set<string>>(new Set());

  useEffect(() => { const saved = localStorage.getItem('visitComplianceState'); if (saved) { try { const state = JSON.parse(saved); if(state.employees) setEmployees(state.employees); if(state.visits) setVisits(state.visits); } catch(e) {} } }, []);
  useEffect(() => { if (employees.length > 0 || visits.length > 0) { localStorage.setItem('visitComplianceState', JSON.stringify({ employees, visits, files })); } }, [employees, visits, files]);
  useEffect(() => { const handler = (e: KeyboardEvent) => { if ((e.metaKey || e.ctrlKey) && e.key === 'k') { e.preventDefault(); setCommandBarOpen(true); } }; window.addEventListener('keydown', handler); return () => window.removeEventListener('keydown', handler); }, []);

  const toggleSelection = (code: string) => { setSelectedForAction(prev => { const next = new Set(prev); if (next.has(code)) next.delete(code); else next.add(code); return next; }); };
  const handleBulkEmail = () => { const names = Array.from(selectedForAction).join(','); window.open(`mailto:?subject=Action Required: Compliance Flags&body=Dear ${names},%0D%0APlease review your field compliance metrics.`); };

  const handleFiles=useCallback(async(fileList:FileList)=>{
    setIsProcessing(true); const newLoaded:LoadedFile[]=[]; let newEmployees:EmployeeRecord[]|null=null; let newFollowUp:Map<string,{total:number;missing:string[]}>|null=null; let newRaw:Set<string>|null=null; let newVisits:VisitRow[]|null=null;
    for(const file of Array.from(fileList)){ try{ const buf=await file.arrayBuffer(); const wb=XLSX.read(buf,{type:"array"}); const kind=detectKind(wb.SheetNames); newLoaded.push({name:file.name,kind,size:fmtBytes(file.size)}); if(kind==="visit_summary"){ newEmployees=parseSummary(XLSX.utils.sheet_to_json(wb.Sheets["Weekly Summary"],{header:1,defval:null}) as unknown[][]); } else if(kind==="visit_followup"){ newFollowUp=parseFollowUp(XLSX.utils.sheet_to_json(wb.Sheets["Submission Matrix"],{header:1,defval:null}) as unknown[][]); } else if(kind==="visit_raw"){ newVisits=parseRawVisits(XLSX.utils.sheet_to_json(wb.Sheets["Mobile"],{header:1,defval:null}) as unknown[][]); newRaw=new Set(newVisits.map(v=>v.spvrCode)); } }catch{ newLoaded.push({name:file.name,kind:"unknown",size:"—"}); } }
    setFiles(prev=>[...prev,...newLoaded]); if(newFollowUp) setFollowUpMap(newFollowUp); if(newRaw) setRawCodes(newRaw); if(newEmployees) setEmployees(newEmployees); if(newVisits) setVisits(newVisits); setIsProcessing(false);
  },[]);

  const enriched = useMemo(()=>employees.map(e=>({ ...e, inFollowUp:followUpMap.has(e.code), inRaw:rawCodes.has(e.code), compliance:pct(e.daysReported,e.daysPossible) })),[employees,followUpMap,rawCodes]);
  const integrityProfiles=useMemo(()=>buildIntegrityProfiles(visits,employees),[visits,employees]);
  const integrityByCode=useMemo(()=>new Map(integrityProfiles.map(p=>[p.code,p])),[integrityProfiles]);
  const audit = useMemo(()=>{ const issues:{severity:"high"|"medium";message:string}[]=[]; for(const e of enriched){ const ip=integrityByCode.get(e.code); if(ip){ if(ip.teleportationFlags.length>0) issues.push({severity:"high",message:`${e.name}: Impossible travel detected.`}); } } return issues; },[enriched,integrityByCode]);
  const departments=["All",...Array.from(new Set(enriched.map(e=>e.department)))];
  const filtered = useMemo(() => { let list = deptFilter==="All"?enriched:enriched.filter(e=>e.department===deptFilter); if (statusFilter === "AT RISK") list = list.filter(e => Math.round(e.compliance*0.5+(integrityByCode.get(e.code)?.integrityScore??100)*0.5) < 50); return list; }, [deptFilter, enriched, integrityByCode, statusFilter]);
  const deptRollup=useMemo(()=> Array.from(new Set(enriched.map(e=>e.department))).map(dept=>{ const members=enriched.filter(e=>e.department===dept); return {dept,headcount:members.length,compliance:pct(members.reduce((s,m)=>s+m.daysReported,0),members.reduce((s,m)=>s+m.daysPossible,0)),chronic:members.filter(m=>m.compliance<50).length}; }).sort((a,b)=>b.compliance-a.compliance), [enriched]);
  const overallCompliance=pct(enriched.reduce((s,e)=>s+e.daysReported,0),enriched.reduce((s,e)=>s+e.daysPossible,0));
  const fieldTrust=useMemo(()=>enriched.map(e=>{ const ip=integrityByCode.get(e.code); const integrity=ip?.integrityScore??100; return {...e,integrity,trust:Math.round(e.compliance*0.5+integrity*0.5)}; }).sort((a,b)=>a.trust-b.trust),[enriched,integrityByCode]);
  const chronicOffenders=fieldTrust.filter(e=>e.trust<50);
  const statusBuckets=useMemo(()=>[ {name:"Healthy", value:fieldTrust.filter(e=>e.trust>=80).length, color:"#34d399", key:"HEALTHY"}, {name:"At Risk", value:fieldTrust.filter(e=>e.trust<50).length, color:"#f87171", key:"AT RISK"} ],[fieldTrust]);
  const dailyTrend=useMemo(()=>{ const map=new Map<string,number>(); for(const v of visits) map.set(v.date.slice(0,10),(map.get(v.date.slice(0,10))||0)+1); return Array.from(map.entries()).sort((a,b)=>a[0].localeCompare(b[0])).map(([date,count])=>({date:date.slice(5),count})); },[visits]);
  const ghostRanking=useMemo(()=>[...integrityProfiles].sort((a,b)=>b.ghostScore-a.ghostScore),[integrityProfiles]);
  const hasData=enriched.length>0; const hasVisits=visits.length>0; const totalStaff=enriched.length;
  const visitsByEmpCode = useMemo(() => { const m = new Map<string, VisitRow[]>(); for (const v of visits) { if (!m.has(v.spvrCode)) m.set(v.spvrCode, []); m.get(v.spvrCode)!.push(v); } return m; }, [visits]);
  const selectedEmployee = useMemo(() => { if (!selectedEmployeeCode) return null; const e = enriched.find(emp => emp.code === selectedEmployeeCode); if (!e) return null; return { employee: e, profile: integrityByCode.get(e.code), visits: visitsByEmpCode.get(e.code) ?? [] }; }, [selectedEmployeeCode, enriched, integrityByCode, visitsByEmpCode]);

  return (
    <div className="flex bg-neutral-950 min-h-screen">
      <SystemSidePanel visits={visits} employees={employees} files={files} />
      <div className="ml-72 p-6 space-y-6 max-w-full flex-1">
        <CommandBar isOpen={commandBarOpen} onClose={() => setCommandBarOpen(false)} employees={enriched} onSelectEmployee={(code) => setSelectedEmployeeCode(code)} onExport={() => exportCEOReport(enriched, integrityProfiles, visits, audit)} />

        <div className="flex items-start justify-between flex-wrap gap-3"><div><div className="flex items-center gap-3 mb-1"><div className="w-1 h-8 rounded-full" style={{background:"linear-gradient(to bottom,#22d3ee,#a855f7,#f87171)"}}/><h2 className="text-xl font-bold text-white tracking-widest">FIELD INTELLIGENCE</h2></div></div><div className="flex items-center gap-2">{hasData && (<button onClick={() => exportCEOReport(enriched, integrityProfiles, visits, audit)} className="flex items-center gap-1.5 text-xs font-medium text-purple-300 bg-purple-500/10 border border-purple-500/30 px-3 py-1.5 rounded-lg hover:bg-purple-500/20"><Download className="w-3.5 h-3.5"/>CEO Report</button>)}<button onClick={()=>{setFiles([]);setEmployees([]);setVisits([]);localStorage.removeItem('visitComplianceState');}} className="text-xs text-red-400 bg-red-500/10 border border-red-500/30 px-3 py-1.5 rounded-lg hover:bg-red-500/20">Clear</button></div></div>

        <div onDragOver={e=>{e.preventDefault();setIsDragging(true);}} onDragLeave={()=>setIsDragging(false)} onDrop={e=>{e.preventDefault();setIsDragging(false);if(e.dataTransfer.files.length)handleFiles(e.dataTransfer.files);}} className={`relative border-2 border-dashed rounded-xl p-8 text-center transition-all ${isDragging?"border-cyan-500 bg-cyan-500/5":"border-neutral-700 bg-neutral-900"}`}>
          <input type="file" accept=".xlsx,.xls" multiple onChange={e=>{if(e.target.files?.length)handleFiles(e.target.files);}} className="absolute inset-0 opacity-0 cursor-pointer"/>
          <Upload className={`w-8 h-8 mx-auto mb-3 ${isDragging?"text-cyan-400":"text-neutral-500"}`}/><p className="text-sm font-medium text-white">Drop Raw / Follow-Up / Summary .xlsx files here</p>{isProcessing&&<div className="mt-3 flex items-center justify-center gap-2 text-xs text-cyan-400"><RefreshCw className="w-3.5 h-3.5 animate-spin"/>Parsing…</div>}
        </div>

        {hasData&&(<>
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">{[ {label:"HEADCOUNT",val:enriched.length,color:"text-white",icon:Users}, {label:"COMPLIANCE",val:`${overallCompliance}%`,color:compColor(overallCompliance),icon:TrendingUp}, {label:"AT-RISK",val:chronicOffenders.length,color:"text-red-400",icon:ShieldAlert}, {label:"FLAGS",val:audit.length,color:"text-amber-400",icon:Flag} ].map(k=>(<div key={k.label} className="bg-neutral-900 border border-neutral-700 rounded-xl p-4"><div className="flex items-center justify-between mb-2"><span className="text-[10px] text-neutral-400 tracking-wider font-medium">{k.label}</span><k.icon className={`w-4 h-4 ${k.color}`}/></div><p className={`text-2xl font-bold font-mono ${k.color}`}>{k.val}</p></div>))}</div>
          
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5"><div className="flex items-center gap-2 mb-4"><Activity className="w-4 h-4 text-cyan-400"/><p className="text-xs text-neutral-300 tracking-wider font-semibold">DAILY VISIT VOLUME</p></div><ResponsiveContainer width="100%" height={200}><AreaChart data={dailyTrend} margin={{top:8,right:8,bottom:0,left:0}}><defs><Defs/></defs><CartesianGrid strokeDasharray="3 3" stroke="#1c1c1c" vertical={false}/><XAxis dataKey="date" tick={{fill:"#525252",fontSize:9}} axisLine={false} tickLine={false}/><YAxis tick={{fill:"#525252",fontSize:9}} axisLine={false} tickLine={false} width={26}/><Tooltip contentStyle={{background:"#111",border:"1px solid #2a2a2a",borderRadius:10,color:"#fff"}}/><Area type="monotone" dataKey="count" stroke="#22d3ee" strokeWidth={2.5} fill="url(#areaCyan)" dot={<GlowDot/>}/></AreaChart></ResponsiveContainer></div>

          <EmployeePerformanceTracker enriched={enriched} integrityByCode={integrityByCode} fieldTrust={fieldTrust} onSelectEmployee={(code: string) => setSelectedEmployeeCode(code)} />
          {hasVisits && <VisitMapWidget visits={visits} integrityByCode={integrityByCode} />}

          {hasVisits&&(
            <div className="bg-neutral-900 border border-purple-500/20 rounded-xl overflow-hidden">
              <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-700"><div className="flex items-center gap-2"><Eye className="w-4 h-4 text-purple-400"/><span className="text-xs text-purple-300 tracking-wider font-semibold">MANAGER ACTION DESK</span></div><div className="flex gap-1 p-1 bg-neutral-800 border border-neutral-700 rounded-lg">{([["trust","⬇ Lowest Trust"],["ghost","👻 Ghost Activity"]] as const).map(([id,label])=>(<button key={id} onClick={()=>setLowestTab(id)} className={`text-[10px] px-3 py-1.5 rounded font-bold tracking-wider ${lowestTab===id?id==="ghost"?"bg-red-500/20 text-red-300 border border-red-500/30":"bg-purple-500/20 text-purple-300 border border-purple-500/30":"text-neutral-500 hover:text-neutral-300"}`}>{label}</button>))}</div></div>
              
              {selectedForAction.size > 0 && (
                <div className="bg-cyan-900/20 border-b border-cyan-800/30 p-3 flex items-center justify-between">
                  <span className="text-xs text-cyan-400 font-bold">{selectedForAction.size} Employee(s) Selected</span>
                  <button onClick={handleBulkEmail} className="flex items-center gap-2 text-xs bg-cyan-600 text-white px-3 py-1.5 rounded-lg hover:bg-cyan-500"><Send className="w-3 h-3"/> Draft Bulk Email</button>
                </div>
              )}

              {lowestTab==="trust"&&(<div className="overflow-x-auto"><table className="w-full text-sm"><thead><tr className="text-[10px] text-neutral-500 border-b border-neutral-800 bg-neutral-950/40"><th className="py-2.5 px-3 w-10"></th>{["EMPLOYEE","VISITS","INTEGRITY","TRUST","WHY FOLLOW UP"].map(h=>(<th key={h} className="text-left py-2.5 px-3 font-semibold tracking-wider whitespace-nowrap">{h}</th>))}</tr></thead><tbody>{fieldTrust.slice(0,15).map(e=>{const ip=integrityByCode.get(e.code); if(!ip) return null; return (<tr key={e.code} onClick={() => setSelectedEmployeeCode(e.code)} className={`border-b border-neutral-800 hover:bg-neutral-800/30 transition-colors cursor-pointer ${e.trust<40?"bg-red-950/10":""}`}><td className="py-2.5 px-3" onClick={(ev) => ev.stopPropagation()}><input type="checkbox" checked={selectedForAction.has(e.code)} onChange={() => toggleSelection(e.code)} className="w-4 h-4 bg-neutral-700 border-neutral-600 rounded focus:ring-cyan-500" /></td><td className="py-2.5 px-3"><p className="text-white text-xs font-semibold">{e.name}</p><p className="text-cyan-400 font-mono text-[10px]">{e.code}</p></td><td className="py-2.5 px-3 font-mono text-neutral-300 text-xs text-center">{ip.totalVisits}</td><td className={`py-2.5 px-3 text-center font-mono font-bold text-xs ${intgColor(ip.integrityScore)}`}>{ip.integrityScore}%</td><td className="py-2.5 px-3 text-center"><span className={`text-xs font-bold font-mono px-2 py-0.5 rounded-full border ${e.trust<50?"bg-red-500/15 text-red-300 border-red-500/30":"bg-emerald-500/15 text-emerald-300 border-emerald-500/30"}`}>{e.trust}%</span></td><td className="py-2.5 px-3 min-w-[180px]"><WhyCell e={e} ip={ip}/></td></tr>);})}</tbody></table></div>)}
              {lowestTab==="ghost"&&(<div className="overflow-x-auto"><table className="w-full text-sm"><thead><tr className="text-[10px] text-neutral-500 border-b border-neutral-800 bg-neutral-950/40"><th className="py-2.5 px-3 w-10"></th>{["EMPLOYEE","GHOST SCORE","RISK","FLAGS"].map(h=>(<th key={h} className="text-left py-2.5 px-3 font-semibold tracking-wider whitespace-nowrap">{h}</th>))}</tr></thead><tbody>{ghostRanking.slice(0,15).map(ip=>{const badge2=ghostBadge(ip.ghostScore); return (<tr key={ip.code} onClick={() => setSelectedEmployeeCode(ip.code)} className={`border-b border-neutral-800 hover:bg-neutral-800/30 transition-colors cursor-pointer ${ip.ghostScore>=60?"bg-red-950/10":""}`}><td className="py-2.5 px-3" onClick={(ev) => ev.stopPropagation()}><input type="checkbox" checked={selectedForAction.has(ip.code)} onChange={() => toggleSelection(ip.code)} className="w-4 h-4 bg-neutral-700 border-neutral-600 rounded focus:ring-cyan-500" /></td><td className="py-2.5 px-3"><p className="text-white text-xs font-semibold">{ip.name}</p><p className="text-cyan-400 font-mono text-[10px]">{ip.code}</p></td><td className="py-2.5 px-3"><GhostMeter score={ip.ghostScore}/></td><td className="py-2.5 px-3"><span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${badge2.bg} ${badge2.text} ${badge2.border}`}>{badge2.label}</span></td><td className="py-2.5 px-3"><FlagTooltip profile={ip}/></td></tr>);})}</tbody></table></div>)}
            </div>
          )}

          <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
            <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-700 flex-wrap gap-2"><span className="text-xs text-neutral-300 tracking-wider font-semibold">EMPLOYEE-LEVEL DETAIL</span><select value={deptFilter} onChange={e=>setDeptFilter(e.target.value)} className="bg-neutral-800 border border-neutral-600 text-xs text-neutral-300 rounded px-3 py-1.5 focus:outline-none focus:border-cyan-500">{departments.map(d=><option key={d}>{d}</option>)}</select></div>
            <div className="overflow-x-auto max-h-[480px] overflow-y-auto"><table className="w-full text-sm"><thead><tr className="text-[10px] text-neutral-500 border-b border-neutral-800 sticky top-0 bg-neutral-900 z-10"><th className="py-2.5 px-3 w-10"></th>{["EMPLOYEE","COMP","INTG","TREND","STATUS"].map(h=>(<th key={h} className="text-left py-2.5 px-3 font-semibold tracking-wider whitespace-nowrap">{h}</th>))}</tr></thead><tbody>{filtered.map(e=>{const ip=integrityByCode.get(e.code); const trust=Math.round(e.compliance*0.5+(ip?.integrityScore??100)*0.5); return (<tr key={e.code} onClick={() => setSelectedEmployeeCode(e.code)} className="border-b border-neutral-800 hover:bg-neutral-800/30 transition-colors cursor-pointer"><td className="py-2 px-3" onClick={(ev) => ev.stopPropagation()}><input type="checkbox" checked={selectedForAction.has(e.code)} onChange={() => toggleSelection(e.code)} className="w-4 h-4 bg-neutral-700 border-neutral-600 rounded" /></td><td className="py-2 px-3 text-white text-xs font-medium">{e.name}<span className="text-neutral-500 ml-2">{e.code}</span></td><td className={`py-2 px-3 text-center font-mono font-bold text-xs ${compColor(e.compliance)}`}>{e.compliance}%</td><td className={`py-2 px-3 text-center font-mono font-bold text-xs ${intgColor(ip?.integrityScore??100)}`}>{ip?.integrityScore??100}%</td><td className="py-2 px-3">{ip && ip.dailyTrend.length > 1 && <SparklineChart data={ip.dailyTrend} color={trust < 50 ? "#f87171" : "#34d399"} />}</td><td className="py-2 px-3">{trust<50?<span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-red-500/15 text-red-300 border border-red-500/30">RISK</span>:<span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-500/15 text-emerald-300 border border-emerald-500/30">OK</span>}</td></tr>);})}</tbody></table></div>
          </div>
        </>)}
      </div>
      {selectedEmployee && <EmployeeScorecardModal employee={selectedEmployee.employee} profile={selectedEmployee.profile} visits={selectedEmployee.visits} onClose={() => setSelectedEmployeeCode(null)} />}
    </div>
  );
}