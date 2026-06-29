"use client";
import { useState, useCallback, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Upload, FileSpreadsheet, CheckCircle, AlertTriangle, XCircle,
  Users, ShieldCheck, ShieldAlert, TrendingDown, TrendingUp,
  Layers, RefreshCw, Crown, Flag, Copy, FileWarning, Eye, Gauge,
  Zap, Clock, Target, Activity, Ghost, MapPin, Download, X,
  FileDown, MessageSquare, ChevronRight, Mail, Printer, FastForward, Radar, Search
} from "lucide-react";
import {
  AreaChart, Area, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, ResponsiveContainer, Tooltip, CartesianGrid, ReferenceLine,
  LabelList, LineChart, Line
} from "recharts";
import { MapContainer, TileLayer, CircleMarker, Popup } from 'react-leaflet';
import 'leaflet/dist/leaflet.css';
import html2canvas from 'html2canvas';

// ─── TYPES ──────────────────────────────────────────────────────────────────
type VisitFileKind = "visit_raw" | "visit_followup" | "visit_summary" | "unknown";
interface LoadedFile { name: string; kind: VisitFileKind; size: string; }
interface EmployeeRecord {
  code: string; name: string; department: string;
  daysReported: number; daysPossible: number;
}
interface VisitRow {
  date: string; spvrCode: string; spvrName: string; shopCode: string;
  shopName: string; area: string; governorate: string;
  samsungShortage: string; compShortage: string; selloutMovement: string;
  brand: string; movement: string; comment: string;
  action1: string; accountFeedback: string; action2: string;
  lat: number; lon: number;
}
interface IntegrityProfile {
  code: string; name: string;
  totalVisits: number; uniqueShops: number;
  commentedVisits: number; blankComments: number; lowInfoComments: number;
  templatedRepeats: number;
  topRepeatedComment: { text: string; count: number } | null;
  uniquenessRatio: number; unsupportedClaims: number;
  integrityScore: number;
  avgCommentLen: number;
  lateNightCount: number;
  singleShopLoop: boolean;
  ghostScore: number;
  suspiciousFlags: string[];
  // NEW enriched fields & Fraud/NLP
  teleportationFlags: string[];
  rushHourFlags: string[];
  commentTopics: string[];
  avgSentiment: number;
  mutatedCopyPastes: number;
  dailyTrend: { date: string; count: number }[];
}

const LOW_INFO_SET = new Set([
  "ok","fine","good","done","no issue","nothing","n/a","na","-","none","good.","ok."
]);
const fmtBytes = (b: number) =>
  b > 1_048_576 ? `${(b/1_048_576).toFixed(1)} MB` : `${(b/1024).toFixed(0)} KB`;
const pct = (a: number, b: number) => b > 0 ? Math.round((a/b)*100) : 0;

// ─── NEW HELPERS: NLP & FRAUD ───────────────────────────────────────────────
function hashStrToCoord(str: string): [number, number] {
  let h1 = 0, h2 = 0;
  for (let i = 0; i < str.length; i++) {
    h1 = ((h1 << 5) - h1 + str.charCodeAt(i)) | 0;
    if (i > 0) h2 = ((h2 << 5) - h2 + str.charCodeAt(i)) | 0;
  }
  return [30.0 + (Math.abs(h1) % 100) / 50.0, 31.0 + (Math.abs(h2) % 100) / 50.0];
}

function calculateDistance(lat1: number, lon1: number, lat2: number, lon2: number): number {
  const R = 6371; const dLat = (lat2 - lat1) * Math.PI / 180; const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) + Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * Math.sin(dLon/2) * Math.sin(dLon/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function levenshtein(a: string, b: string): number {
  const tmp=[]; let i, j, prev;
  for (i = 0; i <= a.length; i++) tmp[i] = [i];
  for (j = 0; j <= b.length; j++) tmp[0][j] = j;
  for (i = 1; i <= a.length; i++) { for (j = 1; j <= b.length; j++) { prev = tmp[i - 1][j - 1]; if (a[i - 1] !== b[j - 1]) prev = Math.min(prev, tmp[i][j - 1], tmp[i - 1][j]) + 1; tmp[i][j] = prev; } }
  return tmp[a.length][b.length];
}

const posWords = new Set(["good","great","excellent","positive","satisfied","happy","clean","organized","active","sold","helped"]);
const negWords = new Set(["bad","poor","terrible","negative","angry","dirty","messy","out of stock","shortage","rude","issue","problem","missing"]);
function analyzeSentiment(text: string): number {
  const words = text.toLowerCase().split(/\s+/); if (words.length === 0) return 0;
  let score = 0; words.forEach(w => { if(posWords.has(w)) score++; if(negWords.has(w)) score--; });
  return Math.max(-1, Math.min(1, score / words.length));
}

function extractTopics(text: string): string[] {
  const t = text.toLowerCase(); const topics: string[] = [];
  if (t.includes("out of stock") || t.includes("shortage") || t.includes("oos")) topics.push("Shortage");
  if (t.includes("display") || t.includes("planogram") || t.includes("shelf")) topics.push("Display");
  if (t.includes("competitor") || t.includes("promo") || t.includes("offer")) topics.push("Competitor");
  if (t.includes("training") || t.includes("staff") || t.includes("employee")) topics.push("Staff");
  if (t.includes("sales") || t.includes("sellout") || t.includes("revenue")) topics.push("Sales");
  if (topics.length === 0 && text.length > 10) topics.push("General Ops");
  return topics;
}

// ─── ORIGINAL PARSING LOGIC ─────────────────────────────────────────────────
function detectKind(sheetNames: string[]): VisitFileKind {
  if (sheetNames.includes("Mobile")) return "visit_raw";
  if (sheetNames.includes("Submission Matrix")) return "visit_followup";
  if (sheetNames.includes("Weekly Summary")) return "visit_summary";
  return "unknown";
}
function parseSummary(rows: unknown[][]): EmployeeRecord[] {
  const headerRow = (rows[1]||[]) as unknown[];
  const dayCols: number[] = [];
  for (let c=3;c<headerRow.length;c++){
    const h=String(headerRow[c]??"");
    if(h&&!h.includes("Total")&&!h.toLowerCase().includes("notes")) dayCols.push(c);
  }
  const out: EmployeeRecord[] = [];
  for(let r=2;r<rows.length;r++){
    const row=rows[r] as unknown[];
    const name=row?.[0],code=row?.[1],dept=row?.[2];
    if(!name||!code) continue;
    let reported=0;
    for(const c of dayCols){
      const v=row[c];
      if(v!==null&&v!==undefined&&v!==""&&Number(v)>0) reported++;
    }
    out.push({code:String(code),name:String(name),department:String(dept??"Unassigned"),daysReported:reported,daysPossible:dayCols.length});
  }
  return out;
}
function parseFollowUp(rows: unknown[][]): Map<string,{total:number;missing:string[]}> {
  const map = new Map<string,{total:number;missing:string[]}>();
  for(let r=1;r<rows.length;r++){
    const row=rows[r] as unknown[];
    const code=row?.[1]; if(!code) continue;
    const total=Number(row?.[2])||0;
    const missingStr=String(row?.[3]??"");
    const missing=missingStr==="–"||!missingStr?[]:missingStr.split(",").map(s=>s.trim()).filter(Boolean);
    map.set(String(code),{total,missing});
  }
  return map;
}
function parseRawVisits(rows: unknown[][]): VisitRow[] {
  const out: VisitRow[] = [];
  for(let r=3;r<rows.length;r++){
    const row=rows[r] as unknown[];
    if(!row?.[1]) continue;
    const areaStr = String(row[5]??"Unknown");
    const shopStr = String(row[4]??row[3]??"");
    let lat = 0, lon = 0;
    const coordMatch = areaStr.match(/(-?\d+\.\d+),\s*(-?\d+\.\d+)/);
    if (coordMatch) {
      lat = parseFloat(coordMatch[1]); lon = parseFloat(coordMatch[2]);
    } else {
      [lat, lon] = hashStrToCoord(areaStr + shopStr);
    }
    out.push({
      date:String(row[0]??""),spvrCode:String(row[1]??""),spvrName:String(row[2]??""),shopCode:String(row[3]??""),shopName:shopStr,area:areaStr,governorate:String(row[6]??""),
      samsungShortage:String(row[7]??""),compShortage:String(row[8]??""),selloutMovement:String(row[9]??""),brand:String(row[10]??""),
      movement:String(row[11]??""),comment:String(row[12]??""),action1:String(row[13]??""),accountFeedback:String(row[14]??""),action2:String(row[15]??""),
      lat, lon
    });
  }
  return out;
}

function getHour(dateStr: string): number|null {
  const m = dateStr.match(/[T ](\d{2}):(\d{2})/);
  return m ? parseInt(m[1],10) : null;
}

// ─── ENHANCED INTEGRITY BUILDER ─────────────────────────────────────────────
function buildIntegrityProfiles(visits: VisitRow[], employees: EmployeeRecord[]): IntegrityProfile[] {
  const byEmp = new Map<string,VisitRow[]>();
  for(const v of visits){ if(!byEmp.has(v.spvrCode)) byEmp.set(v.spvrCode,[]); byEmp.get(v.spvrCode)!.push(v); }
  const nameByCode = new Map(employees.map(e=>[e.code,e.name]));
  const profiles: IntegrityProfile[] = [];

  for(const [code,rows] of byEmp.entries()){
    const sortedRows = [...rows].sort((a,b) => a.date.localeCompare(b.date));
    const shops = new Set(rows.map(r=>r.shopCode));
    let blank=0,lowInfo=0,unsupported=0,lateNight=0,totalLen=0;
    const commentCounts = new Map<string,number>();
    const allComments: string[] = [];
    const allTopicsSet = new Set<string>();
    let totalSentiment = 0, commentsWithSentiment = 0;
    const dailyMap = new Map<string, number>();
    
    let teleportationFlags: string[] = [];
    let rushHourFlags: string[] = [];

    for(const r of rows) { const d = r.date.slice(0,10); dailyMap.set(d, (dailyMap.get(d)||0) + 1); }
    const dailyTrend = Array.from(dailyMap.entries()).map(([date, count]) => ({date:date.slice(5), count})).sort((a,b)=>a.date.localeCompare(b.date));

    for(let i=0; i<sortedRows.length; i++){
      const curr = sortedRows[i];
      const currTime = new Date(curr.date).getTime();
      
      if(i < sortedRows.length - 1){
        const next = sortedRows[i+1]; const nextTime = new Date(next.date).getTime();
        const timeDiffHrs = Math.abs(nextTime - currTime) / (1000 * 3600);
        if(timeDiffHrs > 0 && timeDiffHrs < 12) { 
          const distKm = calculateDistance(curr.lat, curr.lon, next.lat, next.lon);
          const speedKmH = distKm / timeDiffHrs;
          if(speedKmH > 150){ teleportationFlags.push(`Traveled ${Math.round(distKm)}km in ${Math.round(timeDiffHrs*60)}m`); }
        }
      }

      const hour = getHour(curr.date);
      if(hour !== null && hour >= 8 && hour <= 18) {
        const visitsThisHour = sortedRows.filter(r => getHour(r.date) === hour && r.date.slice(0,10) === curr.date.slice(0,10)).length;
        if(visitsThisHour > 5) { if(!rushHourFlags.some(f => f.includes(curr.date.slice(0,10)))) rushHourFlags.push(`${visitsThisHour} visits in 1h on ${curr.date.slice(0,10)}`); }
      }

      const c=curr.comment.toLowerCase(); totalLen+=curr.comment.length;
      if(!c) blank++;
      else {
        if(LOW_INFO_SET.has(c)||c.length<=3) lowInfo++;
        else {
          commentCounts.set(c,(commentCounts.get(c)||0)+1); allComments.push(c);
          const sentiment = analyzeSentiment(curr.comment); if(sentiment !== 0) { totalSentiment += sentiment; commentsWithSentiment++; }
          extractTopics(curr.comment).forEach(t => allTopicsSet.add(t));
        }
      }

      const shortage=curr.samsungShortage&&!["n","no","none",""].includes(curr.samsungShortage.toLowerCase());
      if((shortage||curr.selloutMovement.length>0)&&!c) unsupported++;
      if(hour!==null&&(hour>=23||hour<=4)) lateNight++;
    }

    let mutatedCopyPastes = 0;
    for(let i=0; i<allComments.length; i++){
      for(let j=i+1; j<allComments.length; j++){
        if(allComments[i] === allComments[j]) continue;
        const maxLen = Math.max(allComments[i].length, allComments[j].length); if(maxLen === 0) continue;
        if(1 - (levenshtein(allComments[i], allComments[j]) / maxLen) > 0.85) mutatedCopyPastes++;
      }
    }

    const commented=rows.length-blank; const uniqueComments=commentCounts.size+lowInfo;
    let templatedRepeats=0; let top: {text:string;count:number}|null=null;
    for(const [text,count] of commentCounts.entries()){ if(count>1) templatedRepeats+=count-1; if(!top||count>top.count) top={text,count}; }
    
    const uniquenessRatio=commented>0?pct(uniqueComments,commented):100;
    const avgCommentLen=rows.length>0?Math.round(totalLen/rows.length):0;
    const singleShopLoop=shops.size===1&&rows.length>4;
    const avgSentiment = commentsWithSentiment > 0 ? totalSentiment / commentsWithSentiment : 0;

    const suspiciousFlags: string[] = []; const blankPct=pct(blank,rows.length);
    if(blankPct>30) suspiciousFlags.push(`${blank} blank comments (${blankPct}% of visits)`);
    if(templatedRepeats>=3) suspiciousFlags.push(`Same comment copy-pasted ${templatedRepeats} extra times`);
    if(mutatedCopyPastes > 2) suspiciousFlags.push(`${mutatedCopyPastes} mutated copy-pastes detected`);
    const latePct=pct(lateNight,rows.length);
    if(latePct>50&&rows.length>=4) suspiciousFlags.push(`${lateNight}/${rows.length} reports filed 23:00–04:00 (batch backdating?)`);
    if(unsupported>0) suspiciousFlags.push(`${unsupported} shortage/movement claims with zero explanation`);
    if(avgCommentLen<25&&rows.length>=5) suspiciousFlags.push(`Avg comment only ${avgCommentLen} chars — minimal effort`);
    if(singleShopLoop) suspiciousFlags.push(`All ${rows.length} visits logged at the same single shop`);
    if(top&&top.count>=4) suspiciousFlags.push(`"${top.text.slice(0,55)}..." repeated ${top.count}×`);
    if(teleportationFlags.length > 0) suspiciousFlags.push(`Impossible travel speed detected (${teleportationFlags.length} instances)`);
    if(rushHourFlags.length > 0) suspiciousFlags.push(`Rush-hour filing (>5 shops/hour)`);
    if(avgSentiment < -0.3) suspiciousFlags.push(`Highly negative comment sentiment`);

    const ghostScore=Math.min(100,Math.round(
      blankPct*0.30+Math.min(100,templatedRepeats*8)*0.20+latePct*0.25+Math.min(100,unsupported*15)*0.15+(singleShopLoop?20:0)*0.10
    ));
    const penalty=templatedRepeats*4+lowInfo*1.5+unsupported*3+blank*1+mutatedCopyPastes*5+teleportationFlags.length*15+rushHourFlags.length*8;
    const integrityScore=Math.max(0,Math.min(100,Math.round(100-penalty)));

    profiles.push({
      code,name:nameByCode.get(code)||code,
      totalVisits:rows.length,uniqueShops:shops.size,
      commentedVisits:commented,blankComments:blank,lowInfoComments:lowInfo,
      templatedRepeats,topRepeatedComment:top&&top.count>1?top:null,
      uniquenessRatio,unsupportedClaims:unsupported,integrityScore,
      avgCommentLen,lateNightCount:lateNight,singleShopLoop,
      ghostScore,suspiciousFlags,
      teleportationFlags, rushHourFlags, commentTopics: Array.from(allTopicsSet), avgSentiment, mutatedCopyPastes, dailyTrend
    });
  }
  return profiles.sort((a,b)=>a.integrityScore-b.integrityScore);
}

// ─── COLOR HELPERS ──────────────────────────────────────────────────────────
const compColor = (p:number) => p>=80?"text-emerald-400":p>=50?"text-amber-400":"text-red-400";
const intgColor = (p:number) => p>=80?"text-cyan-400":p>=50?"text-amber-400":"text-red-400";
const ghostBadge = (s:number) =>
  s>=60?{label:"HIGH RISK",bg:"bg-red-500/20",text:"text-red-400",border:"border-red-500/30"}:
  s>=30?{label:"SUSPICIOUS",bg:"bg-amber-500/20",text:"text-amber-400",border:"border-amber-500/30"}:
       {label:"CLEAN",bg:"bg-emerald-500/20",text:"text-emerald-400",border:"border-emerald-500/30"};

// ─── SVG GRADIENT DEFS ──────────────────────────────────────────────────────
function Defs() {
  return (
    <defs>
      <linearGradient id="areaCyan" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%"   stopColor="#22d3ee" stopOpacity={0.55}/>
        <stop offset="55%"  stopColor="#06b6d4" stopOpacity={0.12}/>
        <stop offset="100%" stopColor="#22d3ee" stopOpacity={0}/>
      </linearGradient>
      <linearGradient id="hbarCyan" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%"   stopColor="#0e7490" stopOpacity={0.7}/>
        <stop offset="100%" stopColor="#22d3ee" stopOpacity={1}/>
      </linearGradient>
      <linearGradient id="hbarPurple" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%"   stopColor="#5b21b6" stopOpacity={0.7}/>
        <stop offset="100%" stopColor="#a78bfa" stopOpacity={1}/>
      </linearGradient>
      <linearGradient id="hbarAmber" x1="0" y1="0" x2="1" y2="0">
        <stop offset="0%"   stopColor="#92400e" stopOpacity={0.7}/>
        <stop offset="100%" stopColor="#fbbf24" stopOpacity={1}/>
      </linearGradient>
      <filter id="shadow" height="140%" width="140%">
        <feDropShadow dx="0" dy="3" stdDeviation="4" floodColor="#000" floodOpacity={0.55}/>
      </filter>
      <filter id="glowCyan">
        <feDropShadow dx="0" dy="0" stdDeviation="4" floodColor="#22d3ee" floodOpacity={0.55}/>
      </filter>
    </defs>
  );
}

const GlowDot = (props: any) => {
  const {cx,cy} = props;
  if(!cx||!cy) return null;
  return (
    <g>
      <circle cx={cx} cy={cy} r={7} fill="#22d3ee" opacity={0.15}/>
      <circle cx={cx} cy={cy} r={3.5} fill="#22d3ee" filter="url(#glowCyan)"/>
    </g>
  );
};

const GradBar = (props: any) => {
  const {x,y,width,height,index} = props;
  const grads = ["url(#hbarCyan)","url(#hbarPurple)","url(#hbarAmber)"];
  const fill = grads[index%grads.length];
  return (
    <g filter="url(#shadow)">
      <rect x={x} y={y+1} width={width} height={height-2} rx={5} ry={5} fill={fill}/>
      <rect x={x} y={y+1} width={width*0.38} height={(height-2)*0.4} rx={5} ry={5} fill="white" opacity={0.09}/>
    </g>
  );
};

function GhostMeter({score}:{score:number}) {
  const filled=Math.round(score/10);
  const color=score>=60?"bg-red-400":score>=30?"bg-amber-400":"bg-emerald-400";
  return (
    <div className="flex items-center gap-0.5">
      {Array.from({length:10}).map((_,i)=>(
        <div key={i} className={`w-1.5 h-3.5 rounded-sm ${i<filled?color:"bg-neutral-700"}`}/>
      ))}
    </div>
  );
}

function FlagTooltip({profile}:{profile:IntegrityProfile}) {
  const [open,setOpen]=useState(false);
  const n=profile.suspiciousFlags.length;
  if(n===0) return <span className="text-neutral-600 text-xs">—</span>;
  return (
    <div className="relative inline-block">
      <button
        onMouseEnter={()=>setOpen(true)}
        onMouseLeave={()=>setOpen(false)}
        className="flex items-center gap-1 text-xs font-medium text-amber-400 hover:text-amber-200 transition-colors"
      >
        <Zap className="w-3 h-3"/>
        {n} flag{n>1?"s":""}
      </button>
      {open&&(
        <div className="absolute z-50 bottom-full left-0 mb-2 w-80 bg-neutral-800 border border-amber-500/40 rounded-xl p-3 shadow-2xl shadow-black/60">
          <p className="text-[10px] text-amber-300 font-bold mb-2 tracking-widest">⚠ SUSPICIOUS SIGNALS DETECTED</p>
          <div className="space-y-1.5">
            {profile.suspiciousFlags.map((f,i)=>(
              <div key={i} className="flex items-start gap-2 text-xs text-neutral-200">
                <span className="w-1.5 h-1.5 rounded-full bg-amber-400 mt-1 shrink-0"/>
                {f}
              </div>
            ))}
          </div>
          {profile.topRepeatedComment&&(
            <div className="mt-2.5 pt-2 border-t border-neutral-700">
              <p className="text-[10px] text-neutral-500 mb-1">Top repeated comment ({profile.topRepeatedComment.count}×):</p>
              <p className="text-[10px] text-neutral-300 italic line-clamp-2">"{profile.topRepeatedComment.text}"</p>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function WhyCell({e,ip}:{e:any;ip:IntegrityProfile|undefined}) {
  const reasons: string[] = [];
  if(e.compliance<50) reasons.push(`Only ${e.compliance}% submission rate`);
  if(ip){
    if(ip.blankComments>0) reasons.push(`${ip.blankComments} visits with no comment`);
    if(ip.templatedRepeats>=3) reasons.push("Copy-pastes same comment repeatedly");
    if(ip.lateNightCount>ip.totalVisits*0.5&&ip.totalVisits>=4) reasons.push("Batch-files reports late at night");
    if(ip.unsupportedClaims>0) reasons.push(`${ip.unsupportedClaims} unsubstantiated shortage claims`);
    if(ip.singleShopLoop) reasons.push("All visits at one shop only");
    if(ip.teleportationFlags.length>0) reasons.push("Impossible travel speed");
    if(ip.rushHourFlags.length>0) reasons.push("Rush-hour filing");
  }
  if(reasons.length===0) return <span className="text-neutral-600 text-xs">—</span>;
  return (
    <div className="flex items-start justify-between gap-2">
      <ul className="space-y-0.5">
        {reasons.map((r,i)=>(
          <li key={i} className="text-[10px] text-neutral-300 flex items-start gap-1">
            <span className="text-red-400 shrink-0">•</span>{r}
          </li>
        ))}
      </ul>
      <button onClick={(ev) => {
          ev.stopPropagation();
          const subject = encodeURIComponent(`Action Required: Visit Compliance - ${e.name}`);
          const body = encodeURIComponent(`Dear ${e.name},\n\nFlags:\n- ${reasons.join('\n- ')}\n\nPlease resolve.`);
          window.open(`mailto:?subject=${subject}&body=${body}`);
        }} className="shrink-0 text-neutral-500 hover:text-purple-400 transition-colors mt-0.5" title="Draft Email">
        <Mail className="w-3.5 h-3.5"/>
      </button>
    </div>
  );
}

// ─── NEW UI COMPONENTS ──────────────────────────────────────────────────────
function SparklineChart({ data, color }: { data: {count:number}[], color: string }) {
  if (!data || data.length < 2) return <span className="text-neutral-700 text-[10px]">—</span>;
  return (
    <ResponsiveContainer width={50} height={20}>
      <LineChart data={data} margin={{ top: 2, right: 2, bottom: 2, left: 2 }}>
        <Line type="monotone" dataKey="count" stroke={color} strokeWidth={1.5} dot={false} />
      </LineChart>
    </ResponsiveContainer>
  );
}

function EmployeeScorecardModal({ employee, profile, visits, onClose }: { employee: any; profile: IntegrityProfile | undefined; visits: VisitRow[]; onClose: () => void }) {
  const integrity = profile?.integrityScore ?? 100;
  const trust = Math.round(employee.compliance * 0.5 + integrity * 0.5);
  const badge = ghostBadge(profile?.ghostScore ?? 0);
  const complianceColor = employee.compliance >= 80 ? "#34d399" : employee.compliance >= 50 ? "#fbbf24" : "#f87171";
  const integrityColor = integrity >= 80 ? "#22d3ee" : integrity >= 50 ? "#fbbf24" : "#f87171";
  const trustColor = trust >= 80 ? "#a78bfa" : trust >= 50 ? "#fbbf24" : "#f87171";
  const commentHistory = visits.filter(v => v.comment.trim().length > 0).slice(0, 15);
  const pipRef = useRef<HTMLDivElement>(null);
  
  const handleGeneratePIP = async () => {
    if (!pipRef.current) return;
    const canvas = await html2canvas(pipRef.current, { backgroundColor: '#111' });
    const url = canvas.toDataURL('image/png');
    const a = document.createElement('a'); a.href = url; a.download = `PIP-${employee.name}.png`; a.click();
  };

  useEffect(() => { const handler = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); }; document.addEventListener("keydown", handler); return () => document.removeEventListener("keydown", handler); }, [onClose]);

  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4" onClick={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="absolute inset-0 bg-black/70 backdrop-blur-sm" />
      <div className="relative z-10 bg-neutral-900 border border-neutral-700 rounded-2xl shadow-2xl shadow-black/80 w-full max-w-4xl max-h-[90vh] overflow-y-auto">
        <div className="sticky top-0 z-20 bg-neutral-900/95 backdrop-blur-md border-b border-neutral-700 px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-gradient-to-br from-cyan-500/20 to-purple-500/20 border border-cyan-500/30 flex items-center justify-center"><span className="text-lg font-bold text-cyan-400">{employee.name.charAt(0)}</span></div>
            <div><h3 className="text-white font-bold text-lg">{employee.name}</h3><div className="flex items-center gap-3 mt-0.5"><span className="text-cyan-400 font-mono text-xs">{employee.code}</span><span className="text-neutral-600">·</span><span className="text-neutral-400 text-xs">{employee.department}</span></div></div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={handleGeneratePIP} className="flex items-center gap-1.5 text-xs font-medium text-red-300 bg-red-500/10 border border-red-500/30 px-3 py-1.5 rounded-lg hover:bg-red-500/20"><Printer className="w-3.5 h-3.5"/> Generate PIP</button>
            <button onClick={onClose} className="w-8 h-8 rounded-lg bg-neutral-800 border border-neutral-600 flex items-center justify-center text-neutral-400 hover:text-white"><X className="w-4 h-4" /></button>
          </div>
        </div>
        <div className="p-6 space-y-6">
          <div ref={pipRef} className="absolute left-[-9999px] top-[-9999px] w-[800px] p-8 bg-neutral-900 text-white space-y-4">
            <h1 className="text-2xl font-bold border-b border-neutral-700 pb-2">Performance Improvement Plan</h1>
            <div className="grid grid-cols-2 gap-4 text-sm">
              <p><strong>Employee:</strong> {employee.name} ({employee.code})</p><p><strong>Department:</strong> {employee.department}</p>
              <p><strong>Compliance:</strong> {employee.compliance}%</p><p><strong>Integrity:</strong> {integrity}%</p>
              <p><strong>Ghost Score:</strong> {profile?.ghostScore}/100</p><p><strong>Trust Score:</strong> {trust}%</p>
            </div>
            <div><h3 className="font-bold text-red-400 mb-1">Key Infractions:</h3><ul className="list-disc pl-5 text-neutral-300">{profile?.suspiciousFlags.slice(0,3).map(f => <li key={f}>{f}</li>)}</ul></div>
            {profile?.topRepeatedComment && <p className="text-sm italic text-neutral-400">Worst Offense: "{profile.topRepeatedComment.text}"</p>}
          </div>

          <div className="flex flex-wrap items-center justify-center gap-8">
            {[ { val: employee.compliance, label: "Compliance", color: complianceColor }, { val: integrity, label: "Integrity", color: integrityColor }, { val: trust, label: "Trust", color: trustColor } ].map(g => (
              <div key={g.label} className="flex flex-col items-center gap-2">
                <div style={{ width: 130, height: 130, borderRadius: "50%", background: `conic-gradient(from 135deg, ${g.color} ${(g.val/100)*360}deg, rgba(30,30,42,0.9) ${(g.val/100)*360}deg)`, boxShadow: `0 10px 40px rgba(0,0,0,0.7), 0 0 24px ${g.color}22, inset 0 3px 6px rgba(255,255,255,0.04)`, padding: 12 }}>
                  <div style={{ width: "100%", height: "100%", borderRadius: "50%", background: "radial-gradient(circle at 42% 38%, #1e1e30 0%, #0c0c14 80%)", boxShadow: "inset 0 3px 12px rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column" }}>
                    <span style={{ fontSize: 22, fontWeight: "bold", fontFamily: "monospace", color: g.color }}>{g.val}%</span>
                    <span style={{ fontSize: 8, color: "#737373", textTransform: "uppercase", letterSpacing: 2, marginTop: 4 }}>{g.label}</span>
                  </div>
                </div>
              </div>
            ))}
          </div>

          {profile && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="bg-neutral-950 border border-neutral-800 rounded-xl p-4">
                <div className="flex items-center justify-between mb-3"><div className="flex items-center gap-2"><Ghost className="w-4 h-4 text-purple-400" /><span className="text-xs text-neutral-300 tracking-wider font-bold">GHOST & FRAUD</span></div><span className={`text-xs font-bold px-2.5 py-1 rounded-full border ${badge.bg} ${badge.text} ${badge.border}`}>{badge.label} — {profile.ghostScore}</span></div>
                <div className="space-y-2 mt-3">
                  {profile.teleportationFlags.length > 0 && <div className="bg-red-950/30 border border-red-900/40 rounded-lg p-2"><p className="text-[10px] text-red-400 font-bold tracking-widest mb-1">🚀 IMPOSSIBLE TRAVEL</p>{profile.teleportationFlags.slice(0,2).map((f,i)=><p key={i} className="text-xs text-red-200 ml-2">• {f}</p>)}</div>}
                  {profile.rushHourFlags.length > 0 && <div className="bg-amber-950/30 border border-amber-900/40 rounded-lg p-2"><p className="text-[10px] text-amber-400 font-bold tracking-widest mb-1">⚡ RUSH HOUR FILING</p>{profile.rushHourFlags.slice(0,2).map((f,i)=><p key={i} className="text-xs text-amber-200 ml-2">• {f}</p>)}</div>}
                  {profile.mutatedCopyPastes > 0 && <div className="bg-purple-950/30 border border-purple-900/40 rounded-lg p-2"><p className="text-[10px] text-purple-400 font-bold tracking-widest mb-1">🧬 MUTATED COPY-PASTE</p><p className="text-xs text-purple-200 ml-2">• {profile.mutatedCopyPastes} near-identical comments</p></div>}
                </div>
              </div>
              <div className="bg-neutral-950 border border-neutral-800 rounded-xl p-4">
                <div className="flex items-center gap-2 mb-3"><MessageSquare className="w-4 h-4 text-cyan-400" /><span className="text-xs text-neutral-300 tracking-wider font-bold">NLP & SENTIMENT</span></div>
                <div className="grid grid-cols-2 gap-4 mt-3">
                  <div><p className="text-[10px] text-neutral-500 tracking-wider">AVERAGE SENTIMENT</p><p className={`text-xl font-bold font-mono ${profile.avgSentiment < -0.2 ? "text-red-400" : profile.avgSentiment > 0.2 ? "text-emerald-400" : "text-amber-400"}`}>{profile.avgSentiment.toFixed(2)}</p></div>
                  <div><p className="text-[10px] text-neutral-500 tracking-wider">TOPICS EXTRACTED</p><div className="flex flex-wrap gap-1 mt-1">{profile.commentTopics.length > 0 ? profile.commentTopics.map(t=><span key={t} className="text-[10px] px-2 py-0.5 rounded-full bg-cyan-900/30 text-cyan-300 border border-cyan-800/40">{t}</span>):<span className="text-xs text-neutral-600">None</span>}</div></div>
                </div>
              </div>
            </div>
          )}
          <div>
            <div className="flex items-center gap-2 mb-3"><MessageSquare className="w-4 h-4 text-purple-400" /><span className="text-xs text-neutral-300 tracking-wider font-bold">COMMENT HISTORY</span></div>
            <div className="space-y-1.5 max-h-56 overflow-y-auto pr-1">
              {commentHistory.map((v, i) => { const isLow = LOW_INFO_SET.has(v.comment.toLowerCase().trim()) || v.comment.length <= 3; return (<div key={i} className="bg-neutral-950 border border-neutral-800 rounded-lg px-3 py-2 text-xs"><div className="flex items-center gap-2 mb-1"><span className="text-neutral-500 font-mono">{v.date.slice(0, 10)}</span><span className="text-cyan-400/80">{v.shopName}</span></div><p className={`leading-relaxed ${isLow ? "text-amber-300/70 italic" : "text-neutral-300"}`}>"{v.comment}"</p></div>); })}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function VisitMapWidget({ visits, integrityByCode }: { visits: VisitRow[]; integrityByCode: Map<string, IntegrityProfile> }) {
  const [timeSlider, setTimeSlider] = useState(24);
  const [isClient, setIsClient] = useState(false);
  useEffect(() => { setIsClient(true); }, []);

  const filteredByTime = useMemo(() => {
    if (timeSlider === 24) return visits;
    return visits.filter(v => { const hr = getHour(v.date); return hr !== null && hr <= timeSlider; });
  }, [visits, timeSlider]);

  if (!isClient) return <div className="h-[400px] bg-neutral-900 animate-pulse rounded-xl" />;

  return (
    <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
      <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-700">
        <div className="flex items-center gap-2"><MapPin className="w-4 h-4 text-cyan-400" /><span className="text-xs text-neutral-300 tracking-wider font-semibold">SPATIAL INTELLIGENCE MAP</span></div>
        <div className="flex items-center gap-4">
          <span className="text-[10px] text-neutral-500">Route Playback: {timeSlider === 24 ? 'Full Day' : `Until ${timeSlider}:00`}</span>
          <input type="range" min="8" max="24" value={timeSlider} onChange={e => setTimeSlider(parseInt(e.target.value))} className="w-32 h-1 bg-neutral-700 rounded-lg appearance-none cursor-pointer accent-cyan-500" />
        </div>
      </div>
      <div className="relative h-[500px] w-full bg-neutral-950 z-0">
        <MapContainer center={[30.5, 31.5]} zoom={8} style={{ height: '100%', width: '100%' }} className="z-0">
          <TileLayer url="https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png" />
          {filteredByTime.map((v, idx) => {
            const ip = integrityByCode.get(v.spvrCode); const isGhost = ip && ip.ghostScore >= 50;
            return (
              <CircleMarker key={idx} center={[v.lat, v.lon]} radius={isGhost ? 8 : 4} 
                pathOptions={{ color: isGhost ? '#f87171' : '#22d3ee', weight: isGhost ? 2 : 1, fillColor: isGhost ? 'rgba(248,113,113,0.4)' : 'rgba(34,211,238,0.3)', fillOpacity: 0.8 }}>
                <Popup><div className="text-xs bg-neutral-900 p-2 rounded shadow-lg border border-neutral-700"><p className="font-bold text-white">{v.shopName}</p><p className="text-neutral-400">{v.area} | {v.date.slice(0,10)}</p><p className="text-neutral-500 mt-1 italic">{v.comment || "(No comment)"}</p></div></Popup>
              </CircleMarker>
            );
          })}
        </MapContainer>
        <div className="absolute bottom-4 left-4 z-[1000] bg-neutral-900/90 backdrop-blur-sm border border-neutral-700 rounded-lg p-3 shadow-xl">
          <p className="text-[10px] text-neutral-400 font-bold tracking-widest mb-2">LEGEND</p>
          <div className="flex items-center gap-2 text-[10px] mb-1"><span className="w-3 h-3 rounded-full bg-cyan-400/50 border border-cyan-400" /> Clean Visit</div>
          <div className="flex items-center gap-2 text-[10px]"><span className="w-3 h-3 rounded-full bg-red-400/50 border border-red-400" /> Ghost/High Risk</div>
        </div>
      </div>
    </div>
  );
}

function CommandBar({ isOpen, onClose, employees, onSelectEmployee, onExport }:{ isOpen:boolean, onClose:()=>void, employees:any[], onSelectEmployee:(code:string)=>void, onExport:()=>void }) {
  const [query, setQuery] = useState("");
  if (!isOpen) return null;
  const filtered = employees.filter(e => e.name.toLowerCase().includes(query.toLowerCase()) || e.code.toLowerCase().includes(query.toLowerCase()));
  return (
    <div className="fixed inset-0 z-[200] flex items-start justify-center pt-[20vh]" onClick={onClose}>
      <div className="absolute inset-0 bg-black/60 backdrop-blur-sm" />
      <div className="relative z-10 w-full max-w-lg bg-neutral-900 border border-neutral-600 rounded-2xl shadow-2xl overflow-hidden" onClick={e => e.stopPropagation()}>
        <div className="flex items-center gap-3 px-4 border-b border-neutral-700">
          <Search className="w-5 h-5 text-neutral-400" />
          <input autoFocus placeholder="Search employees or type 'export'..." value={query} onChange={e => setQuery(e.target.value)} className="w-full bg-transparent py-4 text-white text-sm outline-none placeholder-neutral-500" />
          <kbd className="text-[10px] text-neutral-500 border border-neutral-600 rounded px-1.5 py-0.5">ESC</kbd>
        </div>
        <div className="max-h-80 overflow-y-auto p-2">
          {query.toLowerCase() === 'export' && <button onClick={onExport} className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-purple-300 hover:bg-neutral-800"><Download className="w-4 h-4" /> Download Evidence Report (CSV)</button>}
          {filtered.map(e => (<button key={e.code} onClick={() => { onSelectEmployee(e.code); onClose(); }} className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-white hover:bg-neutral-800"><Search className="w-4 h-4 text-neutral-500" /><span>{e.name}</span><span className="text-neutral-500 font-mono text-xs ml-auto">{e.code}</span></button>))}
        </div>
      </div>
    </div>
  );
}

function exportEvidenceReport(enriched: any[], integrityProfiles: IntegrityProfile[], visits: VisitRow[], audit: any[]) {
  const timestamp = new Date().toISOString();
  const header = ["Code","Name","Department","Compliance","Integrity","GhostScore","Teleportations","RushHours","Mutations","Sentiment"];
  const rows = enriched.map(e => {
    const ip = integrityProfiles.find(p => p.code === e.code);
    return [e.code, e.name, e.department, e.compliance, ip?.integrityScore, ip?.ghostScore, ip?.teleportationFlags.length, ip?.rushHourFlags.length, ip?.mutatedCopyPastes, ip?.avgSentiment.toFixed(2)].join(",");
  });
  const csv = [`# Evidence Report ${timestamp}`, header.join(","), ...rows].join("\n");
  const blob = new Blob([csv], { type: "text/csv" }); const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = `evidence-${timestamp.slice(0,10)}.csv`; a.click(); URL.revokeObjectURL(url);
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

  useEffect(() => {
    const saved = localStorage.getItem('visitComplianceState');
    if (saved) { try { const state = JSON.parse(saved); if(state.employees) setEmployees(state.employees); if(state.visits) setVisits(state.visits); if(state.files) setFiles(state.files); } catch(e) {} }
  }, []);
  useEffect(() => { if (employees.length > 0 || visits.length > 0) { localStorage.setItem('visitComplianceState', JSON.stringify({ employees, visits, files })); } }, [employees, visits, files]);
  useEffect(() => { const handler = (e: KeyboardEvent) => { if ((e.metaKey || e.ctrlKey) && e.key === 'k') { e.preventDefault(); setCommandBarOpen(true); } }; window.addEventListener('keydown', handler); return () => window.removeEventListener('keydown', handler); }, []);

  const handleFiles=useCallback(async(fileList:FileList)=>{
    setIsProcessing(true);
    const newLoaded:LoadedFile[]=[];
    let newEmployees:EmployeeRecord[]|null=null;
    let newFollowUp:Map<string,{total:number;missing:string[]}>|null=null;
    let newRaw:Set<string>|null=null;
    let newVisits:VisitRow[]|null=null;
    for(const file of Array.from(fileList)){
      try{
        const buf=await file.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array"});
        const kind=detectKind(wb.SheetNames);
        newLoaded.push({name:file.name,kind,size:fmtBytes(file.size)});
        if(kind==="visit_summary"){
          const ws=wb.Sheets["Weekly Summary"];
          newEmployees=parseSummary(XLSX.utils.sheet_to_json(ws,{header:1,defval:null}) as unknown[][]);
        } else if(kind==="visit_followup"){
          const ws=wb.Sheets["Submission Matrix"];
          newFollowUp=parseFollowUp(XLSX.utils.sheet_to_json(ws,{header:1,defval:null}) as unknown[][]);
        } else if(kind==="visit_raw"){
          const ws=wb.Sheets["Mobile"];
          const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null}) as unknown[][];
          newVisits=parseRawVisits(rows);
          newRaw=new Set(newVisits.map(v=>v.spvrCode));
        }
      }catch{
        newLoaded.push({name:file.name,kind:"unknown",size:"—"});
      }
    }
    setFiles(prev=>{
      const ex=new Set(prev.map(f=>f.name));
      return [...prev,...newLoaded.filter(f=>!ex.has(f.name))];
    });
    if(newFollowUp) setFollowUpMap(newFollowUp);
    if(newRaw) setRawCodes(newRaw);
    if(newEmployees) setEmployees(newEmployees);
    if(newVisits) setVisits(newVisits);
    setIsProcessing(false);
  },[]);

  const enriched = useMemo(()=>employees.map(e=>({
    ...e,
    inFollowUp:followUpMap.has(e.code),
    inRaw:rawCodes.has(e.code),
    compliance:pct(e.daysReported,e.daysPossible),
  })),[employees,followUpMap,rawCodes]);

  const integrityProfiles=useMemo(()=>buildIntegrityProfiles(visits,employees),[visits,employees]);
  const integrityByCode=useMemo(()=>new Map(integrityProfiles.map(p=>[p.code,p])),[integrityProfiles]);

  const audit = useMemo(()=>{
    const issues:{severity:"high"|"medium";message:string}[]=[];
    const seen=new Set<string>();
    for(const e of enriched){
      if(seen.has(e.code)) issues.push({severity:"high",message:`Duplicate code "${e.code}" (${e.name}) in Summary.`});
      seen.add(e.code);
      if(!e.inRaw) issues.push({severity:"medium",message:`${e.name} (${e.code}) has no raw visit entries.`});
      if(!e.inFollowUp) issues.push({severity:"medium",message:`${e.name} (${e.code}) missing from Follow-Up Matrix.`});
      if(e.daysPossible>0&&e.daysReported===0) issues.push({severity:"high",message:`${e.name} (${e.code}) — zero days reported entire period.`});
      const ip=integrityByCode.get(e.code);
      if(ip){
        if(ip.unsupportedClaims>0) issues.push({severity:"high",message:`${e.name}: ${ip.unsupportedClaims} shortage/movement claim(s) with zero supporting comment.`});
        if(ip.topRepeatedComment&&ip.topRepeatedComment.count>=3) issues.push({severity:"medium",message:`${e.name}: reused "${ip.topRepeatedComment.text.slice(0,50)}" ${ip.topRepeatedComment.count}× — likely copy-paste.`});
        if(ip.lateNightCount>ip.totalVisits*0.6&&ip.totalVisits>=4) issues.push({severity:"high",message:`${e.name}: ${ip.lateNightCount}/${ip.totalVisits} reports filed 23:00–04:00 — possible batch backdating.`});
        if(ip.teleportationFlags.length>0) issues.push({severity:"high",message:`${e.name}: Impossible travel detected.`});
        if(ip.rushHourFlags.length>0) issues.push({severity:"high",message:`${e.name}: Rush-hour filing (>5 shops/hour).`});
      }
    }
    return issues;
  },[enriched,integrityByCode]);

  const departments=["All",...Array.from(new Set(enriched.map(e=>e.department)))];
  const filtered = useMemo(() => {
    let list = deptFilter==="All"?enriched:enriched.filter(e=>e.department===deptFilter);
    if (statusFilter === "AT RISK") list = list.filter(e => { const t = Math.round(e.compliance*0.5+(integrityByCode.get(e.code)?.integrityScore??100)*0.5); return t < 50; });
    if (statusFilter === "WATCH") list = list.filter(e => { const t = Math.round(e.compliance*0.5+(integrityByCode.get(e.code)?.integrityScore??100)*0.5); return t >= 50 && t < 80; });
    if (statusFilter === "HEALTHY") list = list.filter(e => { const t = Math.round(e.compliance*0.5+(integrityByCode.get(e.code)?.integrityScore??100)*0.5); return t >= 80; });
    return list;
  }, [deptFilter, enriched, integrityByCode, statusFilter]);

  const deptRollup=useMemo(()=>
    Array.from(new Set(enriched.map(e=>e.department))).map(dept=>{
      const members=enriched.filter(e=>e.department===dept);
      const compliance=pct(members.reduce((s,m)=>s+m.daysReported,0),members.reduce((s,m)=>s+m.daysPossible,0));
      const chronic=members.filter(m=>m.compliance<50).length;
      const avgIntegrity=members.length
        ?Math.round(members.reduce((s,m)=>s+(integrityByCode.get(m.code)?.integrityScore??100),0)/members.length):100;
      return {dept,headcount:members.length,compliance,chronic,avgIntegrity};
    }).sort((a,b)=>b.compliance-a.compliance),
  [enriched,integrityByCode]);

  const overallCompliance=pct(enriched.reduce((s,e)=>s+e.daysReported,0),enriched.reduce((s,e)=>s+e.daysPossible,0));

  const fieldTrust=useMemo(()=>enriched.map(e=>{
    const ip=integrityByCode.get(e.code);
    const integrity=ip?.integrityScore??100;
    const trust=Math.round(e.compliance*0.5+integrity*0.5);
    return {...e,integrity,trust};
  }).sort((a,b)=>a.trust-b.trust),[enriched,integrityByCode]);

  const chronicOffenders=fieldTrust.filter(e=>e.trust<50);

  const statusBuckets=useMemo(()=>[
    {name:"Healthy", value:fieldTrust.filter(e=>e.trust>=80).length,  color:"#34d399", key:"HEALTHY"},
    {name:"Watch",   value:fieldTrust.filter(e=>e.trust>=50&&e.trust<80).length, color:"#fbbf24", key:"WATCH"},
    {name:"At Risk", value:fieldTrust.filter(e=>e.trust<50).length,   color:"#f87171", key:"AT RISK"},
  ],[fieldTrust]);

  const dailyTrend=useMemo(()=>{
    const map=new Map<string,number>();
    for(const v of visits) map.set(v.date.slice(0,10),(map.get(v.date.slice(0,10))||0)+1);
    return Array.from(map.entries()).sort((a,b)=>a[0].localeCompare(b[0])).map(([date,count])=>({date:date.slice(5),count}));
  },[visits]);

  const avgVolume=dailyTrend.length>0?Math.round(dailyTrend.reduce((s,d)=>s+d.count,0)/dailyTrend.length):0;

  const ghostRanking=useMemo(()=>[...integrityProfiles].sort((a,b)=>b.ghostScore-a.ghostScore),[integrityProfiles]);

  const hasData=enriched.length>0;
  const hasVisits=visits.length>0;
  const avgIntegrityAll=Math.round(fieldTrust.reduce((s,e)=>s+e.integrity,0)/Math.max(1,fieldTrust.length));
  const totalStaff=enriched.length;

  const visitsByEmpCode = useMemo(() => { const m = new Map<string, VisitRow[]>(); for (const v of visits) { if (!m.has(v.spvrCode)) m.set(v.spvrCode, []); m.get(v.spvrCode)!.push(v); } return m; }, [visits]);
  const selectedEmployee = useMemo(() => { if (!selectedEmployeeCode) return null; const e = enriched.find(emp => emp.code === selectedEmployeeCode); if (!e) return null; return { employee: e, profile: integrityByCode.get(e.code), visits: visitsByEmpCode.get(e.code) ?? [] }; }, [selectedEmployeeCode, enriched, integrityByCode, visitsByEmpCode]);

  return (
    <div className="p-6 space-y-6 max-w-full">
      <CommandBar isOpen={commandBarOpen} onClose={() => setCommandBarOpen(false)} employees={enriched} onSelectEmployee={(code) => setSelectedEmployeeCode(code)} onExport={() => exportEvidenceReport(enriched, integrityProfiles, visits, audit)} />

      {/* ══ HEADER ══════════════════════════════════════════════════════════ */}
      <div className="flex items-start justify-between flex-wrap gap-3">
        <div>
          <div className="flex items-center gap-3 mb-1">
            <div className="w-1 h-8 rounded-full" style={{background:"linear-gradient(to bottom,#22d3ee,#a855f7,#f87171)"}}/>
            <h2 className="text-xl font-bold text-white tracking-widest">VISIT COMPLIANCE &amp; FIELD INTEGRITY</h2>
            <span className="text-[10px] font-mono text-cyan-500 bg-cyan-500/10 border border-cyan-500/20 px-2 py-0.5 rounded-full">ENHANCED v2.0</span>
          </div>
          <p className="text-xs text-neutral-500 ml-4">Submission tracking · data integrity · ghost-activity detection · <kbd className="px-1 border border-neutral-700 rounded bg-neutral-800 text-neutral-400">Cmd+K</kbd> command bar</p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          {hasData && (<button onClick={() => exportEvidenceReport(enriched, integrityProfiles, visits, audit)} className="flex items-center gap-1.5 text-xs font-medium text-purple-300 bg-purple-500/10 border border-purple-500/30 px-3 py-1.5 rounded-lg hover:bg-purple-500/20 transition-colors"><Download className="w-3.5 h-3.5"/>Export CSV</button>)}
          {files.length>0&&(
            <div className="flex items-center gap-2">
              <span className="text-xs text-neutral-500 bg-neutral-800 border border-neutral-700 px-3 py-1.5 rounded-lg">{files.length} file{files.length>1?"s":""} loaded</span>
              <button onClick={()=>{setFiles([]);setEmployees([]);setFollowUpMap(new Map());setRawCodes(new Set());setVisits([]);localStorage.removeItem('visitComplianceState');}}
                className="text-xs text-red-400 bg-red-500/10 border border-red-500/30 px-3 py-1.5 rounded-lg hover:bg-red-500/20 transition-colors">
                Clear all
              </button>
            </div>
          )}
        </div>
      </div>

      {/* ══ UPLOAD ZONE ═════════════════════════════════════════════════════ */}
      <div
        onDragOver={e=>{e.preventDefault();setIsDragging(true);}}
        onDragLeave={()=>setIsDragging(false)}
        onDrop={e=>{e.preventDefault();setIsDragging(false);if(e.dataTransfer.files.length)handleFiles(e.dataTransfer.files);}}
        className={`relative border-2 border-dashed rounded-xl p-8 text-center transition-all ${isDragging?"border-cyan-500 bg-cyan-500/5":"border-neutral-700 hover:border-neutral-600 bg-neutral-900"}`}
      >
        <input type="file" accept=".xlsx,.xls" multiple onChange={e=>{if(e.target.files?.length)handleFiles(e.target.files);}} className="absolute inset-0 opacity-0 cursor-pointer"/>
        <Upload className={`w-8 h-8 mx-auto mb-3 ${isDragging?"text-cyan-400":"text-neutral-500"}`}/>
        <p className="text-sm font-medium text-white">Drop Raw / Follow-Up / Summary .xlsx files here, or click to browse</p>
        <p className="text-xs text-neutral-500 mt-1">
          Auto-detected by sheet: <span className="text-cyan-400 font-mono">Mobile</span> · <span className="text-purple-400 font-mono">Submission Matrix</span> · <span className="text-emerald-400 font-mono">Weekly Summary</span>
        </p>
        {isProcessing&&<div className="mt-3 flex items-center justify-center gap-2 text-xs text-cyan-400"><RefreshCw className="w-3.5 h-3.5 animate-spin"/>Parsing…</div>}
      </div>

      {files.length>0&&(
        <div className="flex flex-wrap gap-2">
          {files.map(f=>(
            <div key={f.name} className={`flex items-center gap-2 px-4 py-2 rounded-xl border text-xs font-medium ${f.kind==="unknown"?"bg-neutral-900 border-amber-500/40 text-amber-300":"bg-neutral-900 border-neutral-700 text-neutral-300"}`}>
              <FileSpreadsheet className="w-3.5 h-3.5"/>
              <span className="max-w-[200px] truncate">{f.name}</span>
              <span className="text-neutral-600">{f.size}</span>
              {f.kind==="unknown"
                ?<span className="flex items-center gap-1 text-amber-400"><AlertTriangle className="w-3 h-3"/>unrecognised</span>
                :<span className="flex items-center gap-1 text-emerald-400"><CheckCircle className="w-3 h-3"/>{f.kind.replace("visit_","")}</span>}
            </div>
          ))}
        </div>
      )}

      {!hasData&&files.length===0&&(
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {[
            {icon:FileSpreadsheet,color:"cyan",   title:"Raw Visit Log",    desc:"Daily market visits, comments & claims (Mobile sheet)"},
            {icon:ShieldCheck,    color:"purple", title:"Follow-Up Matrix", desc:"Per-employee submission times, missing-date audit"},
            {icon:Layers,         color:"emerald",title:"Weekly Summary",   desc:"Employee × Department rollup, daily reported flag"},
          ].map((c,i)=>(
            <div key={i} className={`bg-neutral-900 border rounded-xl p-5 border-${c.color}-500/30`}>
              <c.icon className={`w-6 h-6 mb-2 text-${c.color}-400`}/>
              <p className="text-white font-semibold mb-1">{c.title}</p>
              <p className="text-xs text-neutral-500 leading-relaxed">{c.desc}</p>
            </div>
          ))}
        </div>
      )}

      {hasData&&(<>

        {/* ══ KPI ROW ══════════════════════════════════════════════════════ */}
        <div className="grid grid-cols-2 lg:grid-cols-5 gap-3">
          {[
            {label:"HEADCOUNT",         val:enriched.length,         sub:"",             color:"text-white",       border:"border-neutral-700",    icon:Users},
            {label:"COMPLIANCE",        val:`${overallCompliance}%`, sub:"",             color:compColor(overallCompliance), border:"border-emerald-500/20", icon:overallCompliance>=70?TrendingUp:TrendingDown},
            {label:"AVG FIELD INTEGRITY",val:`${avgIntegrityAll}%`, sub:"",             color:intgColor(avgIntegrityAll),  border:"border-cyan-500/20",   icon:Gauge},
            {label:"AT-RISK STAFF",     val:chronicOffenders.length, sub:"trust < 50",  color:"text-red-400",     border:"border-red-500/20",     icon:ShieldAlert},
            {label:"AUDIT FLAGS",       val:audit.length,            sub:"issues found", color:"text-amber-400",  border:"border-amber-500/20",   icon:Flag},
          ].map(k=>(
            <div key={k.label} className={`bg-neutral-900 border ${k.border} rounded-xl p-4`}>
              <div className="flex items-center justify-between mb-2">
                <span className="text-[10px] text-neutral-400 tracking-wider font-medium">{k.label}</span>
                <k.icon className={`w-4 h-4 ${k.color}`}/>
              </div>
              <p className={`text-2xl font-bold font-mono ${k.color}`}>{k.val}</p>
              {k.sub&&<p className="text-[10px] text-neutral-600 mt-0.5">{k.sub}</p>}
            </div>
          ))}
        </div>

        {/* ══ VISUAL ROW 1: Area Trend + Trust Donut ════════════════════════ */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">

          {/* ── GRADIENT AREA CHART ──────────────────────────────────────── */}
          <div className="lg:col-span-2 bg-neutral-900 border border-neutral-700 rounded-xl p-5">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2">
                <Activity className="w-4 h-4 text-cyan-400"/>
                <p className="text-xs text-neutral-300 tracking-wider font-semibold">DAILY VISIT VOLUME</p>
              </div>
              {hasVisits&&(
                <div className="flex items-center gap-2">
                  <span className="text-[10px] text-neutral-500">daily avg</span>
                  <span className="text-sm font-bold font-mono text-cyan-400">{avgVolume}</span>
                  <span className="text-[10px] text-neutral-600">visits</span>
                </div>
              )}
            </div>
            {hasVisits?(
              <ResponsiveContainer width="100%" height={200}>
                <AreaChart data={dailyTrend} margin={{top:8,right:8,bottom:0,left:0}}>
                  <defs><Defs/></defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="#1c1c1c" vertical={false}/>
                  <XAxis dataKey="date" tick={{fill:"#525252",fontSize:9}} axisLine={false} tickLine={false} interval="preserveStartEnd"/>
                  <YAxis tick={{fill:"#525252",fontSize:9}} axisLine={false} tickLine={false} width={26}/>
                  <Tooltip
                    contentStyle={{background:"#111",border:"1px solid #2a2a2a",borderRadius:10,color:"#fff",fontSize:11}}
                    labelStyle={{color:"#22d3ee",fontWeight:"bold"}}
                    formatter={(v:any)=>[`${v} visits`,"Volume"]}
                  />
                  <ReferenceLine y={avgVolume} stroke="#22d3ee" strokeDasharray="4 3" strokeOpacity={0.3}
                    label={{value:`avg ${avgVolume}`,position:"insideTopRight",fill:"#22d3ee",fontSize:9,fontFamily:"monospace"}}/>
                  <Area
                    type="monotone" dataKey="count"
                    stroke="#22d3ee" strokeWidth={2.5}
                    fill="url(#areaCyan)"
                    dot={<GlowDot/>}
                    activeDot={{r:6,fill:"#22d3ee",stroke:"#0e7490",strokeWidth:2,filter:"url(#glowCyan)"}}
                  />
                </AreaChart>
              </ResponsiveContainer>
            ):(
              <div className="h-[200px] flex flex-col items-center justify-center gap-2">
                <Activity className="w-8 h-8 text-neutral-700"/>
                <p className="text-xs text-neutral-600">Upload the Raw visit log to see the daily trend</p>
              </div>
            )}
          </div>

          {/* ── FIELD TRUST DONUT ────────────────────────────────────────── */}
          <div className="bg-neutral-900 border border-neutral-700 rounded-xl p-5 flex flex-col">
            <div className="flex items-center gap-2 mb-4">
              <Target className="w-4 h-4 text-purple-400"/>
              <p className="text-xs text-neutral-300 tracking-wider font-semibold">FIELD TRUST SPLIT (Click to Filter)</p>
            </div>
            <div className="flex-1 flex flex-col items-center justify-center relative cursor-pointer">
              <ResponsiveContainer width="100%" height={170}>
                <PieChart>
                  <defs><Defs/></defs>
                  <Pie
                    data={statusBuckets} dataKey="value" nameKey="name"
                    cx="50%" cy="50%"
                    innerRadius={44} outerRadius={68}
                    paddingAngle={5}
                    style={{filter:"url(#shadow)"}}
                    onClick={(data) => setStatusFilter(statusFilter === data.key ? null : data.key)}
                  >
                    {statusBuckets.map((s,i)=>(
                      <Cell key={i} fill={s.color}
                        style={{filter: statusFilter === s.key ? `drop-shadow(0 0 8px ${s.color})` : 'none', opacity: statusFilter && statusFilter !== s.key ? 0.3 : 1 }}/>
                    ))}
                  </Pie>
                  <Tooltip contentStyle={{background:"#111",border:"1px solid #2a2a2a",borderRadius:8,color:"#fff",fontSize:11}}/>
                </PieChart>
              </ResponsiveContainer>
              {/* centre overlay */}
              <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                <p className="text-2xl font-bold font-mono text-white">{totalStaff}</p>
                <p className="text-[10px] text-neutral-500">staff</p>
              </div>
            </div>
            {/* pill legend */}
            <div className="flex justify-center gap-2 mt-2 flex-wrap">
              {statusBuckets.map(s=>(
                <span key={s.name} onClick={() => setStatusFilter(statusFilter === s.key ? null : s.key)}
                  className={`flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[10px] font-semibold border cursor-pointer transition-all ${statusFilter === s.key ? 'bg-white/10' : ''}`}
                  style={{borderColor:`${s.color}44`,background:statusFilter === s.key ? `${s.color}30` : `${s.color}15`,color:s.color}}>
                  <span className="w-1.5 h-1.5 rounded-full" style={{background:s.color}}/>
                  {s.name} {s.value}
                </span>
              ))}
            </div>
          </div>
        </div>

        {/* ══ VISUAL ROW 2: Dept ranking — gradient bars + table ════════════ */}
        <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
          <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-700">
            <div className="flex items-center gap-2">
              <Crown className="w-4 h-4 text-amber-400"/>
              <span className="text-xs text-neutral-300 tracking-wider font-semibold">DEPARTMENT RANKING — COMPLIANCE %</span>
            </div>
            <span className="text-xs text-neutral-500">{deptRollup.length} team{deptRollup.length>1?"s":""}</span>
          </div>

          {/* Gradient horizontal bars */}
          <div className="px-5 pt-4 pb-2">
            <ResponsiveContainer width="100%" height={Math.max(80,deptRollup.length*56)}>
              <BarChart data={deptRollup} layout="vertical" margin={{left:0,right:40,top:4,bottom:4}}>
                <defs><Defs/></defs>
                <XAxis type="number" domain={[0,100]} tick={{fill:"#525252",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>`${v}%`}/>
                <YAxis type="category" dataKey="dept" tick={{fill:"#e5e5e5",fontSize:11,fontWeight:600}} axisLine={false} tickLine={false} width={96}/>
                <Tooltip
                  contentStyle={{background:"#111",border:"1px solid #2a2a2a",borderRadius:8,color:"#fff",fontSize:11}}
                  formatter={(v:any)=>[`${v}%`,"Compliance"]}
                />
                <Bar dataKey="compliance" barSize={26} shape={<GradBar/>} radius={[0,6,6,0]}>
                  <LabelList dataKey="compliance" position="right" formatter={(v:any)=>`${v}%`}
                    style={{fill:"#a3a3a3",fontSize:10,fontFamily:"monospace",fontWeight:"bold"}}/>
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* Summary table under chart */}
          <table className="w-full text-sm border-t border-neutral-800">
            <thead>
              <tr className="text-[10px] text-neutral-500 bg-neutral-950/60">
                {["#","DEPARTMENT","HEADCOUNT","COMPLIANCE","AVG INTEGRITY","CHRONIC"].map(h=>(
                  <th key={h} className="text-left py-2 px-4 font-semibold tracking-wider">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {deptRollup.map((d,i)=>(
                <tr key={d.dept} className="border-t border-neutral-800 hover:bg-neutral-800/40 transition-colors">
                  <td className="py-2.5 px-4 text-neutral-500 font-mono text-xs">#{i+1}</td>
                  <td className="py-2.5 px-4 text-white font-medium">{d.dept}</td>
                  <td className="py-2.5 px-4 font-mono text-neutral-300 text-xs">{d.headcount}</td>
                  <td className={`py-2.5 px-4 font-mono font-bold text-xs ${compColor(d.compliance)}`}>{d.compliance}%</td>
                  <td className={`py-2.5 px-4 font-mono font-bold text-xs ${intgColor(d.avgIntegrity)}`}>{d.avgIntegrity}%</td>
                  <td className="py-2.5 px-4 font-mono text-xs">
                    {d.chronic>0?<span className="text-red-400 font-bold">{d.chronic}</span>:<span className="text-neutral-600">—</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* ══ NEW: SPATIAL MAP WIDGET ═════════════════════════════════════ */}
        {hasVisits && <VisitMapWidget visits={visits} integrityByCode={integrityByCode} />}

        {/* ══ VISUAL ROW 3: Manager Action Desk — lowest trust + ghost ══════ */}
        {hasVisits&&(
          <div className="bg-neutral-900 border border-purple-500/20 rounded-xl overflow-hidden">
            {/* Tab header */}
            <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-700 flex-wrap gap-3">
              <div className="flex items-center gap-2">
                <Eye className="w-4 h-4 text-purple-400"/>
                <span className="text-xs text-purple-300 tracking-wider font-semibold">MANAGER ACTION DESK</span>
              </div>
              <div className="flex gap-1 p-1 bg-neutral-800 border border-neutral-700 rounded-lg">
                {([["trust","⬇ Lowest Trust"],["ghost","👻 Ghost Activity"]] as const).map(([id,label])=>(
                  <button key={id} onClick={()=>setLowestTab(id)}
                    className={`text-[10px] px-3 py-1.5 rounded font-bold tracking-wider transition-all ${
                      lowestTab===id
                        ?id==="ghost"?"bg-red-500/20 text-red-300 border border-red-500/30":"bg-purple-500/20 text-purple-300 border border-purple-500/30"
                        :"text-neutral-500 hover:text-neutral-300"
                    }`}>
                    {label}
                  </button>
                ))}
              </div>
            </div>

            {/* ─ Lowest Trust tab ─────────────────────────────────────────── */}
            {lowestTab==="trust"&&(
              <>
                <div className="px-5 py-2 border-b border-neutral-800 bg-neutral-950/40">
                  <p className="text-[10px] text-neutral-500">Sorted by Field Trust Score (compliance 50% + integrity 50%). Hover the <span className="text-amber-400 font-semibold">⚡ flags</span> badge for exact issues.</p>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-[10px] text-neutral-500 border-b border-neutral-800 bg-neutral-950/40">
                        {["EMPLOYEE","DEPT","VISITS","BLANK","LOW-INFO","COPY-PASTE","UNSUPPORTED","LATE NIGHT","INTEGRITY","TRUST SCORE","WHY FOLLOW UP"].map(h=>(
                          <th key={h} className="text-left py-2.5 px-3 font-semibold tracking-wider whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {fieldTrust.slice(0,15).map(e=>{
                        const ip=integrityByCode.get(e.code);
                        if(!ip) return null;
                        const badge=ghostBadge(ip.ghostScore);
                        return (
                          <tr key={e.code} onClick={() => setSelectedEmployeeCode(e.code)}
                            className={`border-b border-neutral-800 hover:bg-neutral-800/30 transition-colors cursor-pointer ${e.trust<40?"bg-red-950/10":""}`}>
                            <td className="py-2.5 px-3">
                              <div className="flex items-center gap-1.5">
                                <ChevronRight className="w-3 h-3 text-neutral-600" />
                                <div>
                                  <p className="text-white text-xs font-semibold leading-tight">{e.name}</p>
                                  <p className="text-cyan-400 font-mono text-[10px]">{e.code}</p>
                                </div>
                              </div>
                            </td>
                            <td className="py-2.5 px-3 text-neutral-400 text-xs">{e.department}</td>
                            <td className="py-2.5 px-3 font-mono text-neutral-300 text-xs text-center">{ip.totalVisits}</td>
                            <td className="py-2.5 px-3 text-center text-xs font-mono">
                              {ip.blankComments>0?<span className="text-red-400 font-bold">{ip.blankComments}</span>:<span className="text-neutral-600">—</span>}
                            </td>
                            <td className="py-2.5 px-3 text-center text-xs font-mono">
                              {ip.lowInfoComments>0?<span className="text-amber-400">{ip.lowInfoComments}</span>:<span className="text-neutral-600">—</span>}
                            </td>
                            <td className="py-2.5 px-3 text-center text-xs font-mono">
                              {ip.templatedRepeats>0?<span className="flex items-center justify-center gap-1 text-amber-400"><Copy className="w-3 h-3"/>{ip.templatedRepeats}</span>:<span className="text-neutral-600">—</span>}
                            </td>
                            <td className="py-2.5 px-3 text-center text-xs font-mono">
                              {ip.unsupportedClaims>0?<span className="flex items-center justify-center gap-1 text-red-400"><FileWarning className="w-3 h-3"/>{ip.unsupportedClaims}</span>:<span className="text-neutral-600">—</span>}
                            </td>
                            <td className="py-2.5 px-3 text-center text-xs font-mono">
                              {ip.lateNightCount>0?<span className="flex items-center justify-center gap-1 text-violet-400"><Clock className="w-3 h-3"/>{ip.lateNightCount}</span>:<span className="text-neutral-600">—</span>}
                            </td>
                            <td className={`py-2.5 px-3 text-center font-mono font-bold text-xs ${intgColor(ip.integrityScore)}`}>{ip.integrityScore}%</td>
                            <td className="py-2.5 px-3 text-center">
                              <span className={`text-xs font-bold font-mono px-2 py-0.5 rounded-full border ${
                                e.trust<50?"bg-red-500/15 text-red-300 border-red-500/30":
                                e.trust<80?"bg-amber-500/15 text-amber-300 border-amber-500/30":
                                "bg-emerald-500/15 text-emerald-300 border-emerald-500/30"
                              }`}>{e.trust}%</span>
                            </td>
                            <td className="py-2.5 px-3 min-w-[180px]">
                              <WhyCell e={e} ip={ip}/>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </>
            )}

            {/* ─ Ghost Activity tab ────────────────────────────────────────── */}
            {lowestTab==="ghost"&&(
              <>
                <div className="px-5 py-2 border-b border-neutral-800 bg-red-950/10">
                  <p className="text-[10px] text-neutral-500">
                    Ghost Score = blank comments (30%) + copy-paste (20%) + late-night batch filing (25%) + unsupported claims (15%) + single-shop loop (10%).
                    <span className="text-red-400 ml-1">≥60 = High Risk · ≥30 = Suspicious</span>
                  </p>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-[10px] text-neutral-500 border-b border-neutral-800 bg-neutral-950/40">
                        {["EMPLOYEE","VISITS","SHOPS","AVG COMMENT","LATE NIGHT","BLANK %","GHOST SCORE","RISK","SUSPICION FLAGS"].map(h=>(
                          <th key={h} className="text-left py-2.5 px-3 font-semibold tracking-wider whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {ghostRanking.slice(0,15).map(ip=>{
                        const badge=ghostBadge(ip.ghostScore);
                        return (
                          <tr key={ip.code} onClick={() => setSelectedEmployeeCode(ip.code)}
                            className={`border-b border-neutral-800 hover:bg-neutral-800/30 transition-colors cursor-pointer ${ip.ghostScore>=60?"bg-red-950/10":""}`}>
                            <td className="py-2.5 px-3">
                               <div className="flex items-center gap-1.5">
                                <ChevronRight className="w-3 h-3 text-neutral-600" />
                                <div>
                                  <p className="text-white text-xs font-semibold leading-tight">{ip.name}</p>
                                  <p className="text-cyan-400 font-mono text-[10px]">{ip.code}</p>
                                </div>
                              </div>
                            </td>
                            <td className="py-2.5 px-3 font-mono text-neutral-300 text-xs text-center">{ip.totalVisits}</td>
                            <td className="py-2.5 px-3 font-mono text-neutral-300 text-xs text-center">
                              {ip.singleShopLoop?<span className="text-red-400 font-bold">{ip.uniqueShops}</span>:ip.uniqueShops}
                            </td>
                            <td className="py-2.5 px-3 text-center font-mono text-xs">
                              <span className={ip.avgCommentLen<25?"text-red-400":ip.avgCommentLen<60?"text-amber-400":"text-emerald-400"}>
                                {ip.avgCommentLen}c
                              </span>
                            </td>
                            <td className="py-2.5 px-3 text-center font-mono text-xs">
                              {ip.lateNightCount>0
                                ?<span className="flex items-center justify-center gap-1 text-violet-400"><Clock className="w-3 h-3"/>{ip.lateNightCount}</span>
                                :<span className="text-neutral-600">—</span>}
                            </td>
                            <td className="py-2.5 px-3 text-center font-mono text-xs">
                              <span className={ip.blankComments>0?"text-red-400":"text-neutral-600"}>
                                {pct(ip.blankComments,ip.totalVisits)}%
                              </span>
                            </td>
                            <td className="py-2.5 px-3">
                              <div className="flex flex-col gap-1">
                                <GhostMeter score={ip.ghostScore}/>
                                <p className="text-[10px] font-mono text-neutral-400">{ip.ghostScore}/100</p>
                              </div>
                            </td>
                            <td className="py-2.5 px-3">
                              <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${badge.bg} ${badge.text} ${badge.border}`}>
                                {badge.label}
                              </span>
                            </td>
                            <td className="py-2.5 px-3">
                              <FlagTooltip profile={ip}/>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        )}

        {/* ══ AUDIT ISSUES ════════════════════════════════════════════════ */}
        {audit.length>0&&(
          <div className="bg-neutral-900 border border-amber-500/25 rounded-xl p-5 space-y-2">
            <div className="flex items-center gap-2 mb-2">
              <AlertTriangle className="w-4 h-4 text-amber-400"/>
              <span className="text-xs text-amber-300 tracking-wider font-bold">DATA AUDIT — {audit.length} ISSUE{audit.length>1?"S":""}</span>
            </div>
            <div className="space-y-1.5 max-h-56 overflow-y-auto pr-1">
              {audit.map((a,i)=>(
                <div key={i} className={`text-xs flex items-start gap-2 ${a.severity==="high"?"text-red-300":"text-neutral-400"}`}>
                  <span className={`mt-0.5 w-1.5 h-1.5 rounded-full shrink-0 ${a.severity==="high"?"bg-red-400":"bg-amber-400"}`}/>
                  {a.message}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ══ EMPLOYEE DETAIL TABLE ════════════════════════════════════════ */}
        <div className="bg-neutral-900 border border-neutral-700 rounded-xl overflow-hidden">
          <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-700 flex-wrap gap-2">
            <div className="flex items-center gap-3">
              <span className="text-xs text-neutral-300 tracking-wider font-semibold">EMPLOYEE-LEVEL DETAIL</span>
              {statusFilter && <span className="text-[10px] text-cyan-400 bg-cyan-900/30 px-2 py-0.5 rounded-full cursor-pointer" onClick={() => setStatusFilter(null)}>Filtered: {statusFilter} ✕</span>}
            </div>
            <select value={deptFilter} onChange={e=>setDeptFilter(e.target.value)}
              className="bg-neutral-800 border border-neutral-600 text-xs text-neutral-300 rounded px-3 py-1.5 focus:outline-none focus:border-cyan-500">
              {departments.map(d=><option key={d}>{d}</option>)}
            </select>
          </div>
          <div className="overflow-x-auto max-h-[480px] overflow-y-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-[10px] text-neutral-500 border-b border-neutral-800 sticky top-0 bg-neutral-900 z-10">
                  {["EMPLOYEE","CODE","DEPARTMENT","DAYS","COMPLIANCE","INTEGRITY","FOLLOW-UP","RAW LOG","TREND","STATUS"].map(h=>(
                    <th key={h} className="text-left py-2.5 px-3 font-semibold tracking-wider whitespace-nowrap">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map(e=>{
                  const ip=integrityByCode.get(e.code);
                  const trust=Math.round(e.compliance*0.5+(ip?.integrityScore??100)*0.5);
                  return (
                    <tr key={e.code} onClick={() => setSelectedEmployeeCode(e.code)}
                      className="border-b border-neutral-800 hover:bg-neutral-800/30 transition-colors cursor-pointer group">
                      <td className="py-2 px-3 text-white text-xs font-medium">{e.name}</td>
                      <td className="py-2 px-3 font-mono text-cyan-400 text-xs">{e.code}</td>
                      <td className="py-2 px-3 text-neutral-400 text-xs">{e.department}</td>
                      <td className="py-2 px-3 text-center font-mono text-neutral-300 text-xs">{e.daysReported}/{e.daysPossible}</td>
                      <td className={`py-2 px-3 text-center font-mono font-bold text-xs ${compColor(e.compliance)}`}>{e.compliance}%</td>
                      <td className={`py-2 px-3 text-center font-mono font-bold text-xs ${intgColor(ip?.integrityScore??100)}`}>{ip?.integrityScore??100}%</td>
                      <td className="py-2 px-3 text-center">
                        {e.inFollowUp?<CheckCircle className="w-4 h-4 text-emerald-400 inline"/>:<XCircle className="w-4 h-4 text-red-400 inline"/>}
                      </td>
                      <td className="py-2 px-3 text-center">
                        {e.inRaw?<CheckCircle className="w-4 h-4 text-emerald-400 inline"/>:<XCircle className="w-4 h-4 text-red-400 inline"/>}
                      </td>
                      <td className="py-2 px-3">
                        {ip && ip.dailyTrend.length > 1 && <SparklineChart data={ip.dailyTrend} color={trust < 50 ? "#f87171" : "#34d399"} />}
                      </td>
                      <td className="py-2 px-3">
                        {trust<50
                          ?<span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-red-500/15 text-red-300 border border-red-500/30">AT RISK</span>
                          :trust<80
                            ?<span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-amber-500/15 text-amber-300 border border-amber-500/30">WATCH</span>
                            :<span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-500/15 text-emerald-300 border border-emerald-500/30">HEALTHY</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

      </>)}

      {/* ══ MODALS ══════════════════════════════════════════════════════════ */}
      {selectedEmployee && (
        <EmployeeScorecardModal 
          employee={selectedEmployee.employee} 
          profile={selectedEmployee.profile} 
          visits={selectedEmployee.visits} 
          onClose={() => setSelectedEmployeeCode(null)}
        />
      )}
    </div>
  );
}
