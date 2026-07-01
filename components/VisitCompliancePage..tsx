'use client'
import React, { useState, useEffect, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, PieChart, Pie, Cell, LineChart, Line
} from 'recharts'
import {
  Upload, Users, ShieldAlert, TrendingUp, CheckCircle, AlertTriangle,
  Eye, Download, X, Activity, MapPin, Flag, Gauge, FileText,
  Building2, Calendar, Search, Award, Clock, AlertCircle, ChevronDown
} from 'lucide-react'

/* ═══════════════════════ TYPES ═══════════════════════ */
interface EmployeeRow {
  name: string; code: string; dept: string
  totalReports: number; uniqueStores: number
  firstDate: string; lastDate: string
  pending: number; reviewed: number; other: number
}
interface MobileRow {
  date: string; code: string; repName: string
  shopCode: string; shopName: string; area: string; gov: string
  samsungShortage: string; compShortage: string
  sellout: string; compBrand: string; compMovement: string
  comment: string; action: string; feedback: string
}
interface StoreRow {
  shopCode: string; shopName: string; area: string; gov: string
  totalVisits: number; uniqueSPVRs: number
  firstVisit: string; lastVisit: string
}
interface SubmissionRow {
  name: string; code: string
  totalReports: number; submitted: number; total: number
  dailyMap: Record<string, string>; dateKeys: string[]
}
interface WeeklyRow {
  name: string; code: string; dept: string
  w1: number; w2: number; w3: number; w4: number; w5: number
}
interface AppState {
  employees: EmployeeRow[]; mobile: MobileRow[]
  stores: StoreRow[]; submissions: SubmissionRow[]; weekly: WeeklyRow[]
  files: string[]
}

/* ═══════════════════════ PARSERS ═══════════════════════ */
function parseEmployee(ws: XLSX.WorkSheet): EmployeeRow[] {
  const rows = XLSX.utils.sheet_to_json<Record<string,unknown>>(ws, { defval: '' })
  return rows.map(r => ({
    name: String(r['SPVR Name'] || ''),
    code: String(r['SPVR Code'] || ''),
    dept: String(r['Department'] || ''),
    totalReports: Number(r['Total Reports']) || 0,
    uniqueStores: Number(r['Unique Stores']) || 0,
    firstDate: String(r['First Report Date'] || ''),
    lastDate: String(r['Last Report Date'] || ''),
    pending: Number(r['Pending']) || 0,
    reviewed: Number(r['Reviewed']) || 0,
    other: Number(r['Other']) || 0,
  })).filter(r => r.name.trim())
}

function parseMobile(ws: XLSX.WorkSheet): MobileRow[] {
  const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' }) as unknown[][]
  return raw.slice(2)
    .filter(r => r[0] && String(r[0]).trim())
    .map(r => ({
      date: String(r[0] || ''), code: String(r[1] || ''), repName: String(r[2] || ''),
      shopCode: String(r[3] || ''), shopName: String(r[4] || ''),
      area: String(r[5] || ''), gov: String(r[6] || ''),
      samsungShortage: String(r[7] || ''), compShortage: String(r[8] || ''),
      sellout: String(r[9] || ''), compBrand: String(r[10] || ''),
      compMovement: String(r[11] || ''), comment: String(r[12] || ''),
      action: String(r[13] || ''), feedback: String(r[14] || ''),
    }))
}

function parseStores(ws: XLSX.WorkSheet): StoreRow[] {
  const rows = XLSX.utils.sheet_to_json<Record<string,unknown>>(ws, { defval: '' })
  return rows.map(r => ({
    shopCode: String(r['Shop Code'] || ''), shopName: String(r['Shop Name'] || ''),
    area: String(r['Area'] || ''), gov: String(r['Governorate'] || ''),
    totalVisits: Number(r['Total Visits']) || 0, uniqueSPVRs: Number(r['Unique SPVRs']) || 0,
    firstVisit: String(r['First Visit'] || ''), lastVisit: String(r['Last Visit'] || ''),
  })).filter(r => r.shopCode.trim())
}

function parseSubmissions(ws: XLSX.WorkSheet): SubmissionRow[] {
  const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' }) as unknown[][]
  if (raw.length < 2) return []
  const headers = (raw[0] as string[])
  const dateKeys = headers.slice(4)
  return raw.slice(1)
    .filter(r => r[0] && String(r[0]).trim())
    .map(r => {
      const dailyMap: Record<string, string> = {}
      let submitted = 0
      dateKeys.forEach((dk, i) => {
        const val = String(r[4 + i] || '').trim()
        dailyMap[dk] = val
        if (val && val !== '–') submitted++
      })
      return {
        name: String(r[0]), code: String(r[1]),
        totalReports: Number(r[2]) || 0,
        submitted, total: dateKeys.length, dailyMap, dateKeys,
      }
    })
}

function parseWeekly(ws: XLSX.WorkSheet): WeeklyRow[] {
  const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' }) as unknown[][]
  if (raw.length < 3) return []
  return raw.slice(2)
    .filter(r => r[0] && String(r[0]).trim())
    .map(r => ({
      name: String(r[0]), code: String(r[1]), dept: String(r[2]),
      w1: Number(r[7]) || 0,  w2: Number(r[16]) || 0,
      w3: Number(r[25]) || 0, w4: Number(r[34]) || 0, w5: Number(r[41]) || 0,
    }))
}

/* ═══════════════════════ HELPERS ═══════════════════════ */
const PIE_COLORS = ['#06b6d4','#8b5cf6','#10b981','#f59e0b','#ef4444','#ec4899','#3b82f6','#84cc16','#f97316','#a78bfa']

function compliancePct(row: SubmissionRow) {
  if (!row.total) return 0
  return Math.round((row.submitted / row.total) * 100)
}
function grade(pct: number) {
  if (pct >= 90) return { label: 'EXCELLENT', color: '#10b981', bg: '#10b98122' }
  if (pct >= 75) return { label: 'GOOD',      color: '#06b6d4', bg: '#06b6d422' }
  if (pct >= 55) return { label: 'AVERAGE',   color: '#f59e0b', bg: '#f59e0b22' }
  if (pct >= 35) return { label: 'WEAK',      color: '#f97316', bg: '#f9731622' }
  return                 { label: 'CRITICAL',  color: '#ef4444', bg: '#ef444422' }
}
function shortageFlag(v: string) {
  return v && v.trim() && v !== 'N' && v !== 'No' && v.length > 1
}

const FILE_SLOTS = [
  { label: 'Employee Summary',  hint: 'Employee_Summary.xlsx' },
  { label: 'Mobile Reports',    hint: 'Mobile_Dep.xlsx' },
  { label: 'Store Coverage',    hint: 'Store_Coverage.xlsx' },
  { label: 'Submission Matrix', hint: 'Submission_Matrix.xlsx' },
  { label: 'Weekly Summary',    hint: 'Weekly_Summary.xlsx' },
] as const

const TABS = ['LEADERBOARD','SUBMISSION MATRIX','WEEKLY TRENDS','COVERAGE','FIELD REPORTS','INTEL'] as const

/* ═══════════════════════ COMPONENT ═══════════════════════ */
export default function VisitCompliancePage() {
  const [state, setState] = useState<AppState>({
    employees: [], mobile: [], stores: [], submissions: [], weekly: [], files: []
  })
  const [tab, setTab] = useState(0)
  const [search, setSearch] = useState('')
  const [govFilter, setGovFilter] = useState('ALL')
  const [sortField, setSortField] = useState<'totalReports'|'uniqueStores'|'compliancePct'>('totalReports')

  // Restore from localStorage
  useEffect(() => {
    try {
      const saved = localStorage.getItem('auditDash_v5')
      if (saved) { try { setState(JSON.parse(saved)) } catch {} }
    } catch {}
  }, [])

  // Save to localStorage with try/catch (safe for large data)
  useEffect(() => {
    if (state.files.length > 0) {
      try { localStorage.setItem('auditDash_v5', JSON.stringify(state)) }
      catch { console.warn('localStorage quota exceeded — state not saved') }
    }
  }, [state])

  const handleFile = useCallback((file: File, slot: number) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target?.result, { type: 'binary', cellDates: true })
        const ws = wb.Sheets[wb.SheetNames[0]]
        setState(prev => {
          const newFiles = [...prev.files]
          newFiles[slot] = file.name
          switch (slot) {
            case 0: return { ...prev, files: newFiles, employees: parseEmployee(ws) }
            case 1: return { ...prev, files: newFiles, mobile:    parseMobile(ws) }
            case 2: return { ...prev, files: newFiles, stores:    parseStores(ws) }
            case 3: return { ...prev, files: newFiles, submissions: parseSubmissions(ws) }
            case 4: return { ...prev, files: newFiles, weekly:    parseWeekly(ws) }
            default: return prev
          }
        })
      } catch (err) { console.error('Parse error:', err) }
    }
    reader.readAsBinaryString(file)
  }, [])

  /* ── Derived data ── */
  const subMap = useMemo(() => {
    const m: Record<string, SubmissionRow> = {}
    state.submissions.forEach(s => { m[s.code] = s })
    return m
  }, [state.submissions])

  const leaderboard = useMemo(() => {
    return state.employees
      .map(e => {
        const sub = subMap[e.code]
        const pct = sub ? compliancePct(sub) : 0
        return { ...e, compliancePct: pct, grade: grade(pct) }
      })
      .sort((a, b) => (b[sortField] as number) - (a[sortField] as number))
  }, [state.employees, subMap, sortField])

  const totalReports    = useMemo(() => state.employees.reduce((s,e)=>s+e.totalReports,0), [state.employees])
  const totalStores     = useMemo(() => state.stores.length, [state.stores])
  const activeGovs      = useMemo(() => new Set(state.stores.map(s=>s.gov).filter(Boolean)).size, [state.stores])
  const overallCompliance = useMemo(() => {
    if (!state.submissions.length) return 0
    return Math.round(state.submissions.reduce((s,r)=>s+compliancePct(r),0)/state.submissions.length)
  }, [state.submissions])
  const shortageAlerts  = useMemo(() => state.mobile.filter(m=>shortageFlag(m.samsungShortage)).length, [state.mobile])

  const govBarData = useMemo(() => {
    const m: Record<string,number> = {}
    state.stores.forEach(s => { if(s.gov) m[s.gov]=(m[s.gov]||0)+s.totalVisits })
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,12).map(([name,value])=>({name,value}))
  }, [state.stores])

  const weeklyTotalChart = useMemo(() => {
    return [
      { week:'W1 Jun02-05', visits: state.weekly.reduce((s,r)=>s+r.w1,0) },
      { week:'W2 Jun06-12', visits: state.weekly.reduce((s,r)=>s+r.w2,0) },
      { week:'W3 Jun13-19', visits: state.weekly.reduce((s,r)=>s+r.w3,0) },
      { week:'W4 Jun20-26', visits: state.weekly.reduce((s,r)=>s+r.w4,0) },
      { week:'W5 Jun27-Jul1',visits: state.weekly.reduce((s,r)=>s+r.w5,0) },
    ]
  }, [state.weekly])

  const competitorData = useMemo(() => {
    const m: Record<string,number> = {}
    state.mobile.forEach(r => { const b=r.compBrand?.trim(); if(b) m[b]=(m[b]||0)+1 })
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,9).map(([name,value])=>({name,value}))
  }, [state.mobile])

  const allGovs = useMemo(() => ['ALL',...new Set(state.mobile.map(m=>m.gov).filter(Boolean))], [state.mobile])

  const filteredMobile = useMemo(() => {
    const s = search.toLowerCase()
    return state.mobile.filter(m => {
      const matchGov = govFilter === 'ALL' || m.gov === govFilter
      const matchSearch = !s ||
        m.repName.toLowerCase().includes(s) ||
        m.shopName.toLowerCase().includes(s) ||
        m.gov.toLowerCase().includes(s) ||
        m.sellout.toLowerCase().includes(s) ||
        m.samsungShortage.toLowerCase().includes(s)
      return matchGov && matchSearch
    }).slice(0, 120)
  }, [state.mobile, search, govFilter])

  const shortageList = useMemo(() =>
    state.mobile.filter(m => shortageFlag(m.samsungShortage)), [state.mobile])

  const topShortageModels = useMemo(() => {
    const m: Record<string,number> = {}
    shortageList.forEach(r => {
      const model = r.samsungShortage.trim()
      m[model] = (m[model]||0)+1
    })
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,10)
  }, [shortageList])

  const isLoaded = state.employees.length > 0 || state.mobile.length > 0

  const complianceColor = overallCompliance >= 70 ? '#10b981' : overallCompliance >= 50 ? '#f59e0b' : '#ef4444'
  const dateKeys = state.submissions[0]?.dateKeys || []

  /* ═══════════════════════ RENDER ═══════════════════════ */
  return (
    <div className="min-h-screen text-white" style={{ background: '#030712', fontFamily: "'JetBrains Mono', 'Fira Code', monospace" }}>

      {/* ── HEADER ── */}
      <header style={{ background: 'rgba(6,182,212,0.04)', borderBottom: '1px solid rgba(6,182,212,0.15)' }}
        className="px-6 py-3 flex items-center justify-between sticky top-0 z-40 backdrop-blur">
        <div className="flex items-center gap-3">
          <div style={{ background: 'rgba(6,182,212,0.1)', border: '1px solid rgba(6,182,212,0.3)', borderRadius: 8, padding: '6px 8px' }}>
            <ShieldAlert size={20} style={{ color: '#06b6d4' }} />
          </div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 700, letterSpacing: '0.15em', color: '#06b6d4' }}>
              FIELD AUDIT COMMAND · MOBILE DEPT
            </div>
            <div style={{ fontSize: 10, color: '#4b5563', letterSpacing: '0.08em' }}>
              EGYPT COVERAGE INTELLIGENCE · DATA NEVER LIES
            </div>
          </div>
        </div>
        <div className="flex items-center gap-2">
          {isLoaded && (
            <button
              onClick={() => { setState({employees:[],mobile:[],stores:[],submissions:[],weekly:[],files:[]}); try{localStorage.removeItem('auditDash_v5')}catch{} }}
              style={{ fontSize: 11, color: '#6b7280', border: '1px solid #374151', borderRadius: 6, padding: '4px 10px', cursor: 'pointer', background: 'transparent' }}
              className="hover:text-red-400 transition">
              <X size={11} className="inline mr-1"/>CLEAR DATA
            </button>
          )}
          <div style={{ fontSize: 10, color: '#374151', border: '1px solid #1f2937', borderRadius: 6, padding: '4px 10px' }}>
            {new Date().toLocaleDateString('en-EG',{day:'2-digit',month:'short',year:'numeric'})}
          </div>
        </div>
      </header>

      {/* ── FILE UPLOAD PANEL ── */}
      <div style={{ background: 'rgba(17,24,39,0.6)', borderBottom: '1px solid #1f2937', padding: '12px 24px' }}>
        <div className="grid grid-cols-5 gap-3">
          {FILE_SLOTS.map((slot, i) => (
            <label key={i} style={{
              border: `2px dashed ${state.files[i] ? '#06b6d4' : '#374151'}`,
              borderRadius: 8, padding: '10px 8px', cursor: 'pointer',
              background: state.files[i] ? 'rgba(6,182,212,0.06)' : 'rgba(17,24,39,0.4)',
              display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 5,
              transition: 'border-color 0.2s',
            }}>
              <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
                onChange={e => e.target.files?.[0] && handleFile(e.target.files[0], i)} />
              <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: '0.1em', color: state.files[i] ? '#06b6d4' : '#6b7280' }}>
                FILE {i+1}
              </span>
              {state.files[i]
                ? <CheckCircle size={15} style={{ color: '#06b6d4' }} />
                : <Upload size={15} style={{ color: '#6b7280' }} />}
              <span style={{ fontSize: 10, color: state.files[i] ? '#22d3ee' : '#4b5563', textAlign: 'center', lineHeight: 1.3 }}>
                {state.files[i] ? state.files[i].replace(/\.(xlsx|xls)$/,'') : slot.label}
              </span>
            </label>
          ))}
        </div>
        {!isLoaded && (
          <div style={{ fontSize: 10, color: '#4b5563', textAlign: 'center', marginTop: 8, letterSpacing: '0.1em' }}>
            UPLOAD ALL 5 FILES TO ACTIVATE INTELLIGENCE DASHBOARD
          </div>
        )}
      </div>

      {!isLoaded ? (
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '60vh', gap: 16 }}>
          <ShieldAlert size={72} style={{ color: '#1f2937' }} />
          <p style={{ color: '#374151', fontSize: 13, letterSpacing: '0.1em' }}>AWAITING DATA UPLOAD — ALL 5 FILES REQUIRED</p>
        </div>
      ) : (
        <>
          {/* ── KPI ROW ── */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6,1fr)', gap: 12, padding: '16px 24px', borderBottom: '1px solid #111827' }}>
            {[
              { icon: Users,         label: 'FIELD REPS',      value: state.employees.length,      color: '#06b6d4' },
              { icon: FileText,      label: 'TOTAL REPORTS',   value: totalReports.toLocaleString(),color: '#8b5cf6' },
              { icon: Building2,     label: 'STORES COVERED',  value: totalStores.toLocaleString(), color: '#10b981' },
              { icon: Gauge,         label: 'COMPLIANCE RATE', value: `${overallCompliance}%`,      color: complianceColor },
              { icon: MapPin,        label: 'GOVERNORATES',    value: activeGovs,                   color: '#3b82f6' },
              { icon: AlertTriangle, label: 'SHORTAGE ALERTS', value: shortageAlerts,               color: '#f97316' },
            ].map(({ icon: Icon, label, value, color }) => (
              <div key={label} style={{ background: `${color}0d`, border: `1px solid ${color}22`, borderRadius: 10, padding: '10px 14px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 6 }}>
                  <Icon size={13} style={{ color }} />
                  <span style={{ fontSize: 9, color: '#6b7280', letterSpacing: '0.12em' }}>{label}</span>
                </div>
                <div style={{ fontSize: 26, fontWeight: 800, color, lineHeight: 1 }}>{value}</div>
              </div>
            ))}
          </div>

          {/* ── TABS ── */}
          <div style={{ display: 'flex', borderBottom: '1px solid #1f2937', padding: '0 24px', overflowX: 'auto' }}>
            {TABS.map((t, i) => (
              <button key={t} onClick={() => setTab(i)} style={{
                padding: '10px 16px', fontSize: 10, fontWeight: 700, letterSpacing: '0.12em',
                color: tab === i ? '#06b6d4' : '#4b5563',
                borderBottom: `2px solid ${tab === i ? '#06b6d4' : 'transparent'}`,
                background: 'transparent', cursor: 'pointer', transition: 'color 0.2s', whiteSpace: 'nowrap',
              }}>{t}</button>
            ))}
          </div>

          {/* ── CONTENT ── */}
          <div style={{ padding: '20px 24px' }}>

            {/* ─ TAB 0: LEADERBOARD ─ */}
            {tab === 0 && (
              <div>
                <div style={{ display: 'flex', gap: 8, marginBottom: 14, alignItems: 'center' }}>
                  <span style={{ fontSize: 10, color: '#4b5563', letterSpacing: '0.1em' }}>SORT BY:</span>
                  {(['totalReports','uniqueStores','compliancePct'] as const).map(f => (
                    <button key={f} onClick={() => setSortField(f)} style={{
                      fontSize: 10, padding: '3px 10px', borderRadius: 5, cursor: 'pointer', letterSpacing: '0.08em',
                      background: sortField===f ? 'rgba(6,182,212,0.15)' : 'transparent',
                      border: `1px solid ${sortField===f ? '#06b6d4' : '#374151'}`,
                      color: sortField===f ? '#06b6d4' : '#6b7280',
                    }}>
                      {f === 'totalReports' ? 'REPORTS' : f === 'uniqueStores' ? 'STORES' : 'COMPLIANCE'}
                    </button>
                  ))}
                  <span style={{ fontSize: 10, color: '#374151', marginLeft: 'auto' }}>{leaderboard.length} REPS</span>
                </div>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ borderBottom: '1px solid #1f2937' }}>
                        {['#','FIELD REP','CODE','REPORTS','STORES','COMPLIANCE %','LAST ACTIVE','GRADE'].map(h => (
                          <th key={h} style={{ textAlign: h==='#'||h==='GRADE'?'center':'left', padding:'6px 10px 6px 0', fontSize:9, color:'#4b5563', letterSpacing:'0.1em', fontWeight:600 }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {leaderboard.map((emp, idx) => (
                        <tr key={emp.code} style={{ borderBottom: '1px solid rgba(31,41,55,0.5)', transition: 'background 0.15s' }}
                          onMouseEnter={e => (e.currentTarget.style.background='rgba(6,182,212,0.04)')}
                          onMouseLeave={e => (e.currentTarget.style.background='transparent')}>
                          <td style={{ textAlign:'center', padding:'8px 10px 8px 0', fontWeight:700, color: idx===0?'#FFD700':idx===1?'#C0C0C0':idx===2?'#CD7F32':'#374151', fontSize:14 }}>
                            {idx<3 ? ['🥇','🥈','🥉'][idx] : idx+1}
                          </td>
                          <td style={{ padding:'8px 10px 8px 0', color:'#e5e7eb' }}>
                            {emp.name.split(' ').slice(0,4).join(' ')}
                          </td>
                          <td style={{ padding:'8px 10px 8px 0', color:'#06b6d4', fontSize:10 }}>{emp.code}</td>
                          <td style={{ padding:'8px 10px 8px 0', color:'#fff', fontWeight:700 }}>{emp.totalReports}</td>
                          <td style={{ padding:'8px 10px 8px 0', color:'#9ca3af' }}>{emp.uniqueStores}</td>
                          <td style={{ padding:'8px 10px 8px 0' }}>
                            <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                              <div style={{ width:70, height:5, background:'#1f2937', borderRadius:3, overflow:'hidden' }}>
                                <div style={{ height:'100%', width:`${emp.compliancePct}%`, background:emp.grade.color, borderRadius:3 }}/>
                              </div>
                              <span style={{ fontSize:11, color:emp.grade.color, fontWeight:700 }}>{emp.compliancePct}%</span>
                            </div>
                          </td>
                          <td style={{ padding:'8px 10px 8px 0', color:'#6b7280', fontSize:10 }}>
                            {emp.lastDate?.slice(0,10) || '–'}
                          </td>
                          <td style={{ textAlign:'center', padding:'8px 0' }}>
                            <span style={{
                              fontSize:9, padding:'3px 8px', borderRadius:4, fontWeight:700,
                              background:emp.grade.bg, color:emp.grade.color, letterSpacing:'0.08em'
                            }}>{emp.grade.label}</span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* ─ TAB 1: SUBMISSION MATRIX HEATMAP ─ */}
            {tab === 1 && (
              <div style={{ overflowX: 'auto' }}>
                <div style={{ display:'flex', gap:16, marginBottom:12, alignItems:'center' }}>
                  {[['#06b6d4','Submitted'],['#1f2937','Missing']].map(([c,l])=>(
                    <div key={l} style={{ display:'flex', alignItems:'center', gap:5 }}>
                      <div style={{ width:10, height:10, borderRadius:2, background:c }}/>
                      <span style={{ fontSize:10, color:'#6b7280' }}>{l}</span>
                    </div>
                  ))}
                  <span style={{ fontSize:10, color:'#374151', marginLeft:'auto' }}>{dateKeys.length} TRACKED DAYS</span>
                </div>
                <div style={{ minWidth:'max-content' }}>
                  <table style={{ borderCollapse:'collapse', fontSize:10 }}>
                    <thead>
                      <tr>
                        <th style={{ textAlign:'left', paddingRight:12, paddingBottom:4, color:'#4b5563', fontWeight:600, fontSize:9, position:'sticky', left:0, background:'#030712', minWidth:180 }}>EMPLOYEE</th>
                        <th style={{ textAlign:'right', paddingRight:8, paddingBottom:4, color:'#4b5563', fontWeight:600, fontSize:9 }}>RPT</th>
                        <th style={{ textAlign:'right', paddingRight:12, paddingBottom:4, color:'#4b5563', fontWeight:600, fontSize:9 }}>%</th>
                        {dateKeys.map(d => (
                          <th key={d} style={{ paddingBottom:4, width:13, minWidth:13 }}>
                            <div style={{ writingMode:'vertical-rl', fontSize:8, color:'#374151', lineHeight:1 }}>{d}</div>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {state.submissions.map(sub => {
                        const pct = compliancePct(sub); const g = grade(pct)
                        return (
                          <tr key={sub.code}>
                            <td style={{ paddingRight:12, paddingTop:2, paddingBottom:2, color:'#d1d5db', position:'sticky', left:0, background:'#030712', whiteSpace:'nowrap', fontSize:11 }}>
                              {sub.name.split(' ').slice(0,3).join(' ')}
                            </td>
                            <td style={{ paddingRight:8, textAlign:'right', color:'#6b7280' }}>{sub.totalReports}</td>
                            <td style={{ paddingRight:12, textAlign:'right', fontWeight:700, color:g.color }}>{pct}%</td>
                            {sub.dateKeys.map(d => {
                              const val = sub.dailyMap[d]
                              const hit = val && val !== '–'
                              return (
                                <td key={d} style={{ padding:'1px 0.5px' }}>
                                  <div title={`${d}: ${val||'–'}`} style={{
                                    width:11, height:11, borderRadius:2,
                                    background: hit ? '#06b6d4' : '#111827',
                                    opacity: hit ? 0.9 : 0.6,
                                  }}/>
                                </td>
                              )
                            })}
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* ─ TAB 2: WEEKLY TRENDS ─ */}
            {tab === 2 && (
              <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:20 }}>
                <div style={{ background:'rgba(17,24,39,0.6)', border:'1px solid #1f2937', borderRadius:12, padding:16 }}>
                  <div style={{ fontSize:10, color:'#4b5563', fontWeight:700, letterSpacing:'0.12em', marginBottom:14 }}>TEAM TOTAL VISITS — BY WEEK</div>
                  <ResponsiveContainer width="100%" height={220}>
                    <BarChart data={weeklyTotalChart}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#111827"/>
                      <XAxis dataKey="week" tick={{ fill:'#6b7280', fontSize:9 }} />
                      <YAxis tick={{ fill:'#6b7280', fontSize:10 }} />
                      <Tooltip contentStyle={{ background:'#0f172a', border:'1px solid #1f2937', color:'#e5e7eb', fontSize:11 }}/>
                      <Bar dataKey="visits" fill="#06b6d4" radius={[4,4,0,0]}/>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div style={{ background:'rgba(17,24,39,0.6)', border:'1px solid #1f2937', borderRadius:12, padding:16 }}>
                  <div style={{ fontSize:10, color:'#4b5563', fontWeight:700, letterSpacing:'0.12em', marginBottom:10 }}>INDIVIDUAL WEEKLY BREAKDOWN</div>
                  <div style={{ overflowY:'auto', maxHeight:240 }}>
                    <table style={{ width:'100%', fontSize:11, borderCollapse:'collapse' }}>
                      <thead>
                        <tr style={{ borderBottom:'1px solid #1f2937' }}>
                          {['REP','W1','W2','W3','W4','W5','TOT'].map(h=>(
                            <th key={h} style={{ textAlign:h==='REP'?'left':'right', padding:'4px 6px 4px 0', fontSize:9, color:'#4b5563', letterSpacing:'0.08em' }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {state.weekly
                          .map(w=>({...w,tot:w.w1+w.w2+w.w3+w.w4+w.w5}))
                          .sort((a,b)=>b.tot-a.tot)
                          .map(w=>(
                            <tr key={w.code} style={{ borderBottom:'1px solid rgba(31,41,55,0.4)' }}>
                              <td style={{ padding:'5px 8px 5px 0', color:'#d1d5db' }}>{w.name.split(' ').slice(0,2).join(' ')}</td>
                              {[w.w1,w.w2,w.w3,w.w4,w.w5].map((v,i)=>(
                                <td key={i} style={{ textAlign:'right', padding:'5px 6px 5px 0',
                                  color:v>=5?'#10b981':v>=3?'#06b6d4':v>=1?'#f59e0b':'#374151', fontWeight:v>4?700:400 }}>
                                  {v||'–'}
                                </td>
                              ))}
                              <td style={{ textAlign:'right', padding:'5px 0', color:'#06b6d4', fontWeight:700 }}>{w.tot}</td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* ─ TAB 3: COVERAGE ─ */}
            {tab === 3 && (
              <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:20 }}>
                <div style={{ background:'rgba(17,24,39,0.6)', border:'1px solid #1f2937', borderRadius:12, padding:16 }}>
                  <div style={{ fontSize:10, color:'#4b5563', fontWeight:700, letterSpacing:'0.12em', marginBottom:14 }}>STORE VISITS BY GOVERNORATE</div>
                  <ResponsiveContainer width="100%" height={320}>
                    <BarChart data={govBarData} layout="vertical">
                      <CartesianGrid strokeDasharray="3 3" stroke="#111827"/>
                      <XAxis type="number" tick={{ fill:'#6b7280', fontSize:9 }}/>
                      <YAxis type="category" dataKey="name" tick={{ fill:'#9ca3af', fontSize:9 }} width={105}/>
                      <Tooltip contentStyle={{ background:'#0f172a', border:'1px solid #1f2937', color:'#e5e7eb', fontSize:11 }}/>
                      <Bar dataKey="value" fill="#8b5cf6" radius={[0,4,4,0]}/>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div style={{ background:'rgba(17,24,39,0.6)', border:'1px solid #1f2937', borderRadius:12, padding:16 }}>
                  <div style={{ fontSize:10, color:'#4b5563', fontWeight:700, letterSpacing:'0.12em', marginBottom:10 }}>
                    STORE DIRECTORY — {totalStores} STORES
                  </div>
                  <div style={{ overflowY:'auto', maxHeight:320 }}>
                    <table style={{ width:'100%', fontSize:11, borderCollapse:'collapse' }}>
                      <thead style={{ position:'sticky', top:0, background:'rgba(17,24,39,0.95)' }}>
                        <tr style={{ borderBottom:'1px solid #1f2937' }}>
                          {['STORE NAME','GOV','VISITS','SPVRs'].map(h=>(
                            <th key={h} style={{ textAlign:h==='STORE NAME'?'left':'right', padding:'4px 6px 4px 0', fontSize:9, color:'#4b5563', letterSpacing:'0.08em' }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {state.stores.sort((a,b)=>b.totalVisits-a.totalVisits).map(s=>(
                          <tr key={s.shopCode} style={{ borderBottom:'1px solid rgba(31,41,55,0.3)' }}>
                            <td style={{ padding:'4px 8px 4px 0', color:'#d1d5db' }}>{s.shopName}</td>
                            <td style={{ textAlign:'right', padding:'4px 8px 4px 0', color:'#6b7280' }}>{s.gov}</td>
                            <td style={{ textAlign:'right', padding:'4px 6px 4px 0', color:'#06b6d4', fontWeight:700 }}>{s.totalVisits}</td>
                            <td style={{ textAlign:'right', padding:'4px 0', color:'#6b7280' }}>{s.uniqueSPVRs}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* ─ TAB 4: FIELD REPORTS ─ */}
            {tab === 4 && (
              <div>
                <div style={{ display:'flex', gap:10, marginBottom:14, alignItems:'center', flexWrap:'wrap' }}>
                  <div style={{ position:'relative', flex:1, minWidth:200 }}>
                    <Search size={13} style={{ position:'absolute', left:10, top:'50%', transform:'translateY(-50%)', color:'#4b5563' }}/>
                    <input value={search} onChange={e=>setSearch(e.target.value)}
                      placeholder="Search rep, store, model, or governor..."
                      style={{ width:'100%', background:'#0f172a', border:'1px solid #374151', borderRadius:8, padding:'8px 12px 8px 32px', fontSize:12, color:'#e5e7eb', outline:'none', boxSizing:'border-box' }}
                    />
                  </div>
                  <select value={govFilter} onChange={e=>setGovFilter(e.target.value)} style={{
                    background:'#0f172a', border:'1px solid #374151', borderRadius:8, padding:'8px 12px', fontSize:11, color:'#e5e7eb', outline:'none' }}>
                    {allGovs.map(g=><option key={g} value={g}>{g}</option>)}
                  </select>
                  <span style={{ fontSize:10, color:'#374151' }}>{filteredMobile.length} RECORDS</span>
                </div>
                <div style={{ display:'flex', flexDirection:'column', gap:8, maxHeight:560, overflowY:'auto', paddingRight:4 }}>
                  {filteredMobile.map((r, i) => {
                    const hasShortage = shortageFlag(r.samsungShortage)
                    return (
                      <div key={i} style={{
                        background: hasShortage ? 'rgba(249,115,22,0.05)' : 'rgba(17,24,39,0.6)',
                        border: `1px solid ${hasShortage ? 'rgba(249,115,22,0.25)' : '#1f2937'}`,
                        borderRadius:10, padding:'10px 14px', fontSize:11,
                      }}>
                        <div style={{ display:'flex', justifyContent:'space-between', marginBottom:8 }}>
                          <div>
                            <span style={{ color:'#06b6d4', fontWeight:700 }}>{r.repName.split(' ').slice(0,3).join(' ')}</span>
                            <span style={{ color:'#374151', margin:'0 8px' }}>→</span>
                            <span style={{ color:'#f9fafb', fontWeight:600 }}>{r.shopName}</span>
                            <span style={{ color:'#4b5563', marginLeft:8 }}>{r.gov}</span>
                          </div>
                          <span style={{ color:'#4b5563', fontSize:10 }}>{r.date.slice(0,10)}</span>
                        </div>
                        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:8, marginBottom:8 }}>
                          <div style={{ color: hasShortage?'#f97316':'#6b7280' }}>
                            <span style={{ fontWeight:600 }}>📦 Samsung:</span> {r.samsungShortage || 'No shortage'}
                          </div>
                          <div style={{ color:'#8b5cf6' }}>
                            <span style={{ fontWeight:600 }}>⚔️ {r.compBrand}:</span> <span style={{ color:'#6b7280' }}>{r.compMovement?.slice(0,55)}{r.compMovement?.length>55?'…':''}</span>
                          </div>
                          <div style={{ color:'#10b981' }}>
                            <span style={{ fontWeight:600 }}>✅ Action:</span> <span style={{ color:'#6b7280' }}>{r.action?.slice(0,55)}{r.action?.length>55?'…':''}</span>
                          </div>
                        </div>
                        <div style={{ borderTop:'1px solid #1f2937', paddingTop:6, color:'#4b5563', lineHeight:1.5 }}>
                          💬 {r.sellout?.slice(0,140)}{r.sellout?.length>140?'…':''}
                        </div>
                      </div>
                    )
                  })}
                </div>
              </div>
            )}

            {/* ─ TAB 5: INTEL ─ */}
            {tab === 5 && (
              <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:20 }}>
                <div style={{ background:'rgba(17,24,39,0.6)', border:'1px solid #1f2937', borderRadius:12, padding:16 }}>
                  <div style={{ fontSize:10, color:'#4b5563', fontWeight:700, letterSpacing:'0.12em', marginBottom:14 }}>COMPETITOR BRAND INTEL</div>
                  <ResponsiveContainer width="100%" height={260}>
                    <PieChart>
                      <Pie data={competitorData} cx="50%" cy="50%" outerRadius={95} dataKey="value" nameKey="name"
                        label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} labelLine={false}>
                        {competitorData.map((_,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}
                      </Pie>
                      <Tooltip contentStyle={{ background:'#0f172a', border:'1px solid #1f2937', fontSize:11 }}/>
                    </PieChart>
                  </ResponsiveContainer>
                </div>
                <div style={{ display:'flex', flexDirection:'column', gap:14 }}>
                  <div style={{ background:'rgba(17,24,39,0.6)', border:'1px solid #1f2937', borderRadius:12, padding:16 }}>
                    <div style={{ fontSize:10, color:'#4b5563', fontWeight:700, letterSpacing:'0.12em', marginBottom:10 }}>
                      TOP SAMSUNG SHORTAGE MODELS
                    </div>
                    <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
                      {topShortageModels.map(([model, count]) => (
                        <div key={model} style={{ display:'flex', alignItems:'center', gap:8 }}>
                          <AlertTriangle size={11} style={{ color:'#f97316', flexShrink:0 }}/>
                          <span style={{ fontSize:11, color:'#f97316', fontWeight:600, flex:1 }}>{model}</span>
                          <span style={{ fontSize:11, color:'#6b7280' }}>×{count}</span>
                        </div>
                      ))}
                      {topShortageModels.length === 0 && (
                        <span style={{ fontSize:11, color:'#374151' }}>No shortage data from Mobile_Dep.xlsx</span>
                      )}
                    </div>
                  </div>
                  <div style={{ background:'rgba(17,24,39,0.6)', border:'1px solid #1f2937', borderRadius:12, padding:16 }}>
                    <div style={{ fontSize:10, color:'#4b5563', fontWeight:700, letterSpacing:'0.12em', marginBottom:10 }}>
                      SHORTAGE ALERTS ({shortageAlerts})
                    </div>
                    <div style={{ overflowY:'auto', maxHeight:160, display:'flex', flexDirection:'column', gap:4 }}>
                      {shortageList.slice(0,30).map((r, i) => (
                        <div key={i} style={{ display:'flex', gap:8, fontSize:10, background:'rgba(249,115,22,0.07)', border:'1px solid rgba(249,115,22,0.2)', borderRadius:6, padding:'5px 8px' }}>
                          <span style={{ color:'#f97316', fontWeight:700, flexShrink:0 }}>{r.samsungShortage}</span>
                          <span style={{ color:'#4b5563' }}>at</span>
                          <span style={{ color:'#d1d5db' }}>{r.shopName}</span>
                          <span style={{ color:'#374151', marginLeft:'auto', flexShrink:0 }}>{r.gov}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            )}

          </div>
        </>
      )}
    </div>
  )
}
