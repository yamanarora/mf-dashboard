import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { LineChart, Line, CartesianGrid, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";
import {
  ColumnDef,
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getSortedRowModel,
  SortingState,
  useReactTable,
} from "@tanstack/react-table";

// ------------------------------------------------------------
// Types
// ------------------------------------------------------------

type MFInputRow = {
  ID?: string | number;
  "Holder Name"?: string;
  Category?: string;
  Name?: string; // scheme display name
  "Type of Investment"?: string; // SIP | Lump Sum | Both
  "SIP Amount"?: number;
  "Lump Sum Amount"?: number;
  Issuer?: string; // AMC
  "SIP Start Date"?: string; // YYYY-MM-DD
  "SIP Day"?: number; // 1-31
  "Start NAV"?: number; // for lump sum baseline
  "Buy NAV"?: number; // input for Lump Sum; derived for SIP/Both
  "Amount Invested"?: number; // user input cumulative
};

type MFDerivedFields = {
  "Current Value": number;
  "NAV Today": number;
  "NAV Past Week": number;
  "NAV Past 2 Weeks": number;
  "NAV Past 3 Weeks": number;
  "NAV Past Month": number;
  "NAV Past 2 Months": number;
  "NAV Past 3 Months": number;
  "NAV Past 6 Months": number;
  "NAV Past 9 Months": number;
  "NAV Past Year": number;
  "NAV MTD": number;
  "NAV QTD": number;
  "NAV FYTD": number;
  "NAV Previous FY": number;
  "P&L Today": number;
  "P&L Past Week": number;
  "P&L Past 2 Weeks": number;
  "P&L Past 3 Weeks": number;
  "P&L Past Month": number;
  "P&L Past 2 Months": number;
  "P&L Past 3 Months": number;
  "P&L Past 6 Months": number;
  "P&L Past 9 Months": number;
  "P&L Past Year": number;
  "P&L MTD": number;
  "P&L QTD": number;
  "P&L FYTD": number;
  "P&L Previous FY": number;
  "As % of Total Portfolio - Invested": number;
  "As % of Total Portfolio - Current": number;
};

type MFRow = MFInputRow & Partial<MFDerivedFields> & {
  __schemeCode?: string; // mfapi.in code (internal)
  __notes?: string[]; // warnings / notes
};

// ------------------------------------------------------------
// Constants & utilities
// ------------------------------------------------------------

const HEADER_ORDER: (keyof MFRow)[] = [
  "ID",
  "Holder Name",
  "Category",
  "Name",
  "Type of Investment",
  "SIP Amount",
  "Lump Sum Amount",
  "Issuer",
  "SIP Start Date",
  "SIP Day",
  "Start NAV",
  "Buy NAV",
  "Amount Invested",
  "Current Value",
  "NAV Today",
  "NAV Past Week",
  "NAV Past 2 Weeks",
  "NAV Past 3 Weeks",
  "NAV Past Month",
  "NAV Past 2 Months",
  "NAV Past 3 Months",
  "NAV Past 6 Months",
  "NAV Past 9 Months",
  "NAV Past Year",
  "NAV MTD",
  "NAV QTD",
  "NAV FYTD",
  "NAV Previous FY",
  "P&L Today",
  "P&L Past Week",
  "P&L Past 2 Weeks",
  "P&L Past 3 Weeks",
  "P&L Past Month",
  "P&L Past 2 Months",
  "P&L Past 3 Months",
  "P&L Past 6 Months",
  "P&L Past 9 Months",
  "P&L Past Year",
  "P&L MTD",
  "P&L QTD",
  "P&L FYTD",
  "P&L Previous FY",
  "As % of Total Portfolio - Invested",
  "As % of Total Portfolio - Current",
];

const INR = new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 2 });
const NUM2 = new Intl.NumberFormat("en-IN", { maximumFractionDigits: 2 });

function parseDate(s?: string): Date | null {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(d: Date) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function addDays(d: Date, n: number) {
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
}

function addMonths(d: Date, n: number) {
  const x = new Date(d);
  const day = x.getDate();
  x.setMonth(x.getMonth() + n);
  // handle short months
  if (x.getDate() < day) x.setDate(0);
  return x;
}

function startOfMonth(d: Date) {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

function startOfQuarter(d: Date) {
  const qStartMonth = Math.floor(d.getMonth() / 3) * 3;
  return new Date(d.getFullYear(), qStartMonth, 1);
}

function startOfFYIndia(d: Date) {
  const year = d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1; // Apr=3
  return new Date(year, 3, 1);
}

function endOfPrevFYIndia(d: Date) {
  const fyStart = startOfFYIndia(d);
  return addDays(fyStart, -1); // Mar 31 of prev FY
}

// ------------------------------------------------------------
// mfapi.in helpers
// ------------------------------------------------------------

type MFListItem = { schemeName: string; schemeCode: string };

type MFNavPoint = { date: string; nav: string };

type MFHistory = { code: string; data: MFNavPoint[] };

const schemeListCache: { list?: MFListItem[] } = {};
const navCache: Record<string, MFHistory> = {};

async function fetchSchemeList(): Promise<MFListItem[]> {
  if (schemeListCache.list) return schemeListCache.list;
  const resp = await fetch("https://api.mfapi.in/mf");
  const json = await resp.json();
  schemeListCache.list = json as MFListItem[];
  return schemeListCache.list;
}

async function resolveSchemeCodeByName(name: string): Promise<string | null> {
  const list = await fetchSchemeList();
  const n = name.trim().toLowerCase();
  // exact match first
  const exact = list.find((x) => x.schemeName.trim().toLowerCase() === n);
  if (exact) return exact.schemeCode;
  // startsWith / includes fallback
  const starts = list.find((x) => x.schemeName.trim().toLowerCase().startsWith(n));
  if (starts) return starts.schemeCode;
  const inc = list.find((x) => x.schemeName.trim().toLowerCase().includes(n));
  return inc ? inc.schemeCode : null;
}

async function fetchNavHistory(code: string): Promise<MFHistory> {
  if (navCache[code]) return navCache[code];
  const resp = await fetch(`https://api.mfapi.in/mf/${code}`);
  const json = await resp.json();
  navCache[code] = json as MFHistory;
  return navCache[code];
}

function buildNavMap(history: MFHistory) {
  const map = new Map<string, number>();
  const dates: string[] = [];
  for (const p of history.data) {
    const v = parseFloat(p.nav);
    if (!isNaN(v)) {
      map.set(p.date, v);
      dates.push(p.date);
    }
  }
  dates.sort(); // ascending YYYY-MM-DD
  return { map, dates };
}

function getNavOnOrBefore(target: Date, navDates: string[], navMap: Map<string, number>): number | null {
  const t = fmtDate(target);
  // binary search last date <= t
  let lo = 0, hi = navDates.length - 1, idx = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const cmp = navDates[mid].localeCompare(t);
    if (cmp <= 0) { idx = mid; lo = mid + 1; } else { hi = mid - 1; }
  }
  if (idx >= 0) return navMap.get(navDates[idx]) ?? null;
  return null;
}

function getNavOnOrAfter(target: Date, navDates: string[], navMap: Map<string, number>): { date: string | null; nav: number | null } {
  const t = fmtDate(target);
  let lo = 0, hi = navDates.length - 1, idx = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const cmp = navDates[mid].localeCompare(t);
    if (cmp >= 0) { idx = mid; hi = mid - 1; } else { lo = mid + 1; }
  }
  if (idx >= 0) {
    const d = navDates[idx];
    return { date: d, nav: navMap.get(d) ?? null };
  }
  return { date: null, nav: null };
}

// ------------------------------------------------------------
// SIP math
// ------------------------------------------------------------

function daysInMonth(y: number, m0: number) { // m0: 0-11
  return new Date(y, m0 + 1, 0).getDate();
}

function enumerateSipExecutions(start: Date, sipDay: number, navDates: string[], today: Date): string[] {
  // returns actual execution NAV dates (YYYY-MM-DD) using next available business day rule
  const actualDates: string[] = [];
  const lastPubDateStr = navDates[navDates.length - 1];
  const lastPub = lastPubDateStr ? new Date(lastPubDateStr) : today;
  const end = today < lastPub ? today : lastPub;
  let y = start.getFullYear();
  let m = start.getMonth();
  while (new Date(y, m, 1) <= end) {
    const dim = daysInMonth(y, m);
    const day = Math.min(Math.max(1, sipDay), dim);
    const sched = new Date(y, m, day);
    const { date: execDate } = getNavOnOrAfter(sched, navDates, new Map());
    // NOTE: We can't reuse getNavOnOrAfter with a new Map(); but we only need the date; we'll compute with known list
    // Implement simple search for first date >= sched in navDates
    let dStr: string | null = null;
    for (let i = 0; i < navDates.length; i++) {
      if (navDates[i] >= fmtDate(sched)) { dStr = navDates[i]; break; }
    }
    if (dStr) {
      const dd = new Date(dStr);
      if (dd <= end) actualDates.push(dStr);
    }
    // move to next month
    m += 1; if (m > 11) { m = 0; y += 1; }
  }
  return actualDates;
}

// ------------------------------------------------------------
// Main Component
// ------------------------------------------------------------

export default function MFPortfolioApp() {
  const [rows, setRows] = useState<MFRow[]>([]);
  const [fileName, setFileName] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [notes, setNotes] = useState<string[]>([]);
  const [selectedScheme, setSelectedScheme] = useState<string | null>(null);
  const [navSeries, setNavSeries] = useState<{ date: string; nav: number }[]>([]);
  const [sorting, setSorting] = useState<SortingState>([{ id: "P&L Today", desc: true }] as any);
  const [globalFilter, setGlobalFilter] = useState("");

  // Import Excel: expect a sheet named MFPortfolio; otherwise use first sheet
  function onFile(e: React.ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0];
    if (!f) return;
    setFileName(f.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target?.result;
      if (!data) return;
      const wb = XLSX.read(data, { type: "array" });
      const wsName = wb.SheetNames.find((s) => s === "MFPortfolio") || wb.SheetNames[0];
      const ws = wb.Sheets[wsName];
      const json = XLSX.utils.sheet_to_json<MFRow>(ws, { defval: "" });
      // Clean numeric fields
      const cleaned = json.map((r) => ({
        ...r,
        "SIP Amount": num(r["SIP Amount"]),
        "Lump Sum Amount": num(r["Lump Sum Amount"]),
        "SIP Day": num(r["SIP Day"]),
        "Start NAV": num(r["Start NAV"]),
        "Buy NAV": num(r["Buy NAV"]),
        "Amount Invested": num(r["Amount Invested"]),
      }));
      setRows(cleaned);
    };
    reader.readAsArrayBuffer(f);
  }

  // Utility to coerce
  function num(v: any): number | undefined {
    const n = typeof v === "string" ? parseFloat(v) : typeof v === "number" ? v : NaN;
    return isNaN(n) ? undefined : n;
  }

  // Refresh: resolve scheme codes, pull NAVs, compute derived fields
  async function refresh() {
    setLoading(true);
    const newNotes: string[] = [];
    try {
      // Resolve scheme codes for unique Names
      const names = Array.from(new Set(rows.map((r) => r.Name).filter(Boolean))) as string[];
      const nameToCode = new Map<string, string>();
      for (const nm of names) {
        const code = await resolveSchemeCodeByName(nm!);
        if (code) nameToCode.set(nm!, code); else newNotes.push(`No scheme code found for: ${nm}`);
      }

      // Fetch NAV histories
      const codeToHistory = new Map<string, { map: Map<string, number>; dates: string[] }>();
      for (const code of Array.from(new Set(Array.from(nameToCode.values())))) {
        const hist = await fetchNavHistory(code);
        const { map, dates } = buildNavMap(hist);
        codeToHistory.set(code, { map, dates });
      }

      const today = new Date();
      const refDates: Record<string, Date> = {
        today,
        pastWeek: addDays(today, -7),
        past2W: addDays(today, -14),
        past3W: addDays(today, -21),
        past1M: addMonths(today, -1),
        past2M: addMonths(today, -2),
        past3M: addMonths(today, -3),
        past6M: addMonths(today, -6),
        past9M: addMonths(today, -9),
        past1Y: addMonths(today, -12),
        mtd: startOfMonth(today),
        qtd: startOfQuarter(today),
        fytd: startOfFYIndia(today),
        prevFY: endOfPrevFYIndia(today),
      };

      const out: MFRow[] = [];
      for (const r of rows) {
        const code = r.Name ? nameToCode.get(r.Name) : undefined;
        const hist = code ? codeToHistory.get(code) : undefined;
        const map = hist?.map; const dates = hist?.dates;
        const notesRow: string[] = [];

        let navToday = 0;
        let navPastWeek = 0, navPast2W = 0, navPast3W = 0;
        let nav1M = 0, nav2M = 0, nav3M = 0, nav6M = 0, nav9M = 0, nav1Y = 0;
        let navMTD = 0, navQTD = 0, navFYTD = 0, navPrevFY = 0;

        if (map && dates && dates.length) {
          navToday = getNavOnOrBefore(refDates.today, dates, map) ?? 0;
          navPastWeek = getNavOnOrBefore(refDates.pastWeek, dates, map) ?? 0;
          navPast2W = getNavOnOrBefore(refDates.past2W, dates, map) ?? 0;
          navPast3W = getNavOnOrBefore(refDates.past3W, dates, map) ?? 0;
          nav1M = getNavOnOrBefore(refDates.past1M, dates, map) ?? 0;
          nav2M = getNavOnOrBefore(refDates.past2M, dates, map) ?? 0;
          nav3M = getNavOnOrBefore(refDates.past3M, dates, map) ?? 0;
          nav6M = getNavOnOrBefore(refDates.past6M, dates, map) ?? 0;
          nav9M = getNavOnOrBefore(refDates.past9M, dates, map) ?? 0;
          nav1Y = getNavOnOrBefore(refDates.past1Y, dates, map) ?? 0;
          navMTD = getNavOnOrBefore(refDates.mtd, dates, map) ?? 0;
          navQTD = getNavOnOrBefore(refDates.qtd, dates, map) ?? 0;
          navFYTD = getNavOnOrBefore(refDates.fytd, dates, map) ?? 0;
          navPrevFY = getNavOnOrBefore(refDates.prevFY, dates, map) ?? 0;
        } else {
          notesRow.push("No NAV history found (name mismatch?)");
        }

        // Compute Buy NAV for SIP/Both
        let buyNAV = r["Buy NAV"] ?? 0;
        const type = (r["Type of Investment"] || "").toLowerCase();
        const sipAmount = r["SIP Amount"] ?? 0;
        const lump = r["Lump Sum Amount"] ?? 0;
        const startNAV = r["Start NAV"] ?? 0;
        const amountInvested = r["Amount Invested"] ?? 0;

        if ((type.includes("sip") || type.includes("both")) && map && dates && dates.length && r["SIP Start Date"] && r["SIP Day"]) {
          const start = parseDate(r["SIP Start Date"]!);
          const sipDay = r["SIP Day"]!;
          if (start) {
            // Lump sum units
            const unitsLump = lump > 0 && startNAV > 0 ? lump / startNAV : 0;
            // Enumerate executed SIPs
            const execDates = enumerateSipExecutions(start, sipDay, dates, new Date());
            // Units from SIPs
            let unitsSip = 0;
            for (const dStr of execDates) {
              const n = map.get(dStr) ?? 0;
              if (n > 0 && sipAmount > 0) unitsSip += sipAmount / n;
            }
            const totalUnits = unitsLump + unitsSip;
            const totalCost = lump + execDates.length * sipAmount;
            if (totalUnits > 0) buyNAV = totalCost / totalUnits;
            else notesRow.push("No executed SIPs yet; Buy NAV left as-is");

            // sanity vs Amount Invested (user input)
            if (amountInvested > 0 && Math.abs(totalCost - amountInvested) / amountInvested > 0.01) {
              notesRow.push(`Invested mismatch: input ${amountInvested} vs computed ${totalCost}`);
            }
          }
        }

        // Derived values
        const A = amountInvested > 0 ? amountInvested : 0;
        const NAV0 = buyNAV > 0 ? buyNAV : 0;
        const scale = NAV0 > 0 ? (navToday / NAV0) : 0;
        const currentValue = A * scale;

        function pnl(nav: number) { return NAV0 > 0 ? A * (nav / NAV0 - 1) : 0; }

        const derived: MFDerivedFields = {
          "Current Value": currentValue,
          "NAV Today": navToday,
          "NAV Past Week": navPastWeek,
          "NAV Past 2 Weeks": navPast2W,
          "NAV Past 3 Weeks": navPast3W,
          "NAV Past Month": nav1M,
          "NAV Past 2 Months": nav2M,
          "NAV Past 3 Months": nav3M,
          "NAV Past 6 Months": nav6M,
          "NAV Past 9 Months": nav9M,
          "NAV Past Year": nav1Y,
          "NAV MTD": navMTD,
          "NAV QTD": navQTD,
          "NAV FYTD": navFYTD,
          "NAV Previous FY": navPrevFY,
          "P&L Today": pnl(navToday),
          "P&L Past Week": pnl(navPastWeek),
          "P&L Past 2 Weeks": pnl(navPast2W),
          "P&L Past 3 Weeks": pnl(navPast3W),
          "P&L Past Month": pnl(nav1M),
          "P&L Past 2 Months": pnl(nav2M),
          "P&L Past 3 Months": pnl(nav3M),
          "P&L Past 6 Months": pnl(nav6M),
          "P&L Past 9 Months": pnl(nav9M),
          "P&L Past Year": pnl(nav1Y),
          "P&L MTD": pnl(navMTD),
          "P&L QTD": pnl(navQTD),
          "P&L FYTD": pnl(navFYTD),
          "P&L Previous FY": pnl(navPrevFY),
          "As % of Total Portfolio - Invested": 0, // fill later
          "As % of Total Portfolio - Current": 0,
        };

        out.push({ ...r, ...derived, "Buy NAV": buyNAV, __schemeCode: code, __notes: notesRow });
      }

      // Fill portfolio %s
      const totalInvested = out.reduce((s, r) => s + (r["Amount Invested"] ?? 0), 0);
      const totalCurrent = out.reduce((s, r) => s + (r["Current Value"] ?? 0), 0);
      for (const r of out) {
        r["As % of Total Portfolio - Invested"] = totalInvested > 0 ? (r["Amount Invested"] ?? 0) / totalInvested : 0;
        r["As % of Total Portfolio - Current"] = totalCurrent > 0 ? (r["Current Value"] ?? 0) / totalCurrent : 0;
      }

      setRows(out);
      setNotes(newNotes);
    } catch (e: any) {
      setNotes(["Refresh failed: " + (e?.message || String(e))]);
    } finally {
      setLoading(false);
    }
  }

  // Export to Excel (writes all headers in canonical order)
  function exportExcel() {
    const data = rows.map((r) => {
      const o: any = {};
      for (const h of HEADER_ORDER) o[h as string] = (r as any)[h] ?? "";
      return o;
    });
    const ws = XLSX.utils.json_to_sheet(data, { header: HEADER_ORDER as string[] });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MFPortfolio");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fileName ? fileName.replace(/\.xlsx?$/i, "_derived.xlsx") : "MFPortfolio_derived.xlsx";
    a.click();
    URL.revokeObjectURL(a.href);
  }

  // Simple analytics
  const kpis = useMemo(() => {
    const invested = rows.reduce((s, r) => s + (r["Amount Invested"] ?? 0), 0);
    const current = rows.reduce((s, r) => s + (r["Current Value"] ?? 0), 0);
    const pnl = current - invested;
    const ret = invested > 0 ? pnl / invested : 0;
    return { invested, current, pnl, ret };
  }, [rows]);

  // Table setup
  const columns = useMemo<ColumnDef<MFRow>[]>(() => [
    { accessorKey: "Name", header: "Scheme", cell: (info) => <div className="font-medium">{String(info.getValue() || "")}</div> },
    { accessorKey: "Holder Name", header: "Holder" },
    { accessorKey: "Category", header: "Category" },
    { accessorKey: "Type of Investment", header: "Type" },
    { accessorKey: "SIP Amount", header: "SIP", cell: ({ getValue }) => INR.format(Number(getValue() || 0)) },
    { accessorKey: "Lump Sum Amount", header: "Lump", cell: ({ getValue }) => INR.format(Number(getValue() || 0)) },
    { accessorKey: "Buy NAV", header: "Buy NAV", cell: ({ getValue }) => NUM2.format(Number(getValue() || 0)) },
    { accessorKey: "Amount Invested", header: "Invested", cell: ({ getValue }) => INR.format(Number(getValue() || 0)) },
    { accessorKey: "Current Value", header: "Current", cell: ({ getValue }) => INR.format(Number(getValue() || 0)) },
    { accessorKey: "P&L Today", header: "P&L", cell: ({ getValue }) => INR.format(Number(getValue() || 0)) },
    { accessorKey: "As % of Total Portfolio - Current", header: "% of Port", cell: ({ getValue }) => (Number(getValue() || 0) * 100).toFixed(1) + "%" },
  ], []);

  const table = useReactTable({
    data: rows,
    columns,
    state: { sorting, globalFilter },
    onSortingChange: setSorting,
    onGlobalFilterChange: setGlobalFilter,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
  });

  // Chart: when a row is clicked, load last 180 days NAV series
  async function onRowClick(r: MFRow) {
    const code = r.__schemeCode || (r.Name ? await resolveSchemeCodeByName(r.Name) : null);
    if (!code) return;
    const hist = await fetchNavHistory(code);
    const series = hist.data.slice(0, 180).reverse().map((p) => ({ date: p.date, nav: parseFloat(p.nav) })).filter((x) => !isNaN(x.nav));
    setSelectedScheme(r.Name || null);
    setNavSeries(series);
  }

  return (
    <div className="min-h-screen bg-neutral-50 text-neutral-900">
      <header className="sticky top-0 z-10 bg-white/80 backdrop-blur border-b border-neutral-200">
        <div className="max-w-7xl mx-auto px-4 py-3 flex flex-wrap items-center gap-3">
          <h1 className="text-xl font-semibold">Mutual Fund Dashboard (Local)</h1>
          <div className="ml-auto flex items-center gap-2">
            <label className="px-3 py-2 rounded-lg border border-neutral-300 bg-white hover:bg-neutral-50 cursor-pointer text-sm">
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={onFile} />
              {fileName ? `Open: ${fileName}` : "Open Excel"}
            </label>
            <button onClick={refresh} disabled={loading || rows.length === 0} className="px-3 py-2 rounded-lg bg-black text-white text-sm disabled:opacity-40">
              {loading ? "Refreshing..." : "Refresh NAVs"}
            </button>
            <button onClick={exportExcel} disabled={rows.length === 0} className="px-3 py-2 rounded-lg border border-neutral-300 bg-white text-sm disabled:opacity-40">Export Excel</button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-4 space-y-6">
        {/* KPIs */}
        <section className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <KPI title="Invested" value={INR.format(kpis.invested)} />
          <KPI title="Current" value={INR.format(kpis.current)} />
          <KPI title="P&L" value={INR.format(kpis.pnl)} trend={kpis.pnl >= 0 ? "up" : "down"} />
          <KPI title="Return" value={(kpis.ret * 100).toFixed(2) + "%"} />
        </section>

        {/* Search */}
        <div className="flex items-center gap-2">
          <input
            value={globalFilter ?? ""}
            onChange={(e) => setGlobalFilter(e.target.value)}
            placeholder="Search schemes, holder, category..."
            className="w-full md:w-1/2 px-3 py-2 border border-neutral-300 rounded-lg"
          />
          {notes.length > 0 && (
            <div className="text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
              {notes.slice(0, 3).join(" • ")}{notes.length > 3 ? ` (+${notes.length - 3} more)` : ""}
            </div>
          )}
        </div>

        {/* Table */}
        <div className="overflow-auto border border-neutral-200 rounded-xl bg-white">
          <table className="min-w-full text-sm">
            <thead className="bg-neutral-50">
              {table.getHeaderGroups().map((hg) => (
                <tr key={hg.id}>
                  {hg.headers.map((h) => (
                    <th key={h.id} className="text-left px-3 py-2 font-semibold text-neutral-700 select-none cursor-pointer" onClick={h.column.getToggleSortingHandler()}>
                      <div className="flex items-center gap-1">
                        {flexRender(h.column.columnDef.header, h.getContext())}
                        {{ asc: "▲", desc: "▼" }[h.column.getIsSorted() as string] ?? null}
                      </div>
                    </th>
                  ))}
                </tr>
              ))}
            </thead>
            <tbody>
              {table.getRowModel().rows.map((row) => (
                <tr key={row.id} className="border-t border-neutral-100 hover:bg-neutral-50 cursor-pointer" onClick={() => onRowClick(row.original)}>
                  {row.getVisibleCells().map((cell) => (
                    <td key={cell.id} className="px-3 py-2 whitespace-nowrap">
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Charts */}
        {selectedScheme && (
          <section className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="md:col-span-2 p-4 border border-neutral-200 rounded-xl bg-white">
              <h3 className="font-semibold mb-3">{selectedScheme} — Recent NAV (≈180 days)</h3>
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={navSeries}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="date" minTickGap={24} />
                    <YAxis />
                    <Tooltip formatter={(v: any) => NUM2.format(Number(v))} />
                    <Line type="monotone" dataKey="nav" dot={false} />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="p-4 border border-neutral-200 rounded-xl bg-white">
              <h3 className="font-semibold mb-3">Allocation (Current)</h3>
              <AllocationPie rows={rows} />
            </div>
          </section>
        )}

        {/* Helper text */}
        <p className="text-xs text-neutral-500">
          Notes: Buy NAV is auto‑computed for SIP/Both based on executed monthly SIPs (SIP Day, next business day). Current Value & P&L use your Amount Invested input. Use Export to save a derived copy of your workbook.
        </p>
      </main>
    </div>
  );
}

// ------------------------------------------------------------
// Small UI bits
// ------------------------------------------------------------

function KPI({ title, value, trend }: { title: string; value: string; trend?: "up" | "down" }) {
  return (
    <div className="p-4 rounded-xl border border-neutral-200 bg-white flex flex-col gap-1">
      <div className="text-xs uppercase tracking-wide text-neutral-500">{title}</div>
      <div className="text-xl font-semibold">{value}</div>
      {trend && <div className={trend === "up" ? "text-green-600" : "text-red-600"}>{trend === "up" ? "▲" : "▼"}</div>}
    </div>
  );
}

function AllocationPie({ rows }: { rows: MFRow[] }) {
  const data = useMemo(() => {
    const groups = new Map<string, number>();
    for (const r of rows) {
      const cat = (r.Category || "(uncat)").trim();
      groups.set(cat, (groups.get(cat) || 0) + (r["Current Value"] || 0));
    }
    const total = Array.from(groups.values()).reduce((a, b) => a + b, 0) || 1;
    return Array.from(groups.entries()).map(([name, v]) => ({ name, value: v, pct: (v / total) * 100 }));
  }, [rows]);

  const COLORS = ["#4f46e5", "#22c55e", "#f59e0b", "#ef4444", "#06b6d4", "#a855f7", "#84cc16"]; // not strictly necessary but helps visually

  return (
    <div className="h-64">
      <ResponsiveContainer width="100%" height="100%">
        <PieChart>
          <Pie data={data} dataKey="value" nameKey="name" outerRadius={90}>
            {data.map((entry, i) => (
              <Cell key={i} fill={COLORS[i % COLORS.length]} />
            ))}
          </Pie>
          <Tooltip formatter={(v: any) => INR.format(Number(v))} />
        </PieChart>
      </ResponsiveContainer>
      <div className="mt-2 grid grid-cols-2 gap-1 text-xs">
        {data.map((d, i) => (
          <div key={i} className="flex items-center gap-2">
            <span className="inline-block w-2 h-2 rounded-full" style={{ backgroundColor: ["#4f46e5","#22c55e","#f59e0b","#ef4444","#06b6d4","#a855f7","#84cc16"][i % 7] }} />
            <span className="truncate">{d.name}</span>
            <span className="ml-auto">{d.pct.toFixed(1)}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}
