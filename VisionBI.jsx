import React, { useState, useEffect, useRef, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from "recharts";
import {
  Upload, Sparkles, TrendingUp, TrendingDown, Database, MessageSquare,
  Send, Download, CheckCircle2, AlertTriangle, FileSpreadsheet, Moon, Sun,
  Loader2, RefreshCw, Radio, LayoutGrid, Table2, Bot,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/* Pipeline steps (mirrors the VisionBI processing workflow)          */
/* ------------------------------------------------------------------ */
const PIPELINE_STEPS = [
  "Detecting file type",
  "Reading dataset",
  "Cleaning & validating data",
  "Detecting schema",
  "Understanding business context",
  "Discovering relationships",
  "Computing KPIs",
  "Selecting visualizations",
  "Generating AI insights",
];

const CHART_COLORS = ["#56E8CC", "#F5B942", "#A78BFA", "#FB7185", "#60A5FA", "#34D399"];

/* ------------------------------------------------------------------ */
/* Business-meaning inference                                         */
/* ------------------------------------------------------------------ */
const BUSINESS_RULES = [
  { re: /salary|wage|compensation/i, label: "Employee Salary" },
  { re: /revenue|sales/i, label: "Sales Revenue" },
  { re: /profit|margin/i, label: "Profitability Metric" },
  { re: /region|state|country|city|territory/i, label: "Geographic Dimension" },
  { re: /join.*date|hire.*date/i, label: "Time Dimension (Tenure)" },
  { re: /order.*date|txn.*date|transaction.*date/i, label: "Time Dimension (Transaction)" },
  { re: /date$/i, label: "Time Dimension" },
  { re: /customer/i, label: "Customer Attribute" },
  { re: /employee|staff/i, label: "Workforce Attribute" },
  { re: /product|sku|category/i, label: "Product Dimension" },
  { re: /quantity|units|qty/i, label: "Volume Metric" },
  { re: /(^id$|_id$|id$)/i, label: "Identifier" },
];

function businessLabel(name, type) {
  for (const rule of BUSINESS_RULES) if (rule.re.test(name)) return rule.label;
  if (type === "date") return "Time Dimension";
  if (type === "id") return "Identifier";
  if (type === "numeric") return "Numeric Measure";
  return "Categorical Attribute";
}

/* ------------------------------------------------------------------ */
/* Type detection & profiling                                         */
/* ------------------------------------------------------------------ */
function toDateSafe(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "string") {
    const d = new Date(v);
    if (!isNaN(d) && /\d{2,4}[-/]\d{1,2}[-/]\d{1,4}|\d{4}-\d{2}/.test(v)) return d;
  }
  return null;
}

function detectType(colName, values) {
  const nonNull = values.filter((v) => v !== null && v !== undefined && v !== "");
  if (nonNull.length === 0) return "unknown";
  const uniqueCount = new Set(nonNull.map(String)).size;

  if (/(^id$|_id$|id$)/i.test(colName.trim()) && uniqueCount >= nonNull.length * 0.95) return "id";

  const dateLike = nonNull.filter((v) => toDateSafe(v) !== null).length;
  if (dateLike / nonNull.length > 0.85) return "date";

  const numLike = nonNull.filter(
    (v) => typeof v === "number" || (v !== "" && !isNaN(parseFloat(v)) && isFinite(v))
  ).length;
  if (numLike / nonNull.length > 0.85) return "numeric";

  return "categorical";
}

function computeProfile(rows, colNames) {
  const rowCount = rows.length;
  const seen = new Set();
  let duplicates = 0;
  rows.forEach((r) => {
    const key = JSON.stringify(r);
    if (seen.has(key)) duplicates++;
    else seen.add(key);
  });

  const columns = colNames.map((name) => {
    const values = rows.map((r) => r[name]);
    const missing = values.filter((v) => v === null || v === undefined || v === "").length;
    const type = detectType(name, values);
    const unique = new Set(values.filter((v) => v !== null && v !== undefined && v !== "").map(String)).size;
    return {
      name,
      type,
      missing,
      missingPct: rowCount ? Math.round((missing / rowCount) * 1000) / 10 : 0,
      unique,
      businessLabel: businessLabel(name, type),
    };
  });

  const totalCells = rowCount * colNames.length;
  const totalMissing = columns.reduce((s, c) => s + c.missing, 0);
  const completeness = totalCells ? 1 - totalMissing / totalCells : 1;
  const dupRatio = rowCount ? duplicates / rowCount : 0;
  const qualityScore = Math.round(Math.max(0, Math.min(100, (completeness * 0.7 + (1 - dupRatio) * 0.3) * 100)));

  return { rowCount, duplicates, columns, qualityScore };
}

/* ------------------------------------------------------------------ */
/* KPI + chart aggregation                                            */
/* ------------------------------------------------------------------ */
function titleCase(s) {
  return s.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
}

function formatNumber(n, colName = "") {
  const isMoney = /revenue|sales|profit|price|cost|amount|salary|income/i.test(colName);
  const abs = Math.abs(n);
  let str;
  if (abs >= 1e6) str = (n / 1e6).toFixed(1) + "M";
  else if (abs >= 1e3) str = (n / 1e3).toFixed(1) + "K";
  else str = Number.isInteger(n) ? String(n) : n.toFixed(1);
  return isMoney ? `$${str}` : str;
}

function computeGrowth(rows, dateCol, numCol) {
  const monthly = {};
  rows.forEach((r) => {
    const d = toDateSafe(r[dateCol]);
    if (!d) return;
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    monthly[key] = (monthly[key] || 0) + (parseFloat(r[numCol]) || 0);
  });
  const keys = Object.keys(monthly).sort();
  if (keys.length < 2) return null;
  const first = monthly[keys[0]];
  const last = monthly[keys[keys.length - 1]];
  if (!first) return null;
  return ((last - first) / Math.abs(first)) * 100;
}

function aggregateByMonth(rows, dateCol, numCol) {
  const monthly = {};
  rows.forEach((r) => {
    const d = toDateSafe(r[dateCol]);
    if (!d) return;
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    monthly[key] = (monthly[key] || 0) + (parseFloat(r[numCol]) || 0);
  });
  return Object.keys(monthly)
    .sort()
    .map((k) => ({ month: k, value: Math.round(monthly[k] * 100) / 100 }));
}

function aggregateByCategory(rows, catCol, numCol, topN = 8) {
  const map = {};
  rows.forEach((r) => {
    const key = String(r[catCol] ?? "Unknown");
    map[key] = (map[key] || 0) + (numCol ? parseFloat(r[numCol]) || 0 : 1);
  });
  return Object.entries(map)
    .sort((a, b) => b[1] - a[1])
    .slice(0, topN)
    .map(([name, value]) => ({ name, value: Math.round(value * 100) / 100 }));
}

function computeKPIs(rows, profile) {
  const numericCols = profile.columns.filter((c) => c.type === "numeric");
  const dateCols = profile.columns.filter((c) => c.type === "date");
  const priority = ["revenue", "sales", "profit", "amount", "total"];

  const scored = numericCols
    .map((c) => {
      let score = 0;
      priority.forEach((p, i) => {
        if (new RegExp(p, "i").test(c.name)) score += priority.length - i;
      });
      return { ...c, score };
    })
    .sort((a, b) => b.score - a.score);

  const kpis = [];
  const primary = scored[0];
  const secondary = scored[1];

  if (primary) {
    const sum = rows.reduce((s, r) => s + (parseFloat(r[primary.name]) || 0), 0);
    kpis.push({ label: `Total ${titleCase(primary.name)}`, value: formatNumber(sum, primary.name), tone: "up" });
  }
  if (secondary) {
    const sum = rows.reduce((s, r) => s + (parseFloat(r[secondary.name]) || 0), 0);
    kpis.push({ label: `Total ${titleCase(secondary.name)}`, value: formatNumber(sum, secondary.name), tone: "up" });
  }
  if (primary) {
    const avg = rows.reduce((s, r) => s + (parseFloat(r[primary.name]) || 0), 0) / (rows.length || 1);
    kpis.push({ label: `Avg ${titleCase(primary.name)} / Record`, value: formatNumber(avg, primary.name) });
  }
  kpis.push({ label: "Records Analyzed", value: rows.length.toLocaleString() });
  if (dateCols[0] && primary) {
    const growth = computeGrowth(rows, dateCols[0].name, primary.name);
    if (growth !== null) {
      kpis.push({
        label: `${titleCase(primary.name)} Growth`,
        value: `${growth >= 0 ? "+" : ""}${growth.toFixed(1)}%`,
        tone: growth >= 0 ? "up" : "down",
      });
    }
  }
  kpis.push({ label: "Data Quality Score", value: `${profile.qualityScore}/100`, tone: profile.qualityScore >= 70 ? "up" : "down" });

  return kpis.slice(0, 6);
}

/* ------------------------------------------------------------------ */
/* Sample dataset generator                                           */
/* ------------------------------------------------------------------ */
function generateSampleData() {
  const regions = ["North", "South", "East", "West"];
  const categories = ["Electronics", "Apparel", "Home Goods", "Grocery", "Sports"];
  const rows = [];
  for (let m = 0; m < 12; m++) {
    const seasonal = 1 + 0.25 * Math.sin((m / 12) * Math.PI * 2) + (m >= 9 ? 0.3 : 0);
    for (let i = 0; i < 15; i++) {
      const day = 1 + Math.floor(Math.random() * 27);
      const date = new Date(2025, m, day);
      const region = regions[Math.floor(Math.random() * regions.length)];
      const category = categories[Math.floor(Math.random() * categories.length)];
      const baseSales = 200 + Math.random() * 800;
      const sales = Math.round(baseSales * seasonal * (1 + m * 0.015));
      const profit = Math.round(sales * (0.12 + Math.random() * 0.18));
      const units = Math.max(1, Math.round(sales / (30 + Math.random() * 40)));
      rows.push({
        OrderDate: date.toISOString().slice(0, 10),
        Region: region,
        Category: category,
        Product: `${category.slice(0, 3).toUpperCase()}-${100 + Math.floor(Math.random() * 50)}`,
        CustomerID: `CUST-${1000 + Math.floor(Math.random() * 300)}`,
        Sales: sales,
        Profit: profit,
        Units: units,
      });
    }
  }
  return rows;
}

/* ------------------------------------------------------------------ */
/* AI calls (real Claude API — powers the "AI-generated" pieces)      */
/* ------------------------------------------------------------------ */
async function callClaude(userContent, system) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      system,
      messages: [{ role: "user", content: userContent }],
    }),
  });
  const data = await res.json();
  return (data.content || []).map((b) => b.text || "").join("\n");
}

function parseJSONLoose(text) {
  const cleaned = text.replace(/```json|```/g, "").trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    const match = cleaned.match(/\{[\s\S]*\}/);
    if (match) return JSON.parse(match[0]);
    throw e;
  }
}

async function generateAIInsights(profile, kpis, sampleRows) {
  const system = `You are VisionBI, an AI business intelligence analyst. Respond ONLY with a raw JSON object, no markdown fences, no preamble, matching exactly this shape:
{"headline": string, "summary": string, "insights": string[4 to 6 items], "recommendations": string[2 to 3 items]}
Ground every claim in the data provided below. Be specific, quantify where possible, and use a concise business tone.`;
  const user = `Dataset profile (columns):
${JSON.stringify(profile.columns.map((c) => ({ name: c.name, type: c.type, businessLabel: c.businessLabel, missingPct: c.missingPct })))}

Row count: ${profile.rowCount}
Data quality score: ${profile.qualityScore}/100

KPIs:
${JSON.stringify(kpis)}

Sample rows:
${JSON.stringify(sampleRows)}`;
  const text = await callClaude(user, system);
  return parseJSONLoose(text);
}

function generateFallbackInsights(profile, kpis) {
  const primary = kpis[0];
  return {
    headline: primary ? `${primary.label}: ${primary.value}` : "Dataset analyzed successfully.",
    summary: `Analyzed ${profile.rowCount.toLocaleString()} records across ${profile.columns.length} columns with a data quality score of ${profile.qualityScore}/100.`,
    insights: [
      `Dataset contains ${profile.rowCount.toLocaleString()} rows and ${profile.columns.length} columns.`,
      `${profile.duplicates} duplicate record(s) detected during cleaning.`,
      `Data completeness is at ${profile.qualityScore}%, ${profile.qualityScore >= 80 ? "which is strong" : "with room to improve via cleaning"}.`,
      primary ? `${primary.label} currently stands at ${primary.value}.` : "Key numeric measures were identified for tracking.",
    ],
    recommendations: [
      "Review columns with high missing-value percentages before deeper analysis.",
      "Re-run AI insight generation once connectivity to the AI service is restored for deeper narrative analysis.",
    ],
  };
}

async function askAI(question, context) {
  const system = `You are VisionBI's embedded AI data analyst. Answer the user's question about their dataset using only the schema, KPI, and aggregate context provided. Be concise (2-4 sentences), cite real numbers when possible, and say plainly if the context doesn't contain what's needed to answer.`;
  const user = `Context:\n${JSON.stringify(context)}\n\nQuestion: ${question}`;
  return await callClaude(user, system);
}

/* ------------------------------------------------------------------ */
/* UI subcomponents                                                    */
/* ------------------------------------------------------------------ */
function KPICard({ kpi }) {
  const Icon = kpi.tone === "down" ? TrendingDown : TrendingUp;
  return (
    <div className="vb-kpi">
      <div className="vb-kpi-label">{kpi.label}</div>
      <div className="vb-kpi-value">{kpi.value}</div>
      {kpi.tone && (
        <div className={`vb-kpi-tone vb-kpi-tone--${kpi.tone}`}>
          <Icon size={13} />
        </div>
      )}
    </div>
  );
}

function TypeBadge({ type }) {
  return <span className={`vb-badge vb-badge--${type}`}>{type}</span>;
}

/* ------------------------------------------------------------------ */
/* Main app                                                            */
/* ------------------------------------------------------------------ */
export default function VisionBIDashboard() {
  const [theme, setTheme] = useState("dark");
  const [stage, setStage] = useState("upload"); // upload | processing | dashboard
  const [dragOver, setDragOver] = useState(false);
  const [error, setError] = useState(null);
  const [stepIndex, setStepIndex] = useState(0);

  const [rawRows, setRawRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [profile, setProfile] = useState(null);
  const [kpis, setKpis] = useState([]);
  const [trendData, setTrendData] = useState([]);
  const [catData, setCatData] = useState([]);
  const [pieData, setPieData] = useState([]);
  const [chartMeta, setChartMeta] = useState({});
  const [aiInsights, setAiInsights] = useState(null);
  const [tickerIndex, setTickerIndex] = useState(0);

  const [chatMessages, setChatMessages] = useState([]);
  const [chatInput, setChatInput] = useState("");
  const [chatLoading, setChatLoading] = useState(false);
  const fileInputRef = useRef(null);
  const chatEndRef = useRef(null);

  const onDataLoaded = useCallback((data, name) => {
    const cleaned = (data || []).filter((r) => Object.values(r).some((v) => v !== "" && v !== null && v !== undefined));
    if (cleaned.length === 0) {
      setError("This file appears to be empty or unreadable.");
      return;
    }
    setError(null);
    setFileName(name || "sample-data");
    setRawRows(cleaned);
    setStepIndex(0);
    setStage("processing");
  }, []);

  const handleFile = useCallback(
    (file) => {
      const ext = file.name.split(".").pop().toLowerCase();
      if (["csv", "tsv", "txt"].includes(ext)) {
        Papa.parse(file, {
          header: true,
          dynamicTyping: true,
          skipEmptyLines: true,
          complete: (res) => onDataLoaded(res.data, file.name),
          error: (err) => setError("Could not parse file: " + err.message),
        });
      } else if (["xlsx", "xls"].includes(ext)) {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            onDataLoaded(json, file.name);
          } catch (err) {
            setError("Could not read Excel file: " + err.message);
          }
        };
        reader.readAsArrayBuffer(file);
      } else {
        setError("Unsupported file type. Please upload a .csv, .tsv, .xlsx, or .xls file.");
      }
    },
    [onDataLoaded]
  );

  function onDrop(e) {
    e.preventDefault();
    setDragOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) handleFile(f);
  }

  function loadSample() {
    onDataLoaded(generateSampleData(), "sample_retail_sales.csv");
  }

  function resetAll() {
    setStage("upload");
    setRawRows([]);
    setProfile(null);
    setKpis([]);
    setAiInsights(null);
    setChatMessages([]);
    setTrendData([]);
    setCatData([]);
    setPieData([]);
    setError(null);
  }

  /* Pipeline animation + real computation */
  useEffect(() => {
    if (stage !== "processing") return;
    let cancelled = false;

    async function run() {
      for (let i = 0; i < PIPELINE_STEPS.length; i++) {
        await new Promise((r) => setTimeout(r, 230));
        if (cancelled) return;
        setStepIndex(i + 1);
      }
      try {
        const colNames = Object.keys(rawRows[0] || {});
        const prof = computeProfile(rawRows, colNames);
        const kpiList = computeKPIs(rawRows, prof);

        const dateCol = prof.columns.find((c) => c.type === "date");
        const numCols = prof.columns.filter((c) => c.type === "numeric");
        const catCols = prof.columns.filter((c) => c.type === "categorical" && c.unique <= 30 && c.unique >= 2);

        const primaryNum =
          numCols.slice().sort((a, b) => {
            const score = (c) => (/revenue|sales|profit|amount/i.test(c.name) ? 1 : 0);
            return score(b) - score(a);
          })[0] || numCols[0];

        const trend = dateCol && primaryNum ? aggregateByMonth(rawRows, dateCol.name, primaryNum.name) : [];
        const cats = catCols[0] ? aggregateByCategory(rawRows, catCols[0].name, primaryNum ? primaryNum.name : null, 8) : [];
        const pie = catCols[1]
          ? aggregateByCategory(rawRows, catCols[1].name, null, 6)
          : catCols[0]
          ? aggregateByCategory(rawRows, catCols[0].name, null, 6)
          : [];

        if (cancelled) return;
        setProfile(prof);
        setKpis(kpiList);
        setTrendData(trend);
        setCatData(cats);
        setPieData(pie);
        setChartMeta({
          dateCol: dateCol?.name,
          catCol: catCols[0]?.name,
          pieCol: catCols[1]?.name || catCols[0]?.name,
          numCol: primaryNum?.name,
        });

        let ai;
        try {
          ai = await generateAIInsights(prof, kpiList, rawRows.slice(0, 5));
        } catch (e) {
          ai = generateFallbackInsights(prof, kpiList);
        }
        if (cancelled) return;
        setAiInsights(ai);
        setStage("dashboard");
      } catch (e) {
        if (!cancelled) {
          setError("Something went wrong analyzing this dataset: " + e.message);
          setStage("upload");
        }
      }
    }
    run();
    return () => {
      cancelled = true;
    };
  }, [stage, rawRows]);

  /* Ticker rotation */
  useEffect(() => {
    if (!aiInsights) return;
    const items = [aiInsights.headline, ...(aiInsights.insights || [])].filter(Boolean);
    if (items.length < 2) return;
    const id = setInterval(() => setTickerIndex((i) => (i + 1) % items.length), 4200);
    return () => clearInterval(id);
  }, [aiInsights]);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [chatMessages, chatLoading]);

  async function handleAskAI(e) {
    e.preventDefault();
    const q = chatInput.trim();
    if (!q || chatLoading || !profile) return;
    setChatMessages((m) => [...m, { role: "user", text: q }]);
    setChatInput("");
    setChatLoading(true);
    try {
      const context = {
        columns: profile.columns.map((c) => ({ name: c.name, type: c.type, businessLabel: c.businessLabel })),
        kpis,
        monthlyTrend: trendData,
        topCategories: catData,
        rowCount: profile.rowCount,
      };
      const answer = await askAI(q, context);
      setChatMessages((m) => [...m, { role: "ai", text: answer || "I couldn't generate an answer right now." }]);
    } catch (err) {
      setChatMessages((m) => [...m, { role: "ai", text: "I ran into an issue reaching the AI service. Please try again." }]);
    } finally {
      setChatLoading(false);
    }
  }

  function exportSummary() {
    const lines = [];
    lines.push(`# VisionBI Report`);
    lines.push(``);
    lines.push(`**Dataset:** ${fileName}`);
    lines.push(`**Generated:** ${new Date().toLocaleString()}`);
    lines.push(``);
    lines.push(`## Executive Summary`);
    lines.push(aiInsights?.summary || "");
    lines.push(``);
    lines.push(`## Key Metrics`);
    kpis.forEach((k) => lines.push(`- **${k.label}:** ${k.value}`));
    lines.push(``);
    lines.push(`## AI Insights`);
    (aiInsights?.insights || []).forEach((i) => lines.push(`- ${i}`));
    lines.push(``);
    lines.push(`## Recommendations`);
    (aiInsights?.recommendations || []).forEach((r) => lines.push(`- ${r}`));
    lines.push(``);
    lines.push(`## Data Quality`);
    lines.push(`- Rows: ${profile?.rowCount}`);
    lines.push(`- Duplicate rows: ${profile?.duplicates}`);
    lines.push(`- Quality score: ${profile?.qualityScore}/100`);
    const blob = new Blob([lines.join("\n")], { type: "text/markdown" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "visionbi-report.md";
    a.click();
    URL.revokeObjectURL(url);
  }

  const tickerItems = aiInsights ? [aiInsights.headline, ...(aiInsights.insights || [])].filter(Boolean) : [];

  return (
    <div className="vb-shell" data-theme={theme}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

        .vb-shell {
          --bg: #0B1220;
          --bg-elev: #101A30;
          --card: #141F38;
          --card-hover: #182647;
          --border: #223154;
          --text: #E7ECF7;
          --text-dim: #8B9AC0;
          --teal: #56E8CC;
          --amber: #F5B942;
          --rose: #FB7185;
          --violet: #A78BFA;
          font-family: 'Inter', sans-serif;
          color: var(--text);
          background: var(--bg);
          border-radius: 14px;
          overflow: hidden;
          min-height: 640px;
          display: flex;
          flex-direction: column;
          box-shadow: 0 1px 0 var(--border) inset;
        }
        .vb-shell[data-theme="light"] {
          --bg: #F2F4FA;
          --bg-elev: #FFFFFF;
          --card: #FFFFFF;
          --card-hover: #F7F9FE;
          --border: #DCE2F0;
          --text: #131B2E;
          --text-dim: #667088;
          --teal: #0F9E86;
          --amber: #B4790A;
          --rose: #D6415A;
          --violet: #7C5CE0;
        }
        .vb-shell * { box-sizing: border-box; }
        .vb-shell ::-webkit-scrollbar { width: 8px; height: 8px; }
        .vb-shell ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 8px; }

        .vb-topbar {
          display: flex; align-items: center; justify-content: space-between;
          padding: 14px 20px; border-bottom: 1px solid var(--border);
          background: var(--bg-elev);
        }
        .vb-brand { display: flex; align-items: center; gap: 10px; }
        .vb-brand-mark {
          width: 30px; height: 30px; border-radius: 8px;
          background: linear-gradient(135deg, var(--teal), var(--violet));
          display: flex; align-items: center; justify-content: center;
          color: #0B1220; flex-shrink: 0;
        }
        .vb-brand-text { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 17px; letter-spacing: -0.02em; }
        .vb-brand-sub { font-size: 10.5px; color: var(--text-dim); font-family: 'JetBrains Mono', monospace; letter-spacing: 0.04em; }
        .vb-top-actions { display: flex; align-items: center; gap: 8px; }
        .vb-icon-btn {
          width: 34px; height: 34px; border-radius: 8px; border: 1px solid var(--border);
          background: var(--card); color: var(--text-dim); display: flex; align-items: center; justify-content: center;
          cursor: pointer; transition: all .15s ease;
        }
        .vb-icon-btn:hover { color: var(--text); background: var(--card-hover); }
        .vb-btn {
          display: flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 8px;
          border: 1px solid var(--border); background: var(--card); color: var(--text);
          font-size: 13px; font-weight: 500; cursor: pointer; transition: all .15s ease; font-family: inherit;
        }
        .vb-btn:hover { background: var(--card-hover); }
        .vb-btn--primary {
          background: linear-gradient(135deg, var(--teal), #33C9AD); color: #06251F; border: none; font-weight: 600;
        }
        .vb-btn--primary:hover { filter: brightness(1.06); }

        .vb-ticker {
          display: flex; align-items: center; gap: 10px; padding: 9px 20px;
          background: var(--bg); border-bottom: 1px solid var(--border);
          font-family: 'JetBrains Mono', monospace; font-size: 12.5px; color: var(--teal);
          overflow: hidden;
        }
        .vb-ticker-dot { width: 7px; height: 7px; border-radius: 50%; background: var(--teal); animation: vb-pulse 1.6s infinite; flex-shrink: 0; }
        .vb-ticker-label { color: var(--text-dim); flex-shrink: 0; letter-spacing: .06em; }
        .vb-ticker-text { animation: vb-fade-up .4s ease; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--text); }
        @keyframes vb-pulse { 0%,100% { opacity: 1; } 50% { opacity: .25; } }
        @keyframes vb-fade-up { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }

        .vb-body { flex: 1; overflow-y: auto; padding: 22px; }

        /* Upload stage */
        .vb-upload-wrap { display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 480px; gap: 18px; }
        .vb-dropzone {
          width: min(560px, 100%); border: 1.5px dashed var(--border); border-radius: 16px;
          padding: 46px 30px; text-align: center; background: var(--card); transition: all .15s ease; cursor: pointer;
        }
        .vb-dropzone.drag { border-color: var(--teal); background: var(--card-hover); }
        .vb-dropzone-icon {
          width: 52px; height: 52px; border-radius: 12px; margin: 0 auto 16px; background: var(--bg-elev);
          border: 1px solid var(--border); display: flex; align-items: center; justify-content: center; color: var(--teal);
        }
        .vb-dropzone h3 { font-family: 'Space Grotesk', sans-serif; font-size: 18px; margin: 0 0 6px; }
        .vb-dropzone p { color: var(--text-dim); font-size: 13.5px; margin: 0 0 20px; }
        .vb-divider { display: flex; align-items: center; gap: 10px; color: var(--text-dim); font-size: 11.5px; width: min(560px,100%); }
        .vb-divider::before, .vb-divider::after { content: ''; flex: 1; height: 1px; background: var(--border); }
        .vb-error { color: var(--rose); font-size: 13px; background: rgba(251,113,133,.1); border: 1px solid rgba(251,113,133,.3); padding: 10px 14px; border-radius: 8px; width: min(560px,100%); }
        .vb-formats { display: flex; gap: 6px; flex-wrap: wrap; justify-content: center; margin-top: 14px; }
        .vb-format-chip { font-family: 'JetBrains Mono', monospace; font-size: 10.5px; color: var(--text-dim); border: 1px solid var(--border); padding: 3px 8px; border-radius: 5px; }

        /* Processing stage */
        .vb-processing { max-width: 460px; margin: 60px auto; }
        .vb-processing h3 { font-family: 'Space Grotesk', sans-serif; text-align: center; margin-bottom: 4px; }
        .vb-processing p.vb-sub { text-align: center; color: var(--text-dim); font-size: 13px; margin-bottom: 28px; }
        .vb-step-row { display: flex; align-items: center; gap: 10px; padding: 8px 0; opacity: 0; animation: vb-fade-up .35s ease forwards; }
        .vb-step-icon { width: 20px; height: 20px; border-radius: 50%; border: 1.5px solid var(--border); display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
        .vb-step-icon.done { border-color: var(--teal); background: var(--teal); color: #06251F; }
        .vb-step-icon.active { border-color: var(--teal); }
        .vb-step-text { font-size: 13.5px; color: var(--text-dim); }
        .vb-step-text.done { color: var(--text); }

        /* Dashboard */
        .vb-kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin-bottom: 20px; }
        .vb-kpi { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 15px 16px; position: relative; }
        .vb-kpi-label { font-size: 11.5px; color: var(--text-dim); margin-bottom: 8px; letter-spacing: .01em; }
        .vb-kpi-value { font-family: 'JetBrains Mono', monospace; font-size: 22px; font-weight: 500; }
        .vb-kpi-tone { position: absolute; top: 14px; right: 14px; }
        .vb-kpi-tone--up { color: var(--teal); }
        .vb-kpi-tone--down { color: var(--rose); }

        .vb-grid-2 { display: grid; grid-template-columns: 1.3fr 1fr; gap: 16px; margin-bottom: 16px; }
        .vb-grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; margin-bottom: 16px; }
        @media (max-width: 860px) { .vb-grid-2, .vb-grid-3 { grid-template-columns: 1fr; } }

        .vb-card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px; }
        .vb-card-title { font-family: 'Space Grotesk', sans-serif; font-size: 13.5px; font-weight: 600; margin-bottom: 4px; display: flex; align-items: center; gap: 7px; }
        .vb-card-sub { font-size: 11.5px; color: var(--text-dim); margin-bottom: 14px; }

        .vb-insight-list { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 10px; }
        .vb-insight-list li { display: flex; gap: 9px; font-size: 13px; line-height: 1.5; color: var(--text); }
        .vb-insight-list svg { flex-shrink: 0; margin-top: 2px; color: var(--amber); }
        .vb-summary { font-size: 13.5px; line-height: 1.6; color: var(--text); background: var(--bg-elev); border: 1px solid var(--border); border-radius: 10px; padding: 12px 14px; margin-bottom: 16px; }
        .vb-rec-list { list-style: none; margin: 12px 0 0; padding: 12px 0 0; border-top: 1px solid var(--border); display: flex; flex-direction: column; gap: 8px; }
        .vb-rec-list li { font-size: 12.5px; color: var(--text-dim); display: flex; gap: 8px; }
        .vb-rec-list li span { color: var(--violet); flex-shrink: 0; }

        table.vb-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
        table.vb-table th { text-align: left; color: var(--text-dim); font-weight: 500; padding: 6px 8px; border-bottom: 1px solid var(--border); font-size: 11px; text-transform: uppercase; letter-spacing: .04em; }
        table.vb-table td { padding: 7px 8px; border-bottom: 1px solid var(--border); }
        table.vb-table tr:last-child td { border-bottom: none; }
        .vb-badge { font-family: 'JetBrains Mono', monospace; font-size: 10px; padding: 2px 7px; border-radius: 5px; text-transform: uppercase; letter-spacing: .03em; }
        .vb-badge--date { background: rgba(167,139,250,.15); color: var(--violet); }
        .vb-badge--numeric { background: rgba(86,232,204,.15); color: var(--teal); }
        .vb-badge--categorical { background: rgba(245,185,66,.15); color: var(--amber); }
        .vb-badge--id { background: rgba(251,113,133,.15); color: var(--rose); }
        .vb-badge--unknown { background: rgba(139,154,192,.15); color: var(--text-dim); }

        .vb-chat { display: flex; flex-direction: column; height: 340px; }
        .vb-chat-msgs { flex: 1; overflow-y: auto; display: flex; flex-direction: column; gap: 10px; padding-right: 4px; margin-bottom: 10px; }
        .vb-msg { max-width: 85%; font-size: 13px; line-height: 1.5; padding: 9px 12px; border-radius: 10px; }
        .vb-msg--user { align-self: flex-end; background: linear-gradient(135deg, var(--teal), #33C9AD); color: #06251F; }
        .vb-msg--ai { align-self: flex-start; background: var(--bg-elev); border: 1px solid var(--border); }
        .vb-chat-empty { color: var(--text-dim); font-size: 12.5px; text-align: center; margin: auto; }
        .vb-chat-form { display: flex; gap: 8px; }
        .vb-chat-input { flex: 1; background: var(--bg-elev); border: 1px solid var(--border); border-radius: 8px; padding: 9px 12px; color: var(--text); font-size: 13px; font-family: inherit; }
        .vb-chat-input:focus { outline: none; border-color: var(--teal); }
        .vb-chat-send { width: 38px; height: 38px; border-radius: 8px; border: none; background: linear-gradient(135deg, var(--teal), #33C9AD); color: #06251F; display: flex; align-items: center; justify-content: center; cursor: pointer; flex-shrink: 0; }
        .vb-chat-send:disabled { opacity: .5; cursor: default; }
        .vb-typing { display: flex; gap: 3px; padding: 4px 0; }
        .vb-typing span { width: 5px; height: 5px; border-radius: 50%; background: var(--text-dim); animation: vb-pulse 1s infinite; }
        .vb-typing span:nth-child(2) { animation-delay: .2s; }
        .vb-typing span:nth-child(3) { animation-delay: .4s; }

        .vb-quick-qs { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 10px; }
        .vb-quick-q { font-size: 11.5px; border: 1px solid var(--border); background: var(--bg-elev); color: var(--text-dim); padding: 5px 10px; border-radius: 20px; cursor: pointer; }
        .vb-quick-q:hover { color: var(--text); border-color: var(--teal); }
      `}</style>

      {/* ---------------- Topbar ---------------- */}
      <div className="vb-topbar">
        <div className="vb-brand">
          <div className="vb-brand-mark"><LayoutGrid size={16} /></div>
          <div>
            <div className="vb-brand-text">VisionBI</div>
            <div className="vb-brand-sub">AI DATA ANALYST</div>
          </div>
        </div>
        <div className="vb-top-actions">
          {stage === "dashboard" && (
            <>
              <button className="vb-btn" onClick={exportSummary}><Download size={14} /> Export Report</button>
              <button className="vb-btn" onClick={resetAll}><RefreshCw size={14} /> New Dataset</button>
            </>
          )}
          <button className="vb-icon-btn" onClick={() => setTheme(theme === "dark" ? "light" : "dark")} title="Toggle theme">
            {theme === "dark" ? <Sun size={16} /> : <Moon size={16} />}
          </button>
        </div>
      </div>

      {/* ---------------- Ticker (signature element) ---------------- */}
      {stage === "dashboard" && tickerItems.length > 0 && (
        <div className="vb-ticker">
          <span className="vb-ticker-dot" />
          <span className="vb-ticker-label">LIVE INSIGHT</span>
          <span className="vb-ticker-text" key={tickerIndex}>{tickerItems[tickerIndex]}</span>
        </div>
      )}

      <div className="vb-body">
        {/* ---------------- Upload stage ---------------- */}
        {stage === "upload" && (
          <div className="vb-upload-wrap">
            {error && <div className="vb-error">{error}</div>}
            <div
              className={`vb-dropzone${dragOver ? " drag" : ""}`}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
              onClick={() => fileInputRef.current?.click()}
            >
              <div className="vb-dropzone-icon"><Upload size={22} /></div>
              <h3>Upload Any Data</h3>
              <p>Drag & drop a file, or click to browse. VisionBI will clean it, profile it, and build a dashboard automatically.</p>
              <button className="vb-btn vb-btn--primary" onClick={(e) => { e.stopPropagation(); fileInputRef.current?.click(); }}>
                <FileSpreadsheet size={14} /> Choose File
              </button>
              <div className="vb-formats">
                {["CSV", "TSV", "XLSX", "XLS"].map((f) => <span key={f} className="vb-format-chip">{f}</span>)}
              </div>
              <input
                ref={fileInputRef}
                type="file"
                accept=".csv,.tsv,.txt,.xlsx,.xls"
                style={{ display: "none" }}
                onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }}
              />
            </div>
            <div className="vb-divider">OR</div>
            <button className="vb-btn" onClick={loadSample}>
              <Sparkles size={14} /> Try Sample Retail Dataset
            </button>
          </div>
        )}

        {/* ---------------- Processing stage ---------------- */}
        {stage === "processing" && (
          <div className="vb-processing">
            <h3>Analyzing {fileName}</h3>
            <p className="vb-sub">VisionBI is running its automated pipeline</p>
            {PIPELINE_STEPS.map((step, i) => (
              <div className="vb-step-row" key={step} style={{ animationDelay: `${i * 0.03}s` }}>
                <div className={`vb-step-icon${i < stepIndex ? " done" : i === stepIndex ? " active" : ""}`}>
                  {i < stepIndex ? <CheckCircle2 size={12} /> : i === stepIndex ? <Loader2 size={12} className="vb-spin" /> : null}
                </div>
                <div className={`vb-step-text${i < stepIndex ? " done" : ""}`}>{step}</div>
              </div>
            ))}
          </div>
        )}

        {/* ---------------- Dashboard stage ---------------- */}
        {stage === "dashboard" && profile && (
          <div>
            <div className="vb-kpi-grid">
              {kpis.map((k, i) => <KPICard kpi={k} key={i} />)}
            </div>

            <div className="vb-grid-2">
              <div className="vb-card">
                <div className="vb-card-title"><TrendingUp size={14} color="var(--teal)" /> {chartMeta.numCol ? titleCase(chartMeta.numCol) : "Value"} Trend</div>
                <div className="vb-card-sub">{trendData.length ? `Aggregated by month${chartMeta.dateCol ? ` · ${titleCase(chartMeta.dateCol)}` : ""}` : "No date column detected for trend analysis"}</div>
                {trendData.length ? (
                  <ResponsiveContainer width="100%" height={220}>
                    <LineChart data={trendData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" />
                      <XAxis dataKey="month" tick={{ fontSize: 11, fill: "var(--text-dim)" }} />
                      <YAxis tick={{ fontSize: 11, fill: "var(--text-dim)" }} />
                      <Tooltip contentStyle={{ background: "var(--bg-elev)", border: "1px solid var(--border)", borderRadius: 8, fontSize: 12 }} />
                      <Line type="monotone" dataKey="value" stroke="var(--teal)" strokeWidth={2.5} dot={{ r: 3 }} />
                    </LineChart>
                  </ResponsiveContainer>
                ) : (
                  <div style={{ color: "var(--text-dim)", fontSize: 13, padding: "40px 0", textAlign: "center" }}>—</div>
                )}
              </div>

              <div className="vb-card">
                <div className="vb-card-title"><Sparkles size={14} color="var(--amber)" /> AI Insights</div>
                <div className="vb-card-sub">Generated by Claude from this dataset's profile</div>
                <ul className="vb-insight-list">
                  {(aiInsights?.insights || []).map((ins, i) => (
                    <li key={i}><Sparkles size={13} />{ins}</li>
                  ))}
                </ul>
              </div>
            </div>

            <div className="vb-grid-3">
              <div className="vb-card">
                <div className="vb-card-title"><LayoutGrid size={14} color="var(--amber)" /> {chartMeta.catCol ? titleCase(chartMeta.catCol) : "Category"} Breakdown</div>
                <div className="vb-card-sub">Top segments by {chartMeta.numCol ? titleCase(chartMeta.numCol) : "count"}</div>
                {catData.length ? (
                  <ResponsiveContainer width="100%" height={200}>
                    <BarChart data={catData} layout="vertical" margin={{ left: 10 }}>
                      <XAxis type="number" tick={{ fontSize: 10, fill: "var(--text-dim)" }} />
                      <YAxis type="category" dataKey="name" width={70} tick={{ fontSize: 10.5, fill: "var(--text-dim)" }} />
                      <Tooltip contentStyle={{ background: "var(--bg-elev)", border: "1px solid var(--border)", borderRadius: 8, fontSize: 12 }} />
                      <Bar dataKey="value" fill="var(--amber)" radius={[0, 4, 4, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                ) : <div style={{ color: "var(--text-dim)", fontSize: 13, padding: "40px 0", textAlign: "center" }}>—</div>}
              </div>

              <div className="vb-card">
                <div className="vb-card-title"><Radio size={14} color="var(--violet)" /> {chartMeta.pieCol ? titleCase(chartMeta.pieCol) : "Distribution"}</div>
                <div className="vb-card-sub">Share of records</div>
                {pieData.length ? (
                  <ResponsiveContainer width="100%" height={200}>
                    <PieChart>
                      <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} label={{ fontSize: 10, fill: "var(--text-dim)" }}>
                        {pieData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                      </Pie>
                      <Tooltip contentStyle={{ background: "var(--bg-elev)", border: "1px solid var(--border)", borderRadius: 8, fontSize: 12 }} />
                    </PieChart>
                  </ResponsiveContainer>
                ) : <div style={{ color: "var(--text-dim)", fontSize: 13, padding: "40px 0", textAlign: "center" }}>—</div>}
              </div>

              <div className="vb-card">
                <div className="vb-card-title"><Table2 size={14} color="var(--teal)" /> Data Profile</div>
                <div className="vb-card-sub">{profile.rowCount.toLocaleString()} rows · {profile.duplicates} duplicates</div>
                <table className="vb-table">
                  <thead><tr><th>Column</th><th>Type</th><th>Missing</th></tr></thead>
                  <tbody>
                    {profile.columns.slice(0, 6).map((c) => (
                      <tr key={c.name}>
                        <td title={c.businessLabel}>{c.name}</td>
                        <td><TypeBadge type={c.type} /></td>
                        <td>{c.missingPct}%</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {profile.columns.length > 6 && (
                  <div style={{ fontSize: 11, color: "var(--text-dim)", marginTop: 8 }}>+{profile.columns.length - 6} more columns</div>
                )}
              </div>
            </div>

            <div className="vb-grid-2">
              <div className="vb-card">
                <div className="vb-card-title"><Sparkles size={14} color="var(--teal)" /> Executive Summary</div>
                <div className="vb-summary">{aiInsights?.summary}</div>
                <div className="vb-card-title" style={{ fontSize: 12.5, marginBottom: 0 }}>Recommendations</div>
                <ul className="vb-rec-list">
                  {(aiInsights?.recommendations || []).map((r, i) => (
                    <li key={i}><span>→</span>{r}</li>
                  ))}
                </ul>
              </div>

              <div className="vb-card">
                <div className="vb-card-title"><Bot size={14} color="var(--violet)" /> Ask VisionBI</div>
                <div className="vb-card-sub">Chat with your data in plain language</div>
                <div className="vb-chat">
                  <div className="vb-chat-msgs">
                    {chatMessages.length === 0 && (
                      <div className="vb-chat-empty">
                        <MessageSquare size={20} style={{ opacity: .4, marginBottom: 6 }} />
                        <div>Ask a question about this dataset</div>
                      </div>
                    )}
                    {chatMessages.map((m, i) => (
                      <div className={`vb-msg vb-msg--${m.role}`} key={i}>{m.text}</div>
                    ))}
                    {chatLoading && (
                      <div className="vb-msg vb-msg--ai vb-typing"><span /><span /><span /></div>
                    )}
                    <div ref={chatEndRef} />
                  </div>
                  {chatMessages.length === 0 && (
                    <div className="vb-quick-qs">
                      {["What's driving the biggest KPI?", "Which segment leads?", "Summarize data quality"].map((q) => (
                        <button className="vb-quick-q" key={q} onClick={() => setChatInput(q)}>{q}</button>
                      ))}
                    </div>
                  )}
                  <form className="vb-chat-form" onSubmit={handleAskAI}>
                    <input
                      className="vb-chat-input"
                      placeholder="e.g. Why did the trend change?"
                      value={chatInput}
                      onChange={(e) => setChatInput(e.target.value)}
                    />
                    <button className="vb-chat-send" type="submit" disabled={chatLoading || !chatInput.trim()}>
                      <Send size={15} />
                    </button>
                  </form>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      <style>{`.vb-spin { animation: vb-spin 1s linear infinite; } @keyframes vb-spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}
