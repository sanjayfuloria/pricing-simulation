import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  isFirebaseConfigured, createGame, updateGameMeta, saveTeamsToGame,
  submitStudentPrice, joinGame, checkGameExists, listenToGameMeta,
  listenToTeams, processQuarterResults, startQuarter as fbStartQuarter,
  getGameTeamIds,
} from "./firebase.js";

const FIREBASE_ENABLED = isFirebaseConfigured();

/* ═══════════════════════════════════════════════════════════════════════════
   PRICING SIMULATION PLATFORM — B2B Marketing, IBS Hyderabad
   Multi-user competitive pricing game with Faculty & Student portals
   Features: Excel roster upload, individual/team play, SD penalty, scoring
   ═══════════════════════════════════════════════════════════════════════════ */

// ─── CONSTANTS ───────────────────────────────────────────────────────────────

// QR Code component — generates QR using a canvas-based algorithm (no external deps)
function QRCode({ text, size = 200 }) {
  const canvasRef = useRef(null);
  useEffect(() => {
    if (!canvasRef.current || !text) return;
    // Use a simple QR encoding via an offscreen image from a public API
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.src = `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${encodeURIComponent(text)}&margin=4`;
    img.onload = () => {
      const ctx = canvasRef.current.getContext("2d");
      ctx.clearRect(0, 0, size, size);
      ctx.drawImage(img, 0, 0, size, size);
    };
    img.onerror = () => {
      // Fallback: draw a placeholder
      const ctx = canvasRef.current.getContext("2d");
      ctx.fillStyle = "#f4f1eb";
      ctx.fillRect(0, 0, size, size);
      ctx.fillStyle = "#1a1a1a";
      ctx.font = "12px sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("QR Code", size/2, size/2 - 6);
      ctx.fillText(text.substring(0, 30), size/2, size/2 + 10);
    };
  }, [text, size]);
  return <canvas ref={canvasRef} width={size} height={size} style={{borderRadius:8,border:"1px solid var(--border)"}} />;
}
const PHASE_CONFIG = {
  1: { name: "Phase 1 — Market Entry", quarters: [1,2,3,4], vc: 3, fc: 500, promo: false, recession: false,
       color: "#1B4332", accent: "#2D6A4F", tag: "ENTRY",
       desc: "Establish presence in a price-sensitive government tiffin market." },
  2: { name: "Phase 2 — Promotions", quarters: [5,6,7,8], vc: 3, fc: 500, promo: true, recession: false,
       color: "#7B2D8E", accent: "#9B59B6", tag: "GROWTH",
       desc: "Promotions unlocked (max 20% of last quarter revenue). Market conditions worsening." },
  3: { name: "Phase 3 — Recession", quarters: [9,10,11,12], vc: 5, fc: 1000, promo: true, recession: true,
       color: "#922B21", accent: "#C0392B", tag: "SURVIVE",
       desc: "Variable cost ₹5, fixed cost ₹1000. Price drops now attract customers." },
};
const INIT_CUSTOMERS = 100, INIT_PRICE = 10, PRICE_STEP = 0.5;

// ─── GAME LOGIC ──────────────────────────────────────────────────────────────
const getPhase = q => q <= 4 ? 1 : q <= 8 ? 2 : 3;

function getRetention(price, lastAIP) {
  if (!lastAIP || lastAIP === 0) return { rate: 0, label: "N/A" };
  const pct = (lastAIP - price) / lastAIP;
  if (pct > 0.2) return { rate: 0.30, label: ">20% below AIP" };
  if (pct >= 0.1) return { rate: 0.15, label: "10-20% below AIP" };
  if (pct > 0) return { rate: 0.10, label: "<10% below AIP" };
  return { rate: 0, label: "≥ AIP (no retention)" };
}

function calcNewCustomers(price, avgComp, phase, promo = 0, lastPrice = null) {
  let nc;
  if (phase === 1) nc = 400 - 40*price + 21*avgComp;
  else if (phase === 2) nc = 400 - 40*price + 21*avgComp + 0.10*promo;
  else {
    const drop = lastPrice !== null ? Math.max(0, lastPrice - price) : 0;
    nc = 400 - 40*price + 21*avgComp + 0.20*promo + 100*drop;
  }
  return Math.max(0, Math.round(nc));
}

function simulateQuarter(price, avgComp, phase, promo, prevData) {
  const cfg = PHASE_CONFIG[phase];
  const lastAIP = prevData ? prevData.avgComp : INIT_PRICE;
  const lastSales = prevData ? prevData.totalSales : INIT_CUSTOMERS;
  const lastPrice = prevData ? prevData.price : INIT_PRICE;
  const lastRev = prevData ? prevData.revenue : INIT_CUSTOMERS * INIT_PRICE;
  const maxPromo = cfg.promo ? Math.floor(lastRev * 0.2) : 0;
  const clampedPromo = Math.min(promo, maxPromo);

  const ret = getRetention(price, lastAIP);
  const retained = Math.round(lastSales * ret.rate);
  const newCust = calcNewCustomers(price, avgComp, phase, clampedPromo, lastPrice);
  const totalSales = retained + newCust;
  const revenue = totalSales * price;
  const profit = totalSales * (price - cfg.vc) - cfg.fc - clampedPromo;

  return { price, avgComp, promo: clampedPromo, maxPromo, retRate: ret.rate, retLabel: ret.label,
           retained, newCust, totalSales, revenue, profit, phase };
}

const fmt = v => v == null || isNaN(v) ? "—" : "₹" + Number(v).toLocaleString("en-IN", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
const fmtD = v => v == null || isNaN(v) ? "—" : "₹" + Number(v).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// ─── EXCEL ROSTER PARSING ──────────────────────────────────────────────────
function parseRosterExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const roster = rows.map(r => ({
          name: (r["Student Name"] || r.StudentName || r.Name || r.name || r.Student || "").toString().trim(),
          team: (r["Team Name"] || r.TeamName || r.Team || r.team || "").toString().trim(),
          seat: (r.Seat || r.seat || r.SeatNo || "").toString().trim(),
          enrol: (r["Enrollment Number"] || r.Enrolno || r.Enrolment || r.enrol || r.EnrolmentNumber || r.Roll || r["Enrolment Number"] || r["Roll Number"] || "").toString().trim(),
          email: (r["Email"] || r["Email ID"] || r["Inst. EMail"] || r.email || "").toString().trim(),
          section: (r["Section"] || r.section || "").toString().trim(),
        })).filter(r => r.name);
        resolve(roster);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function rosterToTeams(roster, allowIndividual = false) {
  const teamMap = {};
  const individuals = [];
  roster.forEach(r => {
    const teamName = r.team.trim();
    if (!teamName || (allowIndividual && teamName.toLowerCase() === "individual")) {
      individuals.push(r);
    } else {
      if (!teamMap[teamName]) teamMap[teamName] = { members: [], section: r.section || "" };
      teamMap[teamName].members.push(r);
      if (r.section) teamMap[teamName].section = r.section;
    }
  });
  let id = 0;
  const teams = Object.entries(teamMap).map(([name, data]) => ({
    id: `team-${id++}`, name,
    section: data.section || "",
    members: data.members.map(m => ({ name: m.name, seat: m.seat, enrol: m.enrol, email: m.email || "" })),
    quarters: [], pendingPrice: null, pendingPromo: null, submitted: false,
    isIndividual: false,
  }));
  individuals.forEach(r => {
    teams.push({
      id: `ind-${id++}`, name: r.name, section: r.section || "",
      members: [{ name: r.name, seat: r.seat, enrol: r.enrol, email: r.email || "" }],
      quarters: [], pendingPrice: null, pendingPromo: null, submitted: false,
      isIndividual: true,
    });
  });
  return teams;
}

// ─── STANDARD DEVIATION PENALTY ────────────────────────────────────────────
function calcStdDev(values) {
  if (values.length < 2) return 0;
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  const variance = values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / values.length;
  return Math.sqrt(variance);
}

function applySDPenalty(teams, quarterIndex, penaltyEnabled = true) {
  // Get all prices for this quarter
  const prices = teams.map(t => t.quarters[quarterIndex]?.price).filter(p => p != null);
  if (prices.length < 3 || !penaltyEnabled) return teams;

  const mean = prices.reduce((a, b) => a + b, 0) / prices.length;
  const sd = calcStdDev(prices);
  if (sd === 0) return teams;

  return teams.map(t => {
    const q = t.quarters[quarterIndex];
    if (!q) return t;
    const zScore = Math.abs(q.price - mean) / sd;
    let penalty = 0;
    if (zScore > 2) penalty = q.profit * 0.20; // 20% profit penalty for extreme outliers
    else if (zScore > 1.5) penalty = q.profit * 0.10; // 10% for moderate outliers
    const updatedQ = { ...q, sdPenalty: Math.round(penalty), zScore: +zScore.toFixed(2),
                        adjustedProfit: Math.round(q.profit - Math.abs(penalty)) };
    const newQuarters = [...t.quarters];
    newQuarters[quarterIndex] = updatedQ;
    return { ...t, quarters: newQuarters };
  });
}

// ─── COMPOSITE SCORING INDEX ───────────────────────────────────────────────
function calcScores(teams) {
  if (!teams.length || !teams[0].quarters.length) return teams.map(t => ({ ...t, score: 0, rank: 0 }));

  // Component metrics for each team
  const metrics = teams.map(t => {
    const qs = t.quarters;
    const totProfit = qs.reduce((s, q) => s + (q.adjustedProfit ?? q.profit), 0);
    const totRevenue = qs.reduce((s, q) => s + q.revenue, 0);
    const totSales = qs.reduce((s, q) => s + q.totalSales, 0);
    // Consistency: lower std dev of profits = more consistent
    const profitValues = qs.map(q => q.adjustedProfit ?? q.profit);
    const profitSD = calcStdDev(profitValues);
    // Penalty count
    const totalPenalties = qs.reduce((s, q) => s + (q.sdPenalty ? 1 : 0), 0);
    // Growth: compare last 4 quarters profit to first 4
    const early = qs.slice(0, 4).reduce((s, q) => s + (q.adjustedProfit ?? q.profit), 0);
    const late = qs.slice(-4).reduce((s, q) => s + (q.adjustedProfit ?? q.profit), 0);
    const growth = early !== 0 ? (late - early) / Math.abs(early) : 0;

    return { team: t, totProfit, totRevenue, totSales, profitSD, totalPenalties, growth };
  });

  // Normalize each metric to 0–100 using min-max
  const normalize = (arr, key, invert = false) => {
    const vals = arr.map(m => m[key]);
    const min = Math.min(...vals), max = Math.max(...vals);
    const range = max - min || 1;
    return arr.map(m => {
      const norm = ((m[key] - min) / range) * 100;
      return invert ? 100 - norm : norm;
    });
  };

  const nProfit = normalize(metrics, "totProfit");
  const nRevenue = normalize(metrics, "totRevenue");
  const nSales = normalize(metrics, "totSales");
  const nConsistency = normalize(metrics, "profitSD", true); // lower SD = better
  const nPenalties = normalize(metrics, "totalPenalties", true); // fewer = better
  const nGrowth = normalize(metrics, "growth");

  // Weighted composite: Profit 35%, Revenue 15%, Sales 15%, Consistency 15%, Penalties 10%, Growth 10%
  const scored = metrics.map((m, i) => {
    const composite = nProfit[i] * 0.35 + nRevenue[i] * 0.15 + nSales[i] * 0.15 +
                      nConsistency[i] * 0.15 + nPenalties[i] * 0.10 + nGrowth[i] * 0.10;
    return { ...m.team, score: +composite.toFixed(1), totProfit: m.totProfit, totRevenue: m.totRevenue,
             totSales: m.totSales, penalties: m.totalPenalties, growth: +(m.growth * 100).toFixed(1) };
  });

  // Sort by score descending and assign ranks
  scored.sort((a, b) => b.score - a.score);
  scored.forEach((t, i) => { t.rank = i + 1; });
  return scored;
}

// ─── ICONS (inline SVG) ─────────────────────────────────────────────────────
const Icon = ({ d, size = 18, color = "currentColor" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d={d}/></svg>
);
const Icons = {
  dashboard: "M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2V9zM9 22V12h6v10",
  play: "M5 3l14 9-14 9V3z",
  users: "M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4-4v2M9 11a4 4 0 100-8 4 4 0 000 8zM23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75",
  chart: "M18 20V10M12 20V4M6 20v-6",
  settings: "M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10zM12 8v4l3 3",
  lock: "M19 11H5a2 2 0 00-2 2v7a2 2 0 002 2h14a2 2 0 002-2v-7a2 2 0 00-2-2zM7 11V7a5 5 0 0110 0v4",
  logout: "M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4M16 17l5-5-5-5M21 12H9",
  check: "M20 6L9 17l-5-5",
  alert: "M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z",
  trophy: "M6 9H4.5a2.5 2.5 0 010-5H6M18 9h1.5a2.5 2.5 0 000-5H18M4 22h16M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20 7 22M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20 17 22M18 2H6v7a6 6 0 1012 0V2z",
  send: "M22 2L11 13M22 2l-7 20-4-9-9-4 20-7z",
  eye: "M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8zM12 9a3 3 0 100 6 3 3 0 000-6z",
  edit: "M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z",
  info: "M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10zM12 16v-4M12 8h.01",
  grid: "M10 3H3v7h7V3zM21 3h-7v7h7V3zM21 14h-7v7h7v-7zM10 14H3v7h7v-7z",
  download: "M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3",
};

// ─── GOOGLE SHEETS ───────────────────────────────────────────────────────────
async function pushToSheets(url, payload) {
  if (!url) return { ok: false, message: "Google Sheets URL not configured" };
  
  // Approach: Use a hidden iframe with a form POST.
  // Google Apps Script redirects on POST, which causes CORS issues with fetch.
  // Form submissions in iframes bypass CORS entirely.
  return new Promise((resolve) => {
    let resolved = false;
    const done = (result) => { if (!resolved) { resolved = true; resolve(result); } };

    // Create a hidden iframe
    const iframeName = "gs_post_" + Date.now();
    const iframe = document.createElement("iframe");
    iframe.name = iframeName;
    iframe.style.cssText = "position:fixed;left:-9999px;top:-9999px;width:1px;height:1px;";
    document.body.appendChild(iframe);

    // Create a form that targets the iframe
    const form = document.createElement("form");
    form.method = "POST";
    form.action = url;
    form.target = iframeName;

    // Add the payload as a hidden field
    const input = document.createElement("textarea");
    input.name = "payload";
    input.value = JSON.stringify(payload);
    input.style.display = "none";
    form.appendChild(input);
    document.body.appendChild(form);

    // The iframe will load after the form submits and Google redirects
    let loadCount = 0;
    iframe.addEventListener("load", () => {
      loadCount++;
      // First load is the blank iframe, second load is after form submission + redirect
      if (loadCount >= 2) {
        done({ ok: true, message: "Data sent to Google Sheets! Check your spreadsheet." });
        setTimeout(() => { 
          try { document.body.removeChild(iframe); } catch(e) {}
          try { document.body.removeChild(form); } catch(e) {}
        }, 2000);
      }
    });

    // Timeout fallback — if iframe doesn't fire load after 12 seconds
    setTimeout(() => {
      done({ ok: true, message: "Request sent. Please check your Google Sheet to confirm data arrived." });
      try { document.body.removeChild(iframe); } catch(e) {}
      try { document.body.removeChild(form); } catch(e) {}
    }, 12000);

    // Submit the form
    form.submit();
  });
}


// ═══════════════════════════════════════════════════════════════════════════
// PRE-LOADED TEAMS — embedded student roster
// ═══════════════════════════════════════════════════════════════════════════
import PRELOADED_TEAMS from "./preloaded_teams.json";

// Build lookup maps for fast access
const PRELOADED_BY_NAME = {};
const PRELOADED_SECTIONS = {};
PRELOADED_TEAMS.forEach(t => {
  PRELOADED_BY_NAME[t.name] = t;
  const sec = t.section || "Other";
  if (!PRELOADED_SECTIONS[sec]) PRELOADED_SECTIONS[sec] = [];
  PRELOADED_SECTIONS[sec].push(t);
});

// ═══════════════════════════════════════════════════════════════════════════
// LOGIN SCREEN — Role selection + Faculty login + Student 2-step join
// ═══════════════════════════════════════════════════════════════════════════
function LoginScreen({ onLogin }) {
  const [role, setRole] = useState(null);
  const [facultyPass, setFacultyPass] = useState("");
  const [error, setError] = useState("");

  // Student wizard: "gameId" → "teamSelect" → "leadVerify"
  const [studentStep, setStudentStep] = useState("gameId");
  const [gameCode, setGameCode] = useState("");
  const [selectedTeamId, setSelectedTeamId] = useState("");
  const [filterSection, setFilterSection] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [verifyName, setVerifyName] = useState("");
  const [isLoading, setIsLoading] = useState(false);

  // The team object selected by the student
  const selectedTeam = PRELOADED_BY_NAME[selectedTeamId] || null;
  const sections = Object.keys(PRELOADED_SECTIONS).sort();
  const teamLead = selectedTeam ? selectedTeam.members[0] : null;

  // Filter teams for display
  const visibleTeams = PRELOADED_TEAMS.filter(t => {
    const matchSec = !filterSection || t.section === filterSection;
    const q = searchQuery.toLowerCase();
    const matchSearch = !q || t.name.toLowerCase().includes(q) ||
      t.members.some(m => m.name.toLowerCase().includes(q) || m.enrol.toLowerCase().includes(q));
    return matchSec && matchSearch;
  });

  const handleFacultyLogin = () => {
    if (facultyPass === "ibs2026") { onLogin({ role: "faculty" }); }
    else { setError("Invalid faculty password"); }
  };

  const handleGameCodeSubmit = async () => {
    const code = gameCode.trim().toUpperCase();
    if (code.length !== 8) { setError("Please enter a valid 8-character simulation code"); return; }
    setError("");
    setStudentStep("teamSelect");
  };

  const handleTeamSelected = async () => {
    if (!selectedTeam) { setError("Please select your team"); return; }
    setError("");
    setIsLoading(true);
    // Check Firebase for any faculty-assigned lead override
    try {
      const res = await fetch(
        `https://pricing-simulation-4ceee-default-rtdb.firebaseio.com/games/${gameCode.trim().toUpperCase()}/teamLeads/${selectedTeam.id}.json?_t=${Date.now()}`,
        { cache: "no-store" }
      );
      if (res.ok) {
        const override = await res.json();
        if (override && override.leadName) {
          // Faculty has reassigned — use the override
          selectedTeam._leadOverride = override.leadName;
          selectedTeam._leadEnrol = override.leadEnrol || "";
        }
      }
    } catch (e) { /* proceed with default lead */ }
    setIsLoading(false);
    setStudentStep("leadVerify");
  };

  const handleLeadVerify = () => {
    if (!selectedTeam) return;
    const currentLead = selectedTeam._leadOverride
      ? { name: selectedTeam._leadOverride, enrol: selectedTeam._leadEnrol }
      : selectedTeam.members[0];
    const entered = verifyName.trim().toLowerCase();
    const leadName = (currentLead.name || "").toLowerCase();
    const leadEnrol = (currentLead.enrol || "").toLowerCase();

    if (!entered) { setError("Please enter your name or enrolment number to verify"); return; }
    if (entered !== leadName && entered !== leadEnrol) {
      setError(`Only the team lead (${currentLead.name}) can enter decisions. Please ask your team lead to log in, or contact faculty to reassign the lead.`);
      return;
    }
    setError("");
    onLogin({
      role: "student",
      teamName: selectedTeam.name,
      section: selectedTeam.section || "",
      members: selectedTeam.members,
      gameCode: gameCode.trim().toUpperCase(),
      isIndividual: selectedTeam.isIndividual || false,
      selectedTeamId: selectedTeam.id,
      isTeamLead: true,
      leadName: currentLead.name,
    });
  };

  // ─── Role selection ──────────────────────────────────────
  if (!role) {
    return (
      <div className="login-wrapper">
        <div className="login-bg-pattern" />
        <div className="login-container">
          <div className="login-brand">
            <div className="brand-icon">🍽️</div>
            <h1>Pricing Simulation</h1>
            <p className="brand-sub">B2B Marketing — MBA 2026 — IBS Hyderabad</p>
            <p className="brand-desc">Compete in a 12-quarter pricing game for government tiffin centres</p>
          </div>
          <div className="login-role-select">
            <button className="role-btn faculty-btn" onClick={() => { setRole("faculty"); setError(""); }}>
              <div className="role-icon">👨‍🏫</div>
              <span className="role-title">Faculty Portal</span>
              <span className="role-desc">Manage games, view all teams, control rounds</span>
            </button>
            <button className="role-btn student-btn" onClick={() => { setRole("student"); setStudentStep("gameId"); setError(""); }}>
              <div className="role-icon">👩‍🎓</div>
              <span className="role-title">Student Portal</span>
              <span className="role-desc">Join a game, set prices, compete with teams</span>
            </button>
          </div>
          <div className="login-footer-text">
            <p>Supports up to 400 concurrent students • Individual or Team play • Real-time leaderboard • Google Sheets export</p>
          </div>
        </div>
      </div>
    );
  }

  // ─── Faculty login ───────────────────────────────────────
  if (role === "faculty") {
    return (
      <div className="login-wrapper">
        <div className="login-bg-pattern" />
        <div className="login-container login-form-container">
          <button className="back-btn" onClick={() => setRole(null)}>← Back</button>
          <div className="login-brand compact">
            <div className="brand-icon small">👨‍🏫</div>
            <h2>Faculty Login</h2>
          </div>
          <div className="login-form">
            <div className="form-field">
              <label>Faculty Password</label>
              <input type="password" value={facultyPass} onChange={e => setFacultyPass(e.target.value)}
                placeholder="Enter password" onKeyDown={e => e.key === "Enter" && handleFacultyLogin()} autoFocus />
            </div>
            {error && <div className="login-error">{error}</div>}
            <button className="login-submit faculty-submit" onClick={handleFacultyLogin}>Enter Control Room</button>
            <p className="login-hint">Default password: ibs2026</p>
          </div>
        </div>
      </div>
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // STUDENT — Step 1: Enter Game Code
  // ═══════════════════════════════════════════════════════════════════
  if (studentStep === "gameId") {
    return (
      <div className="login-wrapper">
        <div className="login-bg-pattern" />
        <div className="login-container login-form-container">
          <button className="back-btn" onClick={() => setRole(null)}>← Back</button>
          <div className="login-brand compact">
            <div className="brand-icon small">🔑</div>
            <h2>Enter Simulation Code</h2>
            <p style={{color:"var(--text-secondary)",fontSize:"0.9rem",marginTop:"0.25rem"}}>
              Your faculty will share an 8-character code to join the simulation.
            </p>
          </div>
          <div className="login-form">
            <div className="form-field">
              <label>Simulation Code *</label>
              <input value={gameCode} onChange={e => setGameCode(e.target.value.toUpperCase())}
                placeholder="e.g. A7BF3X2K" maxLength={8} autoFocus
                style={{fontFamily:"'Courier New',monospace",fontSize:"1.4rem",letterSpacing:"0.15em",textAlign:"center",textTransform:"uppercase"}}
                onKeyDown={e => e.key === "Enter" && handleGameCodeSubmit()} />
              <span style={{fontSize:"0.78rem",color:"var(--text-muted)",marginTop:"0.25rem",display:"block",textAlign:"center"}}>{gameCode.length}/8 characters</span>
            </div>
            {error && <div className="login-error">{error}</div>}
            <button className="login-submit student-submit" onClick={handleGameCodeSubmit}
              disabled={gameCode.trim().length !== 8}>
              Join Game →
            </button>
          </div>
        </div>
      </div>
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // STUDENT — Step 2: Select Team (pre-loaded, searchable)
  // ═══════════════════════════════════════════════════════════════════
  return (
    <div className="login-wrapper">
      <div className="login-bg-pattern" />
      <div className="login-container login-form-container" style={{maxWidth:640}}>
        <button className="back-btn" onClick={() => { setStudentStep("gameId"); setSelectedTeamId(""); setSearchQuery(""); setFilterSection(""); }}>← Back</button>
        <div className="login-brand compact">
          <div className="brand-icon small">👥</div>
          <h2>Find &amp; Join Your Team</h2>
          <div className="game-id-display" style={{marginTop:"0.5rem",padding:"0.4rem 0.75rem"}}>
            <span className="gid-label">Game Code</span>
            <span className="gid-value" style={{fontSize:"1.1rem"}}>{gameCode.toUpperCase()}</span>
          </div>
        </div>
        <div className="login-form">
          {/* Search + Section filter */}
          <div style={{display:"flex",gap:"0.5rem",marginBottom:"0.5rem"}}>
            <div className="form-field" style={{flex:2,marginBottom:0}}>
              <input value={searchQuery} onChange={e => { setSearchQuery(e.target.value); setSelectedTeamId(""); }}
                placeholder="🔍  Search by team name, student name or enrolment..." autoFocus />
            </div>
            <div className="form-field" style={{flex:1,marginBottom:0}}>
              <select value={filterSection} onChange={e => { setFilterSection(e.target.value); setSelectedTeamId(""); }}>
                <option value="">All Sections</option>
                {sections.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </div>

          {/* Team list */}
          <div className="preloaded-team-list">
            {visibleTeams.length === 0 && (
              <div style={{padding:"1rem",textAlign:"center",color:"var(--text-muted)"}}>No teams found. Try a different search.</div>
            )}
            {visibleTeams.map(team => {
              const isSelected = selectedTeamId === team.name;
              return (
                <div key={team.id}
                  className={`preloaded-team-row ${isSelected ? "selected" : ""}`}
                  onClick={() => setSelectedTeamId(team.name)}>
                  <div className="pt-left">
                    <span className="pt-name">{team.name}</span>
                    <span className="pt-section">{team.section}</span>
                  </div>
                  <div className="pt-members">
                    {team.members.map((m,i) => (
                      <span key={i} className="pt-member">{m.name}</span>
                    ))}
                  </div>
                  {isSelected && <span className="pt-check">✓</span>}
                </div>
              );
            })}
          </div>

          {/* Selected team confirmation */}
          {selectedTeam && (
            <div className="selected-team-confirm">
              <div className="stc-header">
                <span className="stc-icon">🤝</span>
                <div>
                  <div className="stc-name">{selectedTeam.name}</div>
                  <div className="stc-section">{selectedTeam.section} • {selectedTeam.members.length} members</div>
                </div>
              </div>
              <div className="stc-members">
                {selectedTeam.members.map((m, i) => (
                  <div key={i} className="stc-member">
                    <span className="stc-num">{i+1}</span>
                    <span className="stc-mname">{m.name}</span>
                    <span className="stc-enrol">{m.enrol}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {error && <div className="login-error">{error}</div>}
          <button className="login-submit student-submit" onClick={handleTeamSelected}
            disabled={!selectedTeam || isLoading} style={{marginTop:"0.75rem"}}>
            {isLoading ? "Checking..." : selectedTeam ? `Proceed with ${selectedTeam.name} →` : "Select Your Team to Continue"}
          </button>
          <p className="login-hint" style={{marginTop:"0.5rem"}}>
            Can't find your team? Ask your faculty or look under a different section filter.
          </p>
        </div>
      </div>
    </div>
  );

  // ═══════════════════════════════════════════════════════════════════
  // STUDENT — Step 3: Team Lead Verification
  // ═══════════════════════════════════════════════════════════════════
  const currentLead = selectedTeam
    ? (selectedTeam._leadOverride
        ? { name: selectedTeam._leadOverride, enrol: selectedTeam._leadEnrol }
        : selectedTeam.members[0])
    : null;

  return (
    <div className="login-wrapper">
      <div className="login-bg-pattern" />
      <div className="login-container login-form-container" style={{maxWidth:520}}>
        <button className="back-btn" onClick={() => { setStudentStep("teamSelect"); setVerifyName(""); setError(""); }}>← Back</button>
        <div className="login-brand compact">
          <div className="brand-icon small">🔐</div>
          <h2>Team Lead Verification</h2>
          <div className="game-id-display" style={{marginTop:"0.5rem",padding:"0.4rem 0.75rem"}}>
            <span className="gid-label">{selectedTeam?.name}</span>
            <span className="gid-value" style={{fontSize:"1rem"}}>{selectedTeam?.section}</span>
          </div>
        </div>
        <div className="login-form">
          {/* Team info */}
          <div className="lead-verify-card">
            <div className="lv-header">
              <span className="lv-crown">👑</span>
              <div>
                <div className="lv-title">Team Lead</div>
                <div className="lv-lead-name">{currentLead?.name}</div>
                <div className="lv-lead-enrol">{currentLead?.enrol}</div>
              </div>
            </div>
            <p className="lv-desc">
              Only the team lead can submit pricing decisions. Other members can view but not submit.
            </p>
            <div className="lv-all-members">
              <div className="lv-members-label">All team members:</div>
              {selectedTeam?.members.map((m, i) => (
                <div key={i} className={`lv-member-row ${i === 0 ? "lead" : ""}`}>
                  {i === 0 ? "👑" : <span className="lv-num">{i+1}</span>}
                  <span className="lv-mname">{m.name}</span>
                  <span className="lv-menrol">{m.enrol}</span>
                </div>
              ))}
            </div>
          </div>

          <div className="form-field" style={{marginTop:"1rem"}}>
            <label>Confirm you are the team lead — enter your name or enrolment number *</label>
            <input value={verifyName} onChange={e => setVerifyName(e.target.value)}
              placeholder={`Enter: ${currentLead?.name || "team lead name"}`}
              onKeyDown={e => e.key === "Enter" && handleLeadVerify()} autoFocus />
          </div>
          {error && <div className="login-error">{error}</div>}
          <button className="login-submit student-submit" onClick={handleLeadVerify}
            disabled={!verifyName.trim()}>
            Enter as Team Lead →
          </button>
          <p className="login-hint">
            Not the team lead? Only the team lead can submit decisions. Faculty can reassign the lead from the Faculty Portal if needed.
          </p>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// FACULTY DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════
// Generate 8-character alphanumeric Game ID
function generateGameId() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // no I/O/0/1 to avoid confusion
  let id = "";
  for (let i = 0; i < 8; i++) id += chars[Math.floor(Math.random() * chars.length)];
  return id;
}

// ─── Late Submission Panel (Faculty) ─────────────────────────────────────────
function LateSubmissionPanel({ gameId, teams, currentQuarter, firebaseRest }) {
  const [latePenalty, setLatePenalty] = useState(10); // default 10% penalty
  const [grantedTeams, setGrantedTeams] = useState({}); // teamId → true
  const [status, setStatus] = useState("");

  const unsubmitted = teams.filter(t => !t.submitted);

  const grantLate = async (team) => {
    const teamKey = team.id;
    try {
      await fetch(`${firebaseRest}/games/${gameId}/meta/lateAllowed/${teamKey}.json`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(true),
      });
      await fetch(`${firebaseRest}/games/${gameId}/meta/latePenalty.json`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(latePenalty),
      });
      setGrantedTeams(prev => ({ ...prev, [teamKey]: true }));
      setStatus(`Late submission granted to ${team.name} with ${latePenalty}% profit penalty.`);
    } catch (e) {
      setStatus("Error granting late submission. Try again.");
    }
  };

  const grantAll = async () => {
    for (const team of unsubmitted) await grantLate(team);
    setStatus(`Late submission granted to all ${unsubmitted.length} pending teams with ${latePenalty}% penalty.`);
  };

  if (unsubmitted.length === 0) return null;

  return (
    <div className="late-panel">
      <div className="late-panel-header">
        <span className="late-panel-icon">⏰</span>
        <div>
          <strong>Late Submission Override</strong>
          <p>{unsubmitted.length} team{unsubmitted.length > 1 ? "s have" : " has"} not submitted yet.</p>
        </div>
      </div>
      <div className="late-penalty-row">
        <label>Penalty for late submission:</label>
        <div className="penalty-input-row">
          {[5,10,15,20,25].map(p => (
            <button key={p} className={`penalty-chip ${latePenalty === p ? "active" : ""}`}
              onClick={() => setLatePenalty(p)}>{p}%</button>
          ))}
          <input type="number" min={1} max={50} value={latePenalty}
            onChange={e => setLatePenalty(Math.min(50, Math.max(1, parseInt(e.target.value)||10)))}
            style={{width:52,textAlign:"center"}} />
        </div>
      </div>
      <div className="late-teams-list">
        {unsubmitted.map(team => (
          <div key={team.id} className="late-team-row">
            <span className="late-team-name">{team.name}</span>
            {grantedTeams[team.id]
              ? <span className="late-granted">✓ Granted ({latePenalty}% penalty)</span>
              : <button className="btn-push" style={{fontSize:"0.75rem",padding:"0.2rem 0.6rem"}} onClick={() => grantLate(team)}>
                  Allow Late
                </button>}
          </div>
        ))}
      </div>
      {unsubmitted.length > 1 && !Object.keys(grantedTeams).length && (
        <button className="btn-push" style={{width:"100%",marginTop:"0.5rem",background:"var(--amber-600)"}}
          onClick={grantAll}>Allow All {unsubmitted.length} Pending Teams (Late)</button>
      )}
      {status && <p className="late-status">{status}</p>}
    </div>
  );
}

const FIREBASE_REST_BASE = "https://pricing-simulation-4ceee-default-rtdb.firebaseio.com";

function FacultyDashboard({ onLogout }) {
  // Setup wizard: "pregame" → "teams" → "playing"
  const [setupStep, setSetupStep] = useState("pregame"); // "pregame" | "teams" | "playing"
  const [tab, setTab] = useState("control");
  const [gameState, setGameState] = useState({
    currentQuarter: 0,
    status: "setup", // "setup" | "active" | "waiting" | "finished"
    aipHistory: [],
    teams: [],
    sheetsUrl: "",
    gameId: generateGameId(),
    sdPenaltyEnabled: true, allowIndividualPlay: true,
    quarterStarted: false,
  });
  const [aipInput, setAipInput] = useState(INIT_PRICE);
  const [pushStatus, setPushStatus] = useState(null);

  // Persist Sheets URL across sessions
  useEffect(() => {
    const saved = localStorage.getItem("pricing-sim-sheets-url");
    if (saved) setGameState(p => ({ ...p, sheetsUrl: saved }));
  }, []);

  const updateSheetsUrl = (url) => {
    setGameState(p => ({ ...p, sheetsUrl: url }));
    try { localStorage.setItem("pricing-sim-sheets-url", url); } catch(e) {}
  };
  const [rosterStatus, setRosterStatus] = useState(null);
  const [teamLeads, setTeamLeads] = useState({}); // { teamId: { leadName, leadEnrol } }
  const [changingLeadFor, setChangingLeadFor] = useState(null); // teamId being edited
  const [newLeadIndex, setNewLeadIndex] = useState(0); // member index in team
  const fileInputRef = useRef(null);
  const attendanceInputRef = useRef(null);
  const [loadOldGameId, setLoadOldGameId] = useState("");
  const [attendanceList, setAttendanceList] = useState([]); // [{name, enrol, section}]
  const [teamSize, setTeamSize] = useState(4); // for auto-assign

  const currentPhase = getPhase(Math.max(gameState.currentQuarter, 1));
  const cfg = PHASE_CONFIG[currentPhase];

  const handleRosterUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setRosterStatus({ type: "loading", message: "Parsing Excel file..." });
    try {
      const roster = await parseRosterExcel(file);
      if (roster.length === 0) { setRosterStatus({ type: "error", message: "No valid rows found. Ensure columns: Name/StudentName, Team/TeamName" }); return; }
      if (roster.length > 400) { setRosterStatus({ type: "error", message: `Found ${roster.length} students. Maximum is 400.` }); return; }
      const teams = rosterToTeams(roster, gameState.allowIndividualPlay);
      setGameState(p => ({ ...p, teams }));
      const indCount = teams.filter(t => t.isIndividual).length;
      const teamCount = teams.filter(t => !t.isIndividual).length;
      setRosterStatus({ type: "ok", message: `Loaded ${roster.length} students → ${teamCount} teams + ${indCount} individuals` });
    } catch (err) {
      setRosterStatus({ type: "error", message: "Failed to parse: " + err.message });
    }
    e.target.value = "";
  };

  // Handle attendance sheet upload (student list without teams)
  const handleAttendanceUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setRosterStatus({ type: "loading", message: "Parsing attendance file..." });
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      const students = rows.map(r => {
        const name = r["Student Name"] || r["StudentName"] || r["Name"] || r["name"] || r["STUDENT NAME"] || "";
        const enrol = String(r["Enrollment Number"] || r["Enrollment"] || r["EnrollmentNumber"] || r["Enrolment"] || r["enrol"] || r["ENROLLMENT NUMBER"] || "");
        const section = r["Section"] || r["section"] || r["SECTION"] || "";
        return { name: name.toString().trim(), enrol: enrol.trim(), section: section.toString().trim() };
      }).filter(s => s.name);
      if (students.length === 0) { setRosterStatus({ type: "error", message: "No students found. Ensure column: Student Name" }); return; }
      if (students.length > 400) { setRosterStatus({ type: "error", message: `Found ${students.length} students. Maximum is 400.` }); return; }
      setAttendanceList(students);
      setRosterStatus({ type: "ok", message: `Loaded ${students.length} students from attendance sheet. Now assign them to teams below.` });
    } catch (err) {
      setRosterStatus({ type: "error", message: "Failed to parse: " + err.message });
    }
    e.target.value = "";
  };

  // Auto-assign students from attendance list into teams
  const autoAssignTeams = () => {
    if (attendanceList.length === 0) return;
    const size = Math.max(1, Math.min(10, teamSize));
    const teams = [];
    for (let i = 0; i < attendanceList.length; i += size) {
      const chunk = attendanceList.slice(i, i + size);
      const teamNum = Math.floor(i / size) + 1;
      teams.push({
        id: `team-${teamNum}`,
        name: `Team ${teamNum}`,
        section: chunk[0]?.section || "",
        members: chunk.map(s => ({ name: s.name, enrol: s.enrol, seat: "" })),
        quarters: [], pendingPrice: null, pendingPromo: null, submitted: false,
        isIndividual: chunk.length === 1 && size === 1,
      });
    }
    setGameState(p => ({ ...p, teams }));
    const indCount = teams.filter(t => t.isIndividual).length;
    const teamCount = teams.filter(t => !t.isIndividual).length;
    setRosterStatus({ type: "ok", message: `Auto-assigned ${attendanceList.length} students → ${teamCount} teams of ~${size} (${indCount} individuals)` });
  };

  const FIREBASE_REST = "https://pricing-simulation-4ceee-default-rtdb.firebaseio.com";

  // Faculty starts the current quarter (enables student timers)
  function startCurrentQuarter() {
    setGameState(prev => ({ ...prev, quarterStarted: true, status: "active" }));
    // Write to localStorage for same-browser tabs
    try {
      localStorage.setItem("pricing-sim-" + gameState.gameId, JSON.stringify({
        type: "QUARTER_STARTED", quarter: gameState.currentQuarter, aip: aipInput,
        gameId: gameState.gameId, timestamp: Date.now(),
      }));
    } catch(e) {}
    // Write to Firebase REST API for cross-device
    fetch(FIREBASE_REST + "/games/" + gameState.gameId + "/meta.json", {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ quarterStarted: true, status: "active", aipInput: aipInput, currentQuarter: gameState.currentQuarter }),
    }).then(r => console.log("Firebase startQuarter:", r.status)).catch(e => console.error("Firebase error:", e));
  }

  // Faculty processes the current quarter and advances
  async function processQuarterAndAdvance() {
    // Fetch late penalty setting from Firebase before processing
    let latePenaltyPct = 0;
    let lateTeamIds = {};
    try {
      const metaRes = await fetch(FIREBASE_REST + "/games/" + gameState.gameId + "/meta.json?_t=" + Date.now(), { cache: "no-store" });
      if (metaRes.ok) {
        const meta = await metaRes.json();
        latePenaltyPct = meta?.latePenalty || 0;
        lateTeamIds = meta?.lateAllowed || {};
      }
    } catch(e) {}

    setGameState(prev => {
      const qNum = prev.currentQuarter;
      const newAipHistory = [...prev.aipHistory, { quarter: qNum, aip: aipInput }];
      let updatedTeams = prev.teams.map(team => {
        const prevData = team.quarters.length > 0 ? team.quarters[team.quarters.length - 1] : null;
        const phase = getPhase(qNum);
        const result = simulateQuarter(team.pendingPrice || INIT_PRICE, aipInput, phase, team.pendingPromo || 0, prevData);
        const isLate = lateTeamIds[team.id] && !team.submitted;
        const penaltyFactor = isLate && latePenaltyPct > 0 ? (1 - latePenaltyPct / 100) : 1;
        const adjustedResult = isLate && latePenaltyPct > 0
          ? { ...result, profit: Math.round(result.profit * penaltyFactor), latePenaltyApplied: latePenaltyPct, adjustedProfit: Math.round(result.profit * penaltyFactor) }
          : result;
        return { ...team, quarters: [...team.quarters, adjustedResult], pendingPrice: null, pendingPromo: null, submitted: false };
      });
      if (prev.sdPenaltyEnabled && updatedTeams.length > 0) {
        updatedTeams = applySDPenalty(updatedTeams, updatedTeams[0].quarters.length - 1, true);
      }
      const isFinished = qNum >= 12;
      try {
        localStorage.setItem("pricing-sim-" + prev.gameId, JSON.stringify({
          type: "QUARTER_PROCESSED", quarter: qNum, nextQuarter: qNum + 1,
          gameId: prev.gameId, isFinished, timestamp: Date.now(),
        }));
      } catch(e) {}

      // Auto-push to Google Sheets after every quarter (if URL is set)
      if (prev.sheetsUrl) {
        const scored = calcScores(updatedTeams);
        const payload = {
          teams: scored.map(t => ({
            name: t.name, section: t.section, members: t.members,
            isIndividual: t.isIndividual || false,
            rank: t.rank, score: t.score, penalties: t.penalties, growth: t.growth,
            totProfit: t.totProfit, totRevenue: t.totRevenue, totSales: t.totSales,
            quarters: t.quarters.map((q, i) => ({ quarter: i + 1, ...q }))
          })),
          aipHistory: newAipHistory,
          sdPenaltyEnabled: prev.sdPenaltyEnabled,
          gameId: prev.gameId,
          autoSync: true,
          syncedAfterQuarter: qNum,
          timestamp: new Date().toISOString()
        };
        pushToSheets(prev.sheetsUrl, payload)
          .then(r => { if (r.ok) setPushStatus({ ok: true, message: `✓ Q${qNum} data auto-saved to Google Sheets` }); })
          .catch(() => {});
      }

      // Clear lateAllowed in Firebase and advance
      fetch(FIREBASE_REST + "/games/" + prev.gameId + "/meta.json", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          currentQuarter: qNum + 1, status: isFinished ? "finished" : "waiting",
          quarterStarted: false, aipHistory: newAipHistory,
          lateAllowed: null, latePenalty: null,
        }),
      }).catch(e => console.error("Firebase processQuarter error:", e));

      return { ...prev, currentQuarter: qNum + 1, aipHistory: newAipHistory,
               teams: updatedTeams, status: isFinished ? "finished" : "waiting",
               quarterStarted: false };
    });
  }

  // Listen to Firebase for real-time team submissions
  useEffect(() => {
    if (!FIREBASE_ENABLED || setupStep !== "playing") return;
    const unsub = listenToTeams(gameState.gameId, (fbTeams) => {
      setGameState(prev => {
        // Merge Firebase team data (submissions) with local state
        const merged = prev.teams.map(t => {
          const fbTeam = fbTeams.find(ft => ft.id === t.id);
          if (fbTeam && fbTeam.submitted && !t.submitted) {
            return { ...t, pendingPrice: fbTeam.pendingPrice, pendingPromo: fbTeam.pendingPromo, submitted: true };
          }
          return t;
        });
        return { ...prev, teams: merged };
      });
    });
    return unsub;
  }, [gameState.gameId, setupStep]);

  const scoredTeams = useMemo(() => {
    try { return calcScores(gameState.teams); }
    catch(e) { console.error("calcScores error:", e); return gameState.teams.map((t,i) => ({...t, score:0, rank:i+1, totProfit:0, totRevenue:0, totSales:0, penalties:0, growth:0})); }
  }, [gameState.teams]);

  const submittedCount = gameState.teams.filter(t => t.submitted).length;

  // Calculate average price from all teams that have submitted this quarter
  const submittedPrices = gameState.teams.filter(t => t.submitted && t.pendingPrice != null).map(t => t.pendingPrice);
  const calculatedAIP = submittedPrices.length > 0
    ? +(submittedPrices.reduce((a, b) => a + b, 0) / submittedPrices.length).toFixed(2)
    : null;
  const priceMin = submittedPrices.length > 0 ? Math.min(...submittedPrices) : null;
  const priceMax = submittedPrices.length > 0 ? Math.max(...submittedPrices) : null;
  const priceSD = submittedPrices.length >= 2 ? +calcStdDev(submittedPrices).toFixed(2) : null;

  // Auto-update AIP input when calculated AIP changes
  useEffect(() => {
    if (calculatedAIP !== null && gameState.quarterStarted) {
      setAipInput(calculatedAIP);
    }
  }, [calculatedAIP, gameState.quarterStarted]);

  const tabs = [
    { id: "control", label: "Game Control", icon: Icons.settings },
    { id: "leaderboard", label: "Leaderboard", icon: Icons.trophy },
    { id: "scores", label: "Scores & Ranks", icon: Icons.eye },
    { id: "teams", label: "All Teams", icon: Icons.users },
    { id: "analytics", label: "Analytics", icon: Icons.chart },
    { id: "sheets", label: "Google Sheets", icon: Icons.grid },
  ];

  // ═══════════════════════════════════════════════════════════════════
  // STEP 1: PRE-GAME SETUP — Generate code or load old game
  // ═══════════════════════════════════════════════════════════════════
  if (setupStep === "pregame") {
    const quickStart = async () => {
      // Load all 48 pre-loaded teams directly and skip team setup wizard
      const teams = PRELOADED_TEAMS.map(t => ({
        ...t,
        quarters: [], pendingPrice: null, pendingPromo: null, submitted: false,
      }));
      setGameState(p => ({ ...p, currentQuarter: 1, status: "waiting", quarterStarted: false, teams }));
      setSetupStep("playing"); setTab("control");
      // Write game to Firebase
      try {
        await fetch(`${FIREBASE_REST_BASE}/games/${gameState.gameId}/meta.json`, {
          method: "PUT", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ status: "waiting", currentQuarter: 1, quarterStarted: false,
            sdPenaltyEnabled: gameState.sdPenaltyEnabled, aipHistory: [], aipInput: 10,
            createdAt: new Date().toISOString() }),
        });
        const teamsObj = {};
        teams.forEach(t => { teamsObj[t.id] = { name: t.name, section: t.section || "", members: t.members || [], isIndividual: false }; });
        await fetch(`${FIREBASE_REST_BASE}/games/${gameState.gameId}/teams.json`, {
          method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(teamsObj),
        });
      } catch(e) { console.error("Firebase quick start error:", e); }
    };

    return (
      <div className="login-wrapper">
        <div className="login-bg-pattern" />
        <div className="login-container login-form-container" style={{maxWidth:580}}>
          <div className="login-brand compact">
            <div className="brand-icon small">👨‍🏫</div>
            <h2>Faculty Pre-Game Setup</h2>
            <p style={{color:"var(--text-secondary)",fontSize:"0.9rem",marginTop:"0.25rem"}}>
              Start a new simulation or analyse a previous game.
            </p>
          </div>
          <div className="login-form">

            {/* ── QUICK START (primary action) ── */}
            <div className="setup-option-card" style={{borderColor:"var(--green-700)",borderWidth:2,background:"var(--green-50)"}}>
              <h3 style={{color:"var(--green-800)"}}>⚡ Quick Start — Teams Already Loaded</h3>
              <p style={{fontSize:"0.88rem",color:"var(--text-secondary)",marginBottom:"0.75rem"}}>
                All 48 teams ({PRELOADED_TEAMS.length} students) are pre-loaded. Skip team setup and go straight to the control panel.
              </p>
              <div className="game-id-display">
                <span className="gid-label">Game Code</span>
                <span className="gid-value">{gameState.gameId}</span>
                <button type="button" className="btn-push" style={{marginTop:"0.5rem",background:"var(--text-muted)",padding:"0.3rem 0.8rem",fontSize:"0.78rem"}}
                  onClick={() => setGameState(p => ({...p, gameId: generateGameId()}))}>
                  Regenerate Code
                </button>
              </div>
              <div className="qr-section" style={{marginTop:"0.75rem"}}>
                <QRCode text={typeof window !== "undefined" ? window.location.origin : "https://pricing-simulation-new.vercel.app"} size={160} />
                <div className="qr-info">
                  <span className="qr-label">Scan to Join</span>
                  <span className="qr-url">{typeof window !== "undefined" ? window.location.origin : "pricing-simulation-new.vercel.app"}</span>
                  <p className="qr-hint">Students scan this QR code, then enter the Game Code above.</p>
                </div>
              </div>
              <div className="settings-toggles" style={{marginTop:"0.75rem"}}>
                <label className="toggle-row">
                  <input type="checkbox" checked={gameState.sdPenaltyEnabled}
                    onChange={e => setGameState(p => ({...p, sdPenaltyEnabled: e.target.checked}))} />
                  <span>Enable SD Penalty</span>
                </label>
              </div>
              <button className="login-submit faculty-submit" style={{marginTop:"0.75rem",background:"var(--green-700)"}} onClick={quickStart}>
                ⚡ Enter Control Panel →
              </button>
            </div>

            {/* ── MANUAL TEAM SETUP ── */}
            <div className="setup-option-card">
              <h3>🔧 Manual Team Setup</h3>
              <p style={{fontSize:"0.88rem",color:"var(--text-secondary)"}}>
                Upload a new roster, re-assign teams, or customise before starting.
              </p>
              <button className="btn-push" style={{marginTop:"0.5rem",width:"100%"}} onClick={() => setSetupStep("teams")}>
                Go to Team Setup →
              </button>
            </div>

            {/* ── ANALYSE OLDER GAME ── */}
            <div className="setup-option-card" style={{borderColor:"var(--border)",background:"var(--bg-alt)"}}>
              <h3>📊 Analyse an Older Game</h3>
              <p style={{fontSize:"0.88rem",color:"var(--text-secondary)",marginBottom:"0.5rem"}}>
                Enter a previous game code to view its dashboard and data.
              </p>
              <div style={{display:"flex",gap:"0.5rem"}}>
                <input value={loadOldGameId} onChange={e => setLoadOldGameId(e.target.value.toUpperCase())}
                  placeholder="e.g. A7BF3X2K" maxLength={8} style={{flex:1,fontFamily:"'Courier New',monospace",letterSpacing:"0.1em",textTransform:"uppercase"}} />
                <button type="button" className="btn-push" disabled={loadOldGameId.length < 8}
                  onClick={() => { setGameState(p => ({...p, gameId: loadOldGameId, teams: PRELOADED_TEAMS.map(t=>({...t,quarters:[],pendingPrice:null,pendingPromo:null,submitted:false}))})); setSetupStep("playing"); setTab("analytics"); }}>
                  Load
                </button>
              </div>
            </div>

            <button type="button" className="back-btn" onClick={onLogout} style={{marginTop:"1rem"}}>← Back to Login</button>
          </div>
        </div>
      </div>
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // STEP 2: TEAM SETUP — Templates, Upload, Attendance, Auto-assign
  // ═══════════════════════════════════════════════════════════════════
  if (setupStep === "teams") {
    return (
      <div className="login-wrapper">
        <div className="login-bg-pattern" />
        <div className="login-container login-form-container" style={{maxWidth:720}}>
          <button className="back-btn" onClick={() => setSetupStep("pregame")}>← Back to Setup</button>
          <div className="login-brand compact">
            <div className="brand-icon small">👥</div>
            <h2>Team Assignment</h2>
            <div className="game-id-display" style={{marginTop:"0.75rem",padding:"0.5rem 0.75rem"}}>
              <span className="gid-label">Game Code</span>
              <span className="gid-value" style={{fontSize:"1.3rem"}}>{gameState.gameId}</span>
            </div>
          </div>
          <div className="login-form">

            {/* ── DOWNLOAD TEMPLATES ── */}
            <div className="setup-option-card">
              <h3>📥 Download Templates</h3>
              <p style={{fontSize:"0.85rem",color:"var(--text-secondary)",marginBottom:"0.75rem"}}>
                Download a pre-formatted Excel template, fill it with your student data, and upload below.
              </p>
              <div style={{display:"flex",gap:"0.5rem",flexWrap:"wrap"}}>
                <a href="/team-roster-template.xlsx" download className="btn-push" style={{textDecoration:"none",display:"inline-flex",alignItems:"center",gap:"0.3rem"}}>
                  <Icon d={Icons.download} size={14} /> Team Roster Template
                </a>
                <a href="/attendance-template.xlsx" download className="btn-push" style={{textDecoration:"none",display:"inline-flex",alignItems:"center",gap:"0.3rem",background:"var(--text-muted)"}}>
                  <Icon d={Icons.download} size={14} /> Attendance Sheet Template
                </a>
              </div>
            </div>

            {/* ── OPTION A: UPLOAD TEAM ROSTER ── */}
            <div className="setup-option-card">
              <h3>Option A — Upload Team Roster</h3>
              <p style={{fontSize:"0.85rem",color:"var(--text-secondary)"}}>
                Upload Excel with columns: <strong>Team Name</strong>, <strong>Student Name</strong>, <strong>Enrollment Number</strong>, <strong>Section</strong>.
                Students with the same team name are grouped together. Leave Team Name blank for individual players.
              </p>
              <div className="roster-upload-row" style={{marginTop:"0.75rem"}}>
                <input type="file" ref={fileInputRef} accept=".xlsx,.xls,.csv" onChange={handleRosterUpload} style={{display:"none"}} />
                <button type="button" className="btn-push" onClick={() => fileInputRef.current?.click()}>
                  <Icon d={Icons.download} size={16} /> Upload Roster File
                </button>
                {gameState.teams.length > 0 && <span className="roster-count">{gameState.teams.length} teams loaded</span>}
              </div>
            </div>

            {/* ── OPTION B: UPLOAD ATTENDANCE + AUTO-ASSIGN ── */}
            <div className="setup-option-card">
              <h3>Option B — Upload Attendance & Auto-Assign Teams</h3>
              <p style={{fontSize:"0.85rem",color:"var(--text-secondary)"}}>
                Upload a student list (names + enrollment numbers), then automatically group them into equally-sized teams.
              </p>
              <div className="roster-upload-row" style={{marginTop:"0.75rem"}}>
                <input type="file" ref={attendanceInputRef} accept=".xlsx,.xls,.csv" onChange={handleAttendanceUpload} style={{display:"none"}} />
                <button type="button" className="btn-push" onClick={() => attendanceInputRef.current?.click()}>
                  <Icon d={Icons.download} size={16} /> Upload Attendance Sheet
                </button>
                {attendanceList.length > 0 && <span className="roster-count">{attendanceList.length} students loaded</span>}
              </div>
              {attendanceList.length > 0 && (
                <div style={{marginTop:"0.75rem",padding:"0.75rem",background:"var(--bg-alt)",borderRadius:"var(--radius)",border:"1px solid var(--border)"}}>
                  <div style={{display:"flex",gap:"0.5rem",alignItems:"center",flexWrap:"wrap"}}>
                    <label style={{fontSize:"0.85rem",fontWeight:600,whiteSpace:"nowrap"}}>Members per team:</label>
                    <input type="number" min={1} max={10} value={teamSize}
                      onChange={e => setTeamSize(parseInt(e.target.value) || 4)}
                      style={{width:"60px",textAlign:"center"}} />
                    <span style={{fontSize:"0.82rem",color:"var(--text-muted)"}}>
                      → {Math.ceil(attendanceList.length / Math.max(1, teamSize))} teams
                    </span>
                    <button type="button" className="btn-push" onClick={autoAssignTeams}>
                      Auto-Assign to Teams
                    </button>
                  </div>
                  <div style={{marginTop:"0.5rem",maxHeight:120,overflowY:"auto",fontSize:"0.82rem",color:"var(--text-secondary)"}}>
                    <strong>{attendanceList.length} students:</strong> {attendanceList.map(s => s.name).join(", ")}
                  </div>
                </div>
              )}
            </div>

            {/* ── OPTION C: DEFAULT NUMBERED TEAMS ── */}
            {gameState.teams.length === 0 && attendanceList.length === 0 && (
              <div className="setup-option-card">
                <h3>Option C — Default Numbered Teams</h3>
                <p style={{fontSize:"0.85rem",color:"var(--text-secondary)",marginBottom:"0.5rem"}}>
                  Quickly create numbered teams (Team 1, Team 2, ...) without uploading any file.
                </p>
                <div style={{display:"flex",gap:"0.5rem",alignItems:"center"}}>
                  <label style={{fontSize:"0.85rem",fontWeight:600,whiteSpace:"nowrap"}}>Number of teams:</label>
                  <input type="number" min={1} max={400} defaultValue={20} id="defaultTeamCount"
                    style={{width:"80px",textAlign:"center"}} />
                  <button type="button" className="btn-push" onClick={() => {
                    const count = Math.min(400, Math.max(1, parseInt(document.getElementById("defaultTeamCount").value) || 20));
                    const teams = Array.from({length: count}, (_, i) => ({
                      id: `team-${i}`, name: `Team ${i+1}`, section: "",
                      members: [{name:"",seat:"",enrol:""}],
                      quarters: [], pendingPrice: null, pendingPromo: null, submitted: false, isIndividual: false,
                    }));
                    setGameState(p => ({...p, teams}));
                    setRosterStatus({type:"ok",message:`${count} default teams created (Team 1 to Team ${count})`});
                  }}>
                    Create Teams
                  </button>
                </div>
              </div>
            )}

            {/* ── STATUS MESSAGE ── */}
            {rosterStatus && <div className={`push-msg ${rosterStatus.type === "ok" ? "ok" : rosterStatus.type === "error" ? "err" : ""}`}>{rosterStatus.message}</div>}

            {/* ── ROSTER PREVIEW ── */}
            {gameState.teams.length > 0 && (
              <div className="setup-option-card" style={{maxHeight:280,overflowY:"auto"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <h3>Roster — {gameState.teams.filter(t=>!t.isIndividual).length} teams, {gameState.teams.filter(t=>t.isIndividual).length} individuals ({gameState.teams.reduce((s,t)=>s+t.members.length,0)} students)</h3>
                  <button type="button" className="btn-push" style={{background:"var(--red-600)",color:"white",padding:"0.2rem 0.6rem",fontSize:"0.75rem"}}
                    onClick={() => { setGameState(p => ({...p, teams:[]})); setAttendanceList([]); setRosterStatus(null); }}>
                    Clear All
                  </button>
                </div>
                <div className="roster-preview">
                  {gameState.teams.map((t, i) => (
                    <div className="rp-row" key={t.id}>
                      <span className="rp-rank">{i+1}</span>
                      <span className="rp-name">{t.name}</span>
                      <span className={`rp-type ${t.isIndividual ? "ind" : "team"}`}>{t.isIndividual ? "Individual" : `${t.members.length} members`}</span>
                      {t.section && <span className="rp-type team" style={{background:"var(--bg-alt)",color:"var(--text-secondary)"}}>{t.section}</span>}
                      <span className="rp-members">{t.members.map(m => m.name).filter(Boolean).join(", ")}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* ── FINALISE BUTTON ── */}
            <button className="login-submit faculty-submit" style={{marginTop:"1rem"}}
              disabled={gameState.teams.length === 0}
              onClick={async () => {
                const resetTeams = gameState.teams.map(t => ({...t, quarters:[], pendingPrice:null, pendingPromo:null, submitted:false}));
                setGameState(p => ({...p, currentQuarter: 1, status: "waiting", quarterStarted: false, teams: resetTeams}));
                setSetupStep("playing"); setTab("control");
                // Write game to Firebase REST API
                const FIREBASE_REST = FIREBASE_REST_BASE;
                try {
                  await fetch(FIREBASE_REST + "/games/" + gameState.gameId + "/meta.json", {
                    method: "PUT",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                      status: "waiting", currentQuarter: 1, quarterStarted: false,
                      sdPenaltyEnabled: gameState.sdPenaltyEnabled,
                      aipHistory: [], aipInput: 10, createdAt: new Date().toISOString(),
                    }),
                  });
                  const teamsObj = {};
                  resetTeams.forEach(t => { teamsObj[t.id] = { name: t.name, section: t.section || "", members: t.members || [], isIndividual: t.isIndividual || false }; });
                  await fetch(FIREBASE_REST + "/games/" + gameState.gameId + "/teams.json", {
                    method: "PUT",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(teamsObj),
                  });
                } catch (e) {
                  console.error("Firebase REST error:", e);
                }
              }}>
              🚀 Finalise Teams & Enter Game ({gameState.teams.length} participants)
            </button>
          </div>
        </div>
      </div>
    );
  }

  // ═══════════════════════════════════════════════════════════════════
  // STEP 3: GAME DASHBOARD — Full control room
  // ═══════════════════════════════════════════════════════════════════
  return (
    <div className="dashboard">
      <nav className="sidebar">
        <div className="sidebar-brand">
          <span className="sb-icon">🍽️</span>
          <div>
            <div className="sb-title">Pricing Sim</div>
            <div className="sb-role">Faculty Control</div>
          </div>
        </div>
        <div className="sidebar-nav">
          {tabs.map(t => (
            <button type="button" key={t.id} className={`nav-item ${tab === t.id ? "active" : ""}`} onClick={(e) => { e.preventDefault(); setTab(t.id); }}>
              <Icon d={t.icon} size={18} />
              <span>{t.label}</span>
            </button>
          ))}
        </div>
        <div className="sidebar-footer">
          <div className="game-status-badge">
            <span className={`status-dot ${gameState.status}`} />
            {gameState.status === "setup" ? "Setup" : `Q${Math.min(gameState.currentQuarter, 12)} of 12 • Phase ${currentPhase}`}
          </div>
          <div className="game-id-badge">{gameState.gameId}</div>
          <button className="nav-item logout-item" onClick={onLogout}>
            <Icon d={Icons.logout} size={18} /><span>Logout</span>
          </button>
        </div>
      </nav>

      <main className="main-content">
        {tab === "control" && (
          <div className="page">
            <div className="page-header">
              <div>
                <h1>Game Control Room</h1>
                <p className="page-desc">Manage quarters, set AIP, and monitor team submissions</p>
              </div>
              <div className="header-badges">
                <div className="hb" style={{borderColor:"var(--amber-600)",background:"var(--amber-100)"}}>
                  <span className="hb-label">Game ID</span>
                  <span className="hb-value" style={{color:"var(--amber-600)",fontSize:"0.85rem",letterSpacing:"0.03em"}}>{gameState.gameId}</span>
                </div>
                <div className="hb" style={{borderColor: cfg.color}}>
                  <span className="hb-label">Phase</span>
                  <span className="hb-value" style={{color: cfg.color}}>{currentPhase}</span>
                </div>
                <div className="hb">
                  <span className="hb-label">Quarter</span>
                  <span className="hb-value">{Math.max(Math.min(gameState.currentQuarter, 12), 1)}</span>
                </div>
                <div className="hb">
                  <span className="hb-label">Teams</span>
                  <span className="hb-value">{gameState.teams.length}</span>
                </div>
              </div>
            </div>

            {/* GAME ACTIVE — show phase indicator and controls */}
            {(gameState.status === "active" || gameState.status === "waiting") && gameState.currentQuarter <= 12 && (
              <>
                <div className="phase-indicator" style={{background: `linear-gradient(135deg, ${cfg.color}, ${cfg.accent})`}}>
                  <div className="pi-tag">{cfg.tag}</div>
                  <h2>{cfg.name}</h2>
                  <p>{cfg.desc}</p>
                  <div className="pi-params">
                    <span>VC: ₹{cfg.vc}/meal</span><span>FC: ₹{cfg.fc}/qtr</span>
                    {cfg.promo && <span className="pi-badge">Promos ON</span>}
                    {cfg.recession && <span className="pi-badge warn">RECESSION</span>}
                  </div>
                </div>

                <div className="control-panel">
                  <div className="control-card">
                    <h3>Set Average Industry Price (AIP)</h3>
                    <p className="cc-desc">Announce this to the class before teams submit their prices.</p>
                    <div className="aip-input-row">
                      <button className="aip-btn" onClick={() => setAipInput(p => Math.max(0, +(p - PRICE_STEP).toFixed(2)))}>−</button>
                      <div className="aip-display">
                        <span className="aip-currency">₹</span>
                        <input type="number" value={aipInput} onChange={e => setAipInput(+e.target.value)} step={PRICE_STEP} min={0} />
                      </div>
                      <button className="aip-btn" onClick={() => setAipInput(p => +(p + PRICE_STEP).toFixed(2))}>+</button>
                    </div>
                  </div>

                  <div className="control-card">
                    <h3>Submissions</h3>
                    <div className="submission-meter">
                      <div className="meter-bar">
                        <div className="meter-fill" style={{width: `${gameState.teams.length > 0 ? (submittedCount / gameState.teams.length)*100 : 0}%`}} />
                      </div>
                      <span className="meter-text">{submittedCount} / {gameState.teams.length} teams submitted</span>
                    </div>
                  </div>

                  {/* Calculated AIP from submissions — shown before faculty processes */}
                  {submittedPrices.length > 0 && (
                    <div className="control-card calculated-aip-card">
                      <h3>📊 Calculated Average Industry Price</h3>
                      <div className="calc-aip-grid">
                        <div className="calc-aip-main">
                          <span className="calc-aip-label">Average Price</span>
                          <span className="calc-aip-value">₹{calculatedAIP.toFixed(2)}</span>
                        </div>
                        <div className="calc-aip-stat">
                          <span className="calc-aip-label">Min</span>
                          <span className="calc-aip-stat-val">₹{priceMin.toFixed(2)}</span>
                        </div>
                        <div className="calc-aip-stat">
                          <span className="calc-aip-label">Max</span>
                          <span className="calc-aip-stat-val">₹{priceMax.toFixed(2)}</span>
                        </div>
                        {priceSD !== null && (
                          <div className="calc-aip-stat">
                            <span className="calc-aip-label">Std Dev</span>
                            <span className="calc-aip-stat-val">₹{priceSD}</span>
                          </div>
                        )}
                        <div className="calc-aip-stat">
                          <span className="calc-aip-label">Submissions</span>
                          <span className="calc-aip-stat-val">{submittedPrices.length}</span>
                        </div>
                      </div>
                      <p className="calc-aip-note">Auto-populated into the AIP field above. You can adjust before processing.</p>
                    </div>
                  )}

                  {/* Two-step: Start Quarter → then Process & Advance */}
                  {!gameState.quarterStarted ? (
                    <button className="advance-btn" onClick={startCurrentQuarter} style={{background:"var(--green-700)"}}>
                      <Icon d={Icons.play} size={20} />
                      Start Quarter {gameState.currentQuarter} — Begin Timer for Students
                    </button>
                  ) : (
                    <>
                      {/* Late Submission Controls — shown after quarter is started */}
                      <LateSubmissionPanel
                        gameId={gameState.gameId}
                        teams={gameState.teams}
                        currentQuarter={gameState.currentQuarter}
                        firebaseRest={FIREBASE_REST}
                      />
                      <button className="advance-btn" onClick={processQuarterAndAdvance}>
                        <Icon d={Icons.check} size={20} />
                        Process Quarter {gameState.currentQuarter} & Advance
                      </button>
                    </>
                  )}
                </div>
              </>
            )}

            {/* GAME FINISHED */}
            {gameState.status === "finished" && (
              <div className="control-card finished-card">
                <div className="finished-icon">🏁</div>
                <h2>Simulation Complete</h2>
                <p>Game <strong>{gameState.gameId}</strong> — All 12 quarters played. Check Leaderboard, Scores, and Analytics for results.</p>
                <button type="button" className="btn-push" style={{marginTop:"1rem"}} onClick={() => { setSetupStep("pregame"); setGameState(p => ({...p, status:"setup", currentQuarter:0, quarterStarted:false, gameId: generateGameId(), teams: [], aipHistory:[]})); }}>
                  Set Up New Game
                </button>
              </div>
            )}

            {/* RESET GAME — always available */}
            <div className="control-card" style={{borderLeft:"4px solid var(--red-600)",marginTop:"1rem"}}>
              <h3>Reset Game</h3>
              <p className="cc-desc">Discard all current data and return to the pre-game setup. This cannot be undone.</p>
              <button type="button" className="btn-push" style={{background:"var(--red-600)",color:"white",marginTop:"0.5rem"}}
                onClick={() => {
                  if (window.confirm("Are you sure you want to reset? All game data will be lost.")) {
                    setSetupStep("pregame");
                    setGameState({
                      currentQuarter: 0, status: "setup", aipHistory: [], teams: [],
                      sheetsUrl: "", gameId: generateGameId(),
                      sdPenaltyEnabled: true, allowIndividualPlay: true, quarterStarted: false,
                    });
                    setAipInput(INIT_PRICE);
                  }
                }}>
                🔄 Reset & Start Over
              </button>
            </div>

            {gameState.aipHistory.length > 0 && (
              <div className="aip-history">
                <h3>AIP History</h3>
                <div className="aip-chips">
                  {gameState.aipHistory.map((h, i) => (
                    <div className="aip-chip" key={i}>
                      <span className="ac-q">Q{h.quarter}</span>
                      <span className="ac-v">{fmt(h.aip)}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="control-card" style={{marginTop:"1rem"}}>
              <h3>Game Settings</h3>
              <div className="settings-toggles">
                <label className="toggle-row">
                  <input type="checkbox" checked={gameState.sdPenaltyEnabled}
                    onChange={e => setGameState(p => ({...p, sdPenaltyEnabled: e.target.checked}))} />
                  <span>SD Penalty — penalise outlier prices (beyond 1.5 SD from mean)</span>
                </label>
                <label className="toggle-row">
                  <input type="checkbox" checked={gameState.allowIndividualPlay}
                    onChange={e => setGameState(p => ({...p, allowIndividualPlay: e.target.checked}))} />
                  <span>Allow Individual Play — students can join without a team</span>
                </label>
              </div>
            </div>
          </div>
        )}

        {tab === "leaderboard" && (
          <div className="page">
            <div className="page-header">
              <div><h1>Leaderboard</h1><p className="page-desc">Real-time rankings by composite score</p></div>
            </div>
            {scoredTeams.filter(t => t.quarters && t.quarters.length > 0).length === 0 ? (
              <div className="control-card" style={{textAlign:"center",padding:"3rem"}}>
                <div style={{fontSize:"3rem",marginBottom:"0.5rem"}}>📊</div>
                <h2 style={{marginBottom:"0.5rem"}}>No Results Yet</h2>
                <p className="cc-desc">The leaderboard will populate after teams complete their first quarter. Start the game and process Q1 to see rankings here.</p>
              </div>
            ) : (
            <div className="leaderboard">
              <div className="lb-header-row">
                <span className="lb-rank">#</span>
                <span className="lb-team">Team</span>
                <span className="lb-type">Type</span>
                <span className="lb-stat">Revenue</span>
                <span className="lb-stat">Profit</span>
                <span className="lb-stat">Penalties</span>
                <span className="lb-stat">Score</span>
              </div>
              {scoredTeams.filter(t => t.quarters && t.quarters.length > 0).map((team, i) => {
                return (
                  <div className={`lb-row ${i < 3 ? "lb-top" : ""}`} key={team.id}>
                    <span className={`lb-rank rank-${i+1}`}>{i < 3 ? ["🥇","🥈","🥉"][i] : i+1}</span>
                    <span className="lb-team">{team.name}</span>
                    <span className="lb-type">{team.isIndividual ? "👤" : "🤝"}</span>
                    <span className="lb-stat">{fmt(team.totRevenue)}</span>
                    <span className={`lb-stat ${team.totProfit >= 0 ? "profit-pos" : "profit-neg"}`}>{fmt(team.totProfit)}</span>
                    <span className={`lb-stat ${team.penalties > 0 ? "profit-neg" : ""}`}>{team.penalties || "—"}</span>
                    <span className="lb-stat"><strong>{team.score}</strong></span>
                  </div>
                );
              })}
            </div>
            )}
          </div>
        )}

        {tab === "scores" && (
          <div className="page">
            <div className="page-header">
              <div><h1>Scores & Rankings</h1><p className="page-desc">Composite scoring index with component breakdown</p></div>
            </div>

            <div className="ss-card" style={{marginBottom:"1rem"}}>
              <h3>Scoring Index Methodology</h3>
              <p style={{fontSize:"0.88rem",color:"var(--text-secondary)",marginBottom:"0.75rem"}}>
                Each team's final score (0–100) is computed from six normalised components:
              </p>
              <div className="score-weights">
                <div className="sw-item"><span className="sw-pct">35%</span><span className="sw-label">Cumulative Profit</span></div>
                <div className="sw-item"><span className="sw-pct">15%</span><span className="sw-label">Total Revenue</span></div>
                <div className="sw-item"><span className="sw-pct">15%</span><span className="sw-label">Total Meals Sold</span></div>
                <div className="sw-item"><span className="sw-pct">15%</span><span className="sw-label">Consistency (low profit SD)</span></div>
                <div className="sw-item"><span className="sw-pct">10%</span><span className="sw-label">Fewer SD Penalties</span></div>
                <div className="sw-item"><span className="sw-pct">10%</span><span className="sw-label">Profit Growth (late vs early)</span></div>
              </div>
            </div>

            {gameState.sdPenaltyEnabled && (
              <div className="ss-card" style={{marginBottom:"1rem",borderLeft:"4px solid var(--red-600)"}}>
                <h3>SD Penalty Rules</h3>
                <p style={{fontSize:"0.88rem",color:"var(--text-secondary)"}}>
                  Each quarter, prices are checked against the group's standard deviation. Outliers incur profit penalties:
                </p>
                <div className="score-weights" style={{marginTop:"0.5rem"}}>
                  <div className="sw-item"><span className="sw-pct" style={{background:"var(--amber-100)",color:"var(--amber-600)"}}>10%</span><span className="sw-label">Z-score 1.5–2.0 (moderate outlier)</span></div>
                  <div className="sw-item"><span className="sw-pct" style={{background:"var(--red-100)",color:"var(--red-600)"}}>20%</span><span className="sw-label">Z-score &gt; 2.0 (extreme outlier)</span></div>
                </div>
              </div>
            )}

            <div className="ss-card">
              <h3>Full Rankings</h3>
              <div className="summary-table-wrap">
                <table className="summary-table">
                  <thead>
                    <tr>
                      <th>Rank</th><th>Team/Individual</th><th>Type</th><th>Profit</th><th>Revenue</th>
                      <th>Sales</th><th>Penalties</th><th>Growth</th><th>Score</th>
                    </tr>
                  </thead>
                  <tbody>
                    {scoredTeams.map(t => (
                      <tr key={t.id} className={t.rank <= 3 ? "highlight-row" : ""}>
                        <td><strong>{t.rank <= 3 ? ["🥇","🥈","🥉"][t.rank-1] : `#${t.rank}`}</strong></td>
                        <td><strong>{t.name}</strong></td>
                        <td>{t.isIndividual ? "Individual" : "Team"}</td>
                        <td className={t.totProfit >= 0 ? "profit-pos" : "profit-neg"}>{fmt(t.totProfit)}</td>
                        <td>{fmt(t.totRevenue)}</td>
                        <td>{(t.totSales || 0).toLocaleString("en-IN")}</td>
                        <td className={t.penalties > 0 ? "profit-neg" : ""}>{t.penalties || 0}</td>
                        <td>{(t.growth || 0) > 0 ? "+" : ""}{t.growth || 0}%</td>
                        <td><strong style={{fontSize:"1.1em"}}>{t.score}</strong></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {tab === "teams" && (
          <div className="page">
            <div className="page-header">
              <div><h1>All Teams ({gameState.teams.length})</h1><p className="page-desc">View teams, team leads, and performance. Click a team lead to reassign.</p></div>
            </div>
            <div className="teams-grid">
              {gameState.teams.map(team => {
                const totProf = team.quarters.reduce((s,q) => s+q.profit, 0);
                const totRev = team.quarters.reduce((s,q) => s+q.revenue, 0);
                const preloaded = PRELOADED_BY_NAME[team.name];
                const members = preloaded ? preloaded.members : (team.members || []);
                const overrideLead = teamLeads[team.id];
                const currentLead = overrideLead
                  ? { name: overrideLead.leadName, enrol: overrideLead.leadEnrol }
                  : (members[0] || { name: "—", enrol: "" });
                const isChanging = changingLeadFor === team.id;
                return (
                  <div className="team-tile" key={team.id}>
                    <div className="tt-header">
                      <h3>{team.name}</h3>
                      <span className="tt-section">{team.section || preloaded?.section || ""}</span>
                    </div>

                    {/* Team Lead */}
                    <div className="tt-lead">
                      <span className="tt-lead-crown">👑</span>
                      <div className="tt-lead-info">
                        <span className="tt-lead-label">Team Lead</span>
                        <span className="tt-lead-name">{currentLead.name}</span>
                        {currentLead.enrol && <span className="tt-lead-enrol">{currentLead.enrol}</span>}
                      </div>
                      <button className="btn-change-lead" onClick={() => {
                        setChangingLeadFor(isChanging ? null : team.id);
                        setNewLeadIndex(0);
                      }}>
                        {isChanging ? "Cancel" : "Change"}
                      </button>
                    </div>

                    {/* Change Lead Panel */}
                    {isChanging && (
                      <div className="tt-change-lead-panel">
                        <p className="tt-cl-hint">Select the new team lead:</p>
                        {members.map((m, mi) => (
                          <label key={mi} className="tt-cl-option">
                            <input type="radio" name={`lead-${team.id}`} value={mi}
                              checked={newLeadIndex === mi}
                              onChange={() => setNewLeadIndex(mi)} />
                            <span className="tt-cl-name">{m.name}</span>
                            <span className="tt-cl-enrol">{m.enrol}</span>
                            {mi === 0 && !overrideLead && <span className="tt-cl-badge">Current</span>}
                          </label>
                        ))}
                        <button className="btn-push" style={{marginTop:"0.5rem",width:"100%"}}
                          onClick={async () => {
                            const newLead = members[newLeadIndex];
                            const update = { leadName: newLead.name, leadEnrol: newLead.enrol || "" };
                            setTeamLeads(prev => ({ ...prev, [team.id]: update }));
                            setChangingLeadFor(null);
                            // Write to Firebase so student login can read it
                            try {
                              await fetch(
                                `${FIREBASE_REST}/games/${gameState.gameId}/teamLeads/${team.id}.json`,
                                { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(update) }
                              );
                            } catch(e) { console.error("Firebase lead update error:", e); }
                          }}>
                          Assign {members[newLeadIndex]?.name} as Team Lead
                        </button>
                      </div>
                    )}

                    {/* Members */}
                    <div className="tt-members">
                      {members.filter(m => m.name).map((m, i) => (
                        <span key={i} className={`tt-member-tag ${currentLead.name === m.name ? "lead" : ""}`}>
                          {currentLead.name === m.name && "👑 "}{m.name}
                        </span>
                      ))}
                    </div>

                    <div className="tt-stats">
                      <div><span className="tt-label">Revenue</span><span className="tt-val">{fmt(totRev)}</span></div>
                      <div><span className="tt-label">Profit</span><span className={`tt-val ${totProf >= 0 ? "profit-pos" : "profit-neg"}`}>{fmt(totProf)}</span></div>
                      <div><span className="tt-label">Quarters</span><span className="tt-val">{team.quarters.length}</span></div>
                    </div>
                    {team.quarters.length > 0 && (
                      <div className="tt-sparkline">
                        {team.quarters.map((q, qi) => {
                          const maxP = Math.max(...team.quarters.map(x => Math.abs(x.profit)), 1);
                          const h = Math.max(4, (Math.abs(q.profit) / maxP) * 32);
                          return <div key={qi} className={`spark-bar ${q.profit >= 0 ? "pos" : "neg"}`} style={{height: h}} title={`Q${qi+1}: ${fmt(q.profit)}`} />;
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab === "analytics" && (
          <div className="page">
            <div className="page-header">
              <div><h1>Analytics</h1><p className="page-desc">Aggregate metrics across all teams</p></div>
            </div>
            <AnalyticsView teams={gameState.teams} aipHistory={gameState.aipHistory} />
          </div>
        )}

        {tab === "sheets" && (
          <div className="page">
            <div className="page-header">
              <div>
                <h1>Google Sheets Export</h1>
                <p className="page-desc">
                  {gameState.sheetsUrl
                    ? "✅ Auto-sync enabled — data saves to Google Sheets after every quarter"
                    : "Paste your Web App URL below to enable auto-sync after every quarter"}
                </p>
              </div>
            </div>
            <div className="sheets-setup">
              <div className="ss-card">
                <h3>Web App URL</h3>
                <p>Paste the Google Apps Script Web App URL below. Once set, data will <strong>automatically save to Google Sheets after every quarter</strong> — no manual push needed.</p>
                <div className="ss-input-row">
                  <input type="url" placeholder="https://script.google.com/macros/s/.../exec"
                    value={gameState.sheetsUrl} onChange={e => updateSheetsUrl(e.target.value)} />
                  <button className="btn-push" style={{background:"#7f8c8d"}} onClick={async () => {
                    setPushStatus({ ok: true, message: "Testing connection..." });
                    const testPayload = { teamName: "TEST", timestamp: new Date().toISOString(), members: [{name:"Test User"}], strategy: {}, quarters: [{ quarter: 1, phase: 1, ownPrice: 8, avgCompPrice: 10, totalSales: 50, revenue: 400, profit: 100 }], conclusions: {} };
                    const res = await pushToSheets(gameState.sheetsUrl, testPayload);
                    setPushStatus(res);
                  }} disabled={!gameState.sheetsUrl}>
                    Test Connection
                  </button>
                  <button className="btn-push" onClick={async () => {
                    setPushStatus(null);
                    const scored = calcScores(gameState.teams);
                    const payload = { teams: scored.map(t => ({
                      name: t.name, section: t.section, members: t.members,
                      isIndividual: t.isIndividual || false,
                      rank: t.rank, score: t.score, penalties: t.penalties, growth: t.growth,
                      totProfit: t.totProfit, totRevenue: t.totRevenue, totSales: t.totSales,
                      quarters: t.quarters.map((q,i) => ({ quarter: i+1, ...q }))
                    })), aipHistory: gameState.aipHistory, sdPenaltyEnabled: gameState.sdPenaltyEnabled,
                    gameId: gameState.gameId, timestamp: new Date().toISOString() };
                    const res = await pushToSheets(gameState.sheetsUrl, payload);
                    setPushStatus(res);
                  }} disabled={!gameState.sheetsUrl}>
                    <Icon d={Icons.send} size={16} /> Push All Data
                  </button>
                </div>
                {pushStatus && <div className={`push-msg ${pushStatus.ok ? "ok" : "err"}`}>{pushStatus.message}</div>}
              </div>
              <SheetsSetupGuide />
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

function AnalyticsView({ teams, aipHistory }) {
  if (!teams.length || !teams[0].quarters.length) {
    return <div className="empty-state"><p>No data yet. Process at least one quarter to see analytics.</p></div>;
  }

  const maxQ = Math.max(...teams.map(t => t.quarters.length));
  const quarterStats = [];
  for (let qi = 0; qi < maxQ; qi++) {
    const prices = teams.map(t => t.quarters[qi]?.price).filter(Boolean);
    const profits = teams.map(t => t.quarters[qi]?.profit).filter(p => p !== undefined);
    const revenues = teams.map(t => t.quarters[qi]?.revenue).filter(Boolean);
    quarterStats.push({
      quarter: qi + 1,
      avgPrice: prices.length ? prices.reduce((a,b)=>a+b,0)/prices.length : 0,
      minPrice: prices.length ? Math.min(...prices) : 0,
      maxPrice: prices.length ? Math.max(...prices) : 0,
      avgProfit: profits.length ? profits.reduce((a,b)=>a+b,0)/profits.length : 0,
      totalRevenue: revenues.reduce((a,b)=>a+b,0),
      aip: aipHistory[qi]?.aip || 0,
    });
  }

  const maxAbsProfit = Math.max(...quarterStats.map(q => Math.abs(q.avgProfit)), 1);

  return (
    <div className="analytics-content">
      <div className="analytics-grid">
        <div className="an-card">
          <h3>Price Distribution by Quarter</h3>
          <div className="price-range-chart">
            {quarterStats.map((qs, i) => (
              <div className="prc-col" key={i}>
                <div className="prc-bar-wrap">
                  <div className="prc-range" style={{
                    bottom: `${(qs.minPrice / 15) * 100}%`,
                    height: `${Math.max(4, ((qs.maxPrice - qs.minPrice) / 15) * 100)}%`
                  }}>
                    <div className="prc-avg" style={{ bottom: `${((qs.avgPrice - qs.minPrice) / Math.max(qs.maxPrice - qs.minPrice, 0.01)) * 100}%` }} />
                  </div>
                </div>
                <span className="prc-label">Q{i+1}</span>
              </div>
            ))}
          </div>
          <div className="chart-legend">
            <span><span className="legend-box range" />Price Range</span>
            <span><span className="legend-dot avg" />Avg Price</span>
          </div>
        </div>

        <div className="an-card">
          <h3>Average Profit per Team</h3>
          <div className="profit-bars">
            {quarterStats.map((qs, i) => (
              <div className="pb-row" key={i}>
                <span className="pb-label">Q{i+1}</span>
                <div className="pb-track">
                  <div className={`pb-fill ${qs.avgProfit >= 0 ? "pos" : "neg"}`}
                    style={{width: `${Math.max(3,(Math.abs(qs.avgProfit)/maxAbsProfit)*100)}%`}} />
                </div>
                <span className={`pb-val ${qs.avgProfit < 0 ? "profit-neg" : ""}`}>{fmt(Math.round(qs.avgProfit))}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="an-card">
        <h3>Quarter Summary Table</h3>
        <div className="summary-table-wrap">
          <table className="summary-table">
            <thead><tr><th>Qtr</th><th>Phase</th><th>AIP</th><th>Avg Price</th><th>Min</th><th>Max</th><th>Avg Profit</th><th>Total Revenue</th></tr></thead>
            <tbody>
              {quarterStats.map((qs, i) => (
                <tr key={i}>
                  <td><strong>Q{qs.quarter}</strong></td>
                  <td><span className={`phase-chip p${getPhase(qs.quarter)}`}>P{getPhase(qs.quarter)}</span></td>
                  <td>{fmtD(qs.aip)}</td><td>{fmtD(qs.avgPrice)}</td><td>{fmtD(qs.minPrice)}</td><td>{fmtD(qs.maxPrice)}</td>
                  <td className={qs.avgProfit >= 0 ? "profit-pos" : "profit-neg"}>{fmt(Math.round(qs.avgProfit))}</td>
                  <td>{fmt(Math.round(qs.totalRevenue))}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function SheetsSetupGuide() {
  return (
    <div className="ss-card guide-card">
      <h3>Setup Instructions</h3>
      <div className="setup-steps">
        <div className="step"><span className="step-num">1</span><div><strong>Create a Google Sheet</strong><p>Open Google Sheets and create a new blank spreadsheet.</p></div></div>
        <div className="step"><span className="step-num">2</span><div><strong>Open Apps Script</strong><p>Extensions → Apps Script in the menu bar.</p></div></div>
        <div className="step"><span className="step-num">3</span><div><strong>Paste the Script</strong><p>Delete default code. Paste contents of <code>google-apps-script.js</code>.</p></div></div>
        <div className="step"><span className="step-num">4</span><div><strong>Deploy as Web App</strong><p>Deploy → New Deployment → Web app. Execute as: Me, Access: Anyone.</p></div></div>
        <div className="step"><span className="step-num">5</span><div><strong>Copy URL</strong><p>Copy the Web App URL and paste it above.</p></div></div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// STUDENT DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════
function StudentDashboard({ session, onLogout }) {
  const [gamePhase, setGamePhase] = useState("pregame"); // "pregame" or "playing"
  const [tab, setTab] = useState("play");
  const [quarters, setQuarters] = useState([]);
  const [currentInput, setCurrentInput] = useState({ price: INIT_PRICE, promo: 0 });
  const [strategy, setStrategy] = useState({ generic: "", year1: "", year2: "", year3: "" });
  const [aip, setAip] = useState(INIT_PRICE);
  const [submitted, setSubmitted] = useState(false);
  const [conclusions, setConclusions] = useState({ revenue: "", profit: "" });
  const [quarterComments, setQuarterComments] = useState({ observations: "", nextStrategy: "" });
  const [quarterReady, setQuarterReady] = useState(false); // false = waiting for faculty to start
  const [timerSeconds, setTimerSeconds] = useState(0);
  const [timerRunning, setTimerRunning] = useState(false);
  const timerRef = useRef(null);

  const currentQuarter = quarters.length + 1;
  const phase = getPhase(Math.min(currentQuarter, 12));
  const cfg = PHASE_CONFIG[phase];
  const prevData = quarters.length > 0 ? quarters[quarters.length - 1] : null;
  const lastRev = prevData ? prevData.revenue : INIT_CUSTOMERS * INIT_PRICE;
  const maxPromo = cfg.promo ? Math.floor(lastRev * 0.2) : 0;

  const preview = currentQuarter <= 12 ? simulateQuarter(currentInput.price, aip, phase, Math.min(currentInput.promo, maxPromo), prevData) : null;

  // Timer logic: Q1 of each phase = 6 min, Q2-Q4 of each phase = 3 min
  const getTimerDuration = (qNum) => {
    const qInPhase = ((qNum - 1) % 4) + 1; // 1,2,3,4 within each phase
    return qInPhase === 1 ? 360 : 180; // 6 min or 3 min in seconds
  };

  const startTimerRef = useRef(null);
  const [timerExpired, setTimerExpired] = useState(false); // true when timer hits 0 without submission
  const [lateSubmitAllowed, setLateSubmitAllowed] = useState(false); // faculty grants extension
  const [isLateSubmission, setIsLateSubmission] = useState(false); // this submission was late

  const startTimer = useCallback(() => {
    const duration = getTimerDuration(currentQuarter);
    setTimerSeconds(duration);
    setTimerRunning(true);
    setQuarterReady(true);
    setSubmitted(false);
    setTimerExpired(false);
    setLateSubmitAllowed(false);
    setIsLateSubmission(false);
    if (timerRef.current) clearInterval(timerRef.current);
    timerRef.current = setInterval(() => {
      setTimerSeconds(prev => {
        if (prev <= 1) {
          clearInterval(timerRef.current);
          setTimerRunning(false);
          setTimerExpired(true); // AUTO-LOCK when timer hits 0
          return 0;
        }
        return prev - 1;
      });
    }, 1000);
  }, [currentQuarter]);
  startTimerRef.current = startTimer;

  // Q1 does NOT auto-start. Student must wait for faculty to start every quarter.

  // SYNC: Poll Firebase REST API directly every 2 seconds
  const quarterReadyRef = useRef(quarterReady);
  const timerRunningRef = useRef(timerRunning);
  const submittedRef = useRef(submitted);
  const lastStartedQuarterRef = useRef(0); // Track which quarter we last started
  quarterReadyRef.current = quarterReady;
  timerRunningRef.current = timerRunning;
  submittedRef.current = submitted;
  const [debugInfo, setDebugInfo] = useState("");

  useEffect(() => {
    if (gamePhase !== "playing") return;
    if (!session.gameCode) return;
    const gameCode = session.gameCode;
    const firebaseUrl = "https://pricing-simulation-4ceee-default-rtdb.firebaseio.com/games/" + gameCode + "/meta.json";
    const storageKey = "pricing-sim-" + gameCode;
    let cancelled = false;

    const checkForQuarterStart = async () => {
      if (cancelled) return;
      // Skip if already ready, timer running, or student already submitted this quarter
      if (quarterReadyRef.current || timerRunningRef.current || submittedRef.current) return;

      // Check Firebase REST API directly (cache: no-store prevents stale responses)
      try {
        const res = await fetch(firebaseUrl + "?_t=" + Date.now(), { cache: "no-store" });
        if (res.ok) {
          const meta = await res.json();
          if (meta && meta.quarterStarted && meta.currentQuarter > lastStartedQuarterRef.current) {
            // Only start if this is a NEW quarter we haven't started yet
            lastStartedQuarterRef.current = meta.currentQuarter;
            if (meta.aipInput != null) setAip(meta.aipInput);
            if (startTimerRef.current) startTimerRef.current();
            return;
          }
          // Faculty granted late submission for this team
          if (meta && meta.lateAllowed) {
            const teamKey = session.selectedTeamId;
            if (meta.lateAllowed[teamKey] && !submittedRef.current) {
              setLateSubmitAllowed(true);
              setIsLateSubmission(true);
            }
          }
        }
      } catch (e) {}

      // Fallback: Check localStorage (same-browser)
      try {
        const raw = localStorage.getItem(storageKey);
        if (raw) {
          const msg = JSON.parse(raw);
          if (msg.type === "QUARTER_STARTED" && msg.quarter > lastStartedQuarterRef.current) {
            lastStartedQuarterRef.current = msg.quarter;
            if (msg.aip != null) setAip(msg.aip);
            if (startTimerRef.current) startTimerRef.current();
            return;
          }
        }
      } catch (e) {}
    };

    // Check immediately, then every 2 seconds
    checkForQuarterStart();
    const interval = setInterval(checkForQuarterStart, 2000);

    // Also listen for storage events (instant when same browser)
    const handleStorage = (e) => {
      if (e.key === storageKey) checkForQuarterStart();
    };
    window.addEventListener("storage", handleStorage);

    return () => {
      cancelled = true;
      clearInterval(interval);
      window.removeEventListener("storage", handleStorage);
    };
  }, [gamePhase, session.gameCode]);

  // Cleanup timer
  useEffect(() => { return () => { if (timerRef.current) clearInterval(timerRef.current); }; }, []);

  const formatTimer = (secs) => {
    const m = Math.floor(secs / 60);
    const s = secs % 60;
    return `${m}:${String(s).padStart(2, "0")}`;
  };

  const timerColor = timerSeconds <= 30 ? "var(--red-600)" : timerSeconds <= 60 ? "var(--amber-600)" : "var(--green-700)";
  const timerPct = timerRunning ? (timerSeconds / getTimerDuration(currentQuarter)) * 100 : 0;

  // Validation: comments must be at least 100 characters each
  const [commentError, setCommentError] = useState("");
  const obsLen = quarterComments.observations.trim().length;
  const stratLen = quarterComments.nextStrategy.trim().length;
  const commentsValid = obsLen >= 100 && stratLen >= 100;

  const submitQuarter = () => {
    // Validate comments
    if (obsLen < 100 || stratLen < 100) {
      setCommentError("Please elaborate — both Observations and Strategy must be at least 100 characters each.");
      return;
    }
    setCommentError("");
    const result = {
      ...simulateQuarter(currentInput.price, aip, phase, Math.min(currentInput.promo, maxPromo), prevData),
      observations: quarterComments.observations,
      nextStrategy: quarterComments.nextStrategy,
      lateSubmission: isLateSubmission, // flag for faculty/scoring
    };
    setQuarters(prev => [...prev, result]);
    setSubmitted(true);
    setQuarterReady(false); // LOCK — stays false until faculty processes
    // Sync to Firebase REST API
    if (session.gameCode && session.selectedTeamId) {
      const teamKey = session.selectedTeamId.replace(/\s+/g, "_");
      fetch("https://pricing-simulation-4ceee-default-rtdb.firebaseio.com/games/" + session.gameCode + "/teams/" + teamKey + ".json", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pendingPrice: currentInput.price, pendingPromo: currentInput.promo || 0, submitted: true, lateSubmission: isLateSubmission }),
      }).catch(e => console.error("Firebase submit error:", e));
    }
    if (timerRef.current) clearInterval(timerRef.current);
    setTimerRunning(false);
    setTimerSeconds(0);
    setTimerExpired(false);
    setLateSubmitAllowed(false);
    setTimeout(() => {
      setSubmitted(false);
      setCurrentInput({ price: currentInput.price, promo: 0 });
      setQuarterComments({ observations: "", nextStrategy: "" });
      setIsLateSubmission(false);
    }, 1500);
  };

  // Faculty unlock: student can manually trigger next quarter timer
  // In a real multi-player setup, this would come from a shared server/firebase
  const unlockNextQuarter = () => {
    startTimer();
  };

  const totRev = quarters.reduce((s,q) => s+q.revenue, 0);
  const totProf = quarters.reduce((s,q) => s+q.profit, 0);
  const totSales = quarters.reduce((s,q) => s+q.totalSales, 0);

  const unlockedPhase = currentQuarter <= 4 ? 1 : currentQuarter <= 8 ? 2 : 3;

  // PRE-GAME STRATEGY SCREEN
  if (gamePhase === "pregame") {
    return (
      <div className="login-wrapper">
        <div className="login-bg-pattern" />
        <div className="login-container login-form-container student-form" style={{maxWidth:620}}>
          <div className="login-brand compact">
            <div className="brand-icon small">📋</div>
            <h2>Game Details & Strategy</h2>
          </div>

          {/* Game Info Summary */}
          <div className="setup-option-card" style={{marginBottom:"0.75rem",background:"var(--bg-alt)"}}>
            <div style={{display:"flex",flexWrap:"wrap",gap:"1rem",fontSize:"0.88rem"}}>
              <div><span style={{fontWeight:700,color:"var(--text-muted)",fontSize:"0.72rem",textTransform:"uppercase",letterSpacing:"0.06em",display:"block"}}>Game Code</span><span style={{fontFamily:"'Courier New',monospace",fontWeight:700,color:"var(--amber-600)"}}>{session.gameCode}</span></div>
              <div><span style={{fontWeight:700,color:"var(--text-muted)",fontSize:"0.72rem",textTransform:"uppercase",letterSpacing:"0.06em",display:"block"}}>Team</span>{session.teamName}</div>
              <div><span style={{fontWeight:700,color:"var(--text-muted)",fontSize:"0.72rem",textTransform:"uppercase",letterSpacing:"0.06em",display:"block"}}>Members</span>{session.members?.filter(m=>m.name).map(m=>m.name).join(", ") || "—"}</div>
            </div>
            <div style={{marginTop:"0.75rem",fontSize:"0.85rem",color:"var(--text-secondary)"}}>
              <strong>About the game:</strong> You run a food canteen for the government. Over 12 quarters (3 phases), you set meal prices to maximise profit.
              Phase 1 is market entry, Phase 2 adds promotions, Phase 3 introduces a recession with higher costs.
            </div>
          </div>

          <div className="login-form">
            <p style={{color:"var(--text-secondary)",fontSize:"0.9rem",marginBottom:"0.75rem"}}>
              Before the simulation begins, outline your pricing strategy for the next 3 years.
            </p>
            <div className="form-field">
              <label>Overall Generic Strategy *</label>
              <textarea rows={3} value={strategy.generic} onChange={e => setStrategy(p => ({...p, generic: e.target.value}))} placeholder="e.g. Penetration pricing to gain market share, then gradually increase..." />
            </div>
            <div className="form-field">
              <label>Year 1 Objectives (Q1–Q4) — Market Entry *</label>
              <textarea rows={2} value={strategy.year1} onChange={e => setStrategy(p => ({...p, year1: e.target.value}))} placeholder="How will you establish your presence?" />
            </div>
            <div className="form-field">
              <label>Year 2 Objectives (Q5–Q8) — Promotions</label>
              <textarea rows={2} value={strategy.year2} onChange={e => setStrategy(p => ({...p, year2: e.target.value}))} placeholder="How will you use promotions? Growth targets?" />
            </div>
            <div className="form-field">
              <label>Year 3 Objectives (Q9–Q12) — Recession</label>
              <textarea rows={2} value={strategy.year3} onChange={e => setStrategy(p => ({...p, year3: e.target.value}))} placeholder="How will you survive the recession?" />
            </div>
            <button className="login-submit student-submit"
              onClick={() => { if (!strategy.generic.trim()) return; setGamePhase("playing"); }}
              disabled={!strategy.generic.trim()}>
              Enter Simulation →
            </button>
            {!strategy.generic.trim() && <p className="login-hint">Fill in at least the generic strategy to proceed</p>}
          </div>
        </div>
      </div>
    );
  }

  const tabs = [
    { id: "play", label: "Play", icon: Icons.play },
    { id: "strategy", label: "Strategy", icon: Icons.edit },
    { id: "history", label: "History", icon: Icons.chart },
    { id: "formulas", label: "Formulas", icon: Icons.info },
  ];

  return (
    <div className="dashboard student-dashboard">
      <nav className="sidebar student-sidebar">
        <div className="sidebar-brand">
          <span className="sb-icon">🍽️</span>
          <div>
            <div className="sb-title">{session.teamName}</div>
            <div className="sb-role">{session.section}{session.isIndividual ? " (Individual)" : ""}</div>
          </div>
        </div>
        <div className="sidebar-nav">
          {tabs.map(t => (
            <button type="button" key={t.id} className={`nav-item ${tab === t.id ? "active" : ""}`} onClick={(e) => { e.preventDefault(); setTab(t.id); }}>
              <Icon d={t.icon} size={18} /><span>{t.label}</span>
            </button>
          ))}
        </div>
        <div className="sidebar-footer">
          <div className="game-status-badge student-badge">
            <span className="status-dot active" />
            Q{Math.min(currentQuarter, 12)} • Phase {phase}
          </div>
          <div className="student-kpis">
            <div className="sk"><span className="sk-l">Revenue</span><span className="sk-v">{fmt(totRev)}</span></div>
            <div className="sk"><span className="sk-l">Profit</span><span className={`sk-v ${totProf >= 0 ? "profit-pos" : "profit-neg"}`}>{fmt(totProf)}</span></div>
          </div>
          <button type="button" className="nav-item logout-item" onClick={onLogout}>
            <Icon d={Icons.logout} size={18} /><span>Logout</span>
          </button>
        </div>
      </nav>

      <main className="main-content">
        {tab === "play" && (
          <div className="page">
            <div className="page-header">
              <div>
                <h1>Quarter {Math.min(currentQuarter, 13)}</h1>
                <p className="page-desc">{currentQuarter <= 12 ? cfg.desc : "Simulation complete!"}</p>
              </div>
              <div className="header-badges">
                <div className="hb" style={{borderColor: cfg.color}}><span className="hb-label">Phase</span><span className="hb-value" style={{color: cfg.color}}>{phase}</span></div>
                <div className="hb"><span className="hb-label">VC</span><span className="hb-value">₹{cfg.vc}</span></div>
                <div className="hb"><span className="hb-label">FC</span><span className="hb-value">₹{cfg.fc}</span></div>
              </div>
            </div>

            {currentQuarter <= 12 ? (
              <>
                {/* TIMER BAR */}
                <div className="timer-bar" style={{borderColor: timerExpired ? "var(--red-600)" : timerColor}}>
                  <div className="timer-progress" style={{width: `${timerPct}%`, background: timerColor}} />
                  <div className="timer-content">
                    <span className="timer-label">
                      {timerRunning
                        ? (currentQuarter % 4 === 1 || currentQuarter === 1 ? "First quarter of phase — 6 min" : "3 min round")
                        : timerExpired && !submitted
                          ? lateSubmitAllowed ? "⚠️ Late submission allowed by faculty (penalty applies)" : "🔒 Time's up — submission locked"
                          : quarterReady ? "Timer ended" : "Waiting for faculty..."}
                    </span>
                    <span className="timer-display" style={{color: timerExpired ? "var(--red-600)" : timerColor}}>
                      {timerRunning ? formatTimer(timerSeconds) : timerExpired ? "0:00" : quarterReady ? "0:00" : "⏳"}
                    </span>
                  </div>
                </div>

                {/* TIMER EXPIRED BANNER */}
                {timerExpired && !submitted && !lateSubmitAllowed && (
                  <div className="timer-expired-banner">
                    <span className="teb-icon">🔒</span>
                    <div>
                      <strong>Time's up — submissions are locked.</strong>
                      <p>Your faculty can grant a late submission from the Game Control screen. A penalty will apply.</p>
                    </div>
                  </div>
                )}

                {lateSubmitAllowed && !submitted && (
                  <div className="late-allowed-banner">
                    <span>⚠️</span>
                    <div>
                      <strong>Late submission granted by faculty.</strong>
                      <p>A penalty will be applied to your score for this quarter. Submit as soon as possible.</p>
                    </div>
                  </div>
                )}

                <div className="play-grid">
                  <div className="play-input-card">
                    <h3>Set Your Price</h3>
                    <div className="price-setter">
                      <button className="price-btn" onClick={() => setCurrentInput(p => ({...p, price: Math.max(0, +(p.price - PRICE_STEP).toFixed(2))}))}>−</button>
                      <div className="price-display">
                        <span className="pd-currency">₹</span>
                        <span className="pd-value">{currentInput.price.toFixed(2)}</span>
                        <span className="pd-unit">per meal</span>
                      </div>
                      <button className="price-btn" onClick={() => setCurrentInput(p => ({...p, price: +(p.price + PRICE_STEP).toFixed(2)}))}>+</button>
                    </div>
                    {cfg.promo && (
                      <div className="promo-setter">
                        <label>Promotion Expense <span className="promo-max">(max {fmt(maxPromo)})</span></label>
                        <div className="promo-inline"><span>₹</span><input type="number" value={currentInput.promo} min={0} max={maxPromo} onChange={e => setCurrentInput(p => ({...p, promo: Math.min(+e.target.value, maxPromo)}))} /></div>
                      </div>
                    )}
                    <button className={`submit-btn ${submitted ? "submitted" : ""} ${(!quarterReady && !lateSubmitAllowed) ? "locked" : ""} ${timerExpired && !lateSubmitAllowed ? "locked expired" : ""} ${(!commentsValid && (quarterReady || lateSubmitAllowed) && !submitted) ? "needs-comments" : ""}`}
                      onClick={submitQuarter}
                      disabled={submitted || (!quarterReady && !lateSubmitAllowed) || (timerExpired && !lateSubmitAllowed)}>
                      {submitted ? <><Icon d={Icons.check} size={18} /> Submitted!</>
                        : timerExpired && !lateSubmitAllowed ? <><Icon d={Icons.lock} size={18} /> Locked — Time Expired</>
                        : !quarterReady ? <><Icon d={Icons.lock} size={18} /> Waiting for Faculty...</>
                        : lateSubmitAllowed ? <><Icon d={Icons.send} size={18} /> Submit Late (Penalty Applies)</>
                        : <><Icon d={Icons.send} size={18} /> Submit Quarter {currentQuarter}</>}
                    </button>
                    {commentError && <div className="comment-error">{commentError}</div>}
                  </div>

                  <div className="play-preview-card">
                    <h3>Live Preview</h3>
                    <p className="preview-note">Updates as you change your price</p>
                    <div className="aip-display-readonly">
                      <span className="aip-label-ro">Avg Industry Price (set by faculty)</span>
                      <div className="aip-value-ro">₹{aip.toFixed(2)}</div>
                    </div>
                    {preview && (
                      <div className="preview-grid">
                        <div className="pv-item"><span className="pv-l">Retention</span><span className="pv-v">{(preview.retRate*100).toFixed(0)}%</span><span className="pv-hint">{preview.retLabel}</span></div>
                        <div className="pv-item"><span className="pv-l">Retained</span><span className="pv-v">{preview.retained}</span></div>
                        <div className="pv-item"><span className="pv-l">New Customers</span><span className="pv-v">{preview.newCust}</span></div>
                        <div className="pv-item accent"><span className="pv-l">Total Sales</span><span className="pv-v">{preview.totalSales}</span></div>
                        <div className="pv-item"><span className="pv-l">Revenue</span><span className="pv-v">{fmt(preview.revenue)}</span></div>
                        <div className={`pv-item ${preview.profit >= 0 ? "positive" : "negative"}`}><span className="pv-l">Profit</span><span className="pv-v">{fmt(preview.profit)}</span></div>
                      </div>
                    )}
                  </div>
                </div>

                {/* QUARTER COMMENTS — Mandatory, min 100 chars */}
                <div className="quarter-comments-section">
                  <div className="qc-field">
                    <label>Observations for Quarter {currentQuarter} *
                      <span className={`char-count ${obsLen >= 100 ? "ok" : "short"}`}>{obsLen}/100 min</span>
                    </label>
                    <textarea rows={3} value={quarterComments.observations}
                      onChange={e => { setQuarterComments(p => ({...p, observations: e.target.value})); setCommentError(""); }}
                      placeholder="What do you observe about the market this quarter? (minimum 100 characters)"
                      className={obsLen > 0 && obsLen < 100 ? "field-warning" : ""} />
                    {obsLen > 0 && obsLen < 100 && <span className="elaborate-msg">Please elaborate</span>}
                  </div>
                  <div className="qc-field">
                    <label>Strategy for Next Quarter *
                      <span className={`char-count ${stratLen >= 100 ? "ok" : "short"}`}>{stratLen}/100 min</span>
                    </label>
                    <textarea rows={3} value={quarterComments.nextStrategy}
                      onChange={e => { setQuarterComments(p => ({...p, nextStrategy: e.target.value})); setCommentError(""); }}
                      placeholder="What will be your approach next quarter? (minimum 100 characters)"
                      className={stratLen > 0 && stratLen < 100 ? "field-warning" : ""} />
                    {stratLen > 0 && stratLen < 100 && <span className="elaborate-msg">Please elaborate</span>}
                  </div>
                </div>

                {/* Waiting for Faculty overlay */}
                {!quarterReady && !submitted && (
                  <div className="faculty-wait-banner">
                    <div className="fw-icon">⏳</div>
                    <div>
                      <h3>Waiting for Faculty to Start Q{currentQuarter}</h3>
                      <p>{quarters.length === 0
                        ? "The faculty will start the simulation shortly. Your timer will begin automatically when the faculty clicks 'Start Quarter'."
                        : `The faculty needs to process Q${quarters.length} results before Q${currentQuarter} can begin. The timer will start automatically.`}</p>
                    </div>
                  </div>
                )}

                {quarters.length > 0 && (
                  <div className="last-quarter-strip">
                    <h3>Last Quarter Results (Q{quarters.length})</h3>
                    <div className="lqs-grid">
                      <span>Price: {fmtD(quarters[quarters.length-1].price)}</span>
                      <span>Sales: {quarters[quarters.length-1].totalSales}</span>
                      <span>Revenue: {fmt(quarters[quarters.length-1].revenue)}</span>
                      <span className={quarters[quarters.length-1].profit >= 0 ? "profit-pos" : "profit-neg"}>Profit: {fmt(quarters[quarters.length-1].profit)}</span>
                    </div>
                  </div>
                )}
              </>
            ) : (
              <div className="game-over-card">
                <div className="go-icon">🏁</div>
                <h2>Simulation Complete!</h2>
                <div className="go-stats">
                  <div className="go-stat"><span className="go-l">Total Revenue</span><span className="go-v">{fmt(totRev)}</span></div>
                  <div className="go-stat"><span className="go-l">Total Profit</span><span className={`go-v ${totProf >= 0 ? "profit-pos" : "profit-neg"}`}>{fmt(totProf)}</span></div>
                  <div className="go-stat"><span className="go-l">Meals Sold</span><span className="go-v">{totSales.toLocaleString("en-IN")}</span></div>
                </div>
                <div className="conclusions-section">
                  <h3>Your Conclusions</h3>
                  <div className="form-field"><label>Revenue Analysis</label><textarea rows={3} value={conclusions.revenue} onChange={e => setConclusions(p => ({...p, revenue: e.target.value}))} placeholder="Summarise your revenue performance..." /></div>
                  <div className="form-field"><label>Profit Analysis</label><textarea rows={3} value={conclusions.profit} onChange={e => setConclusions(p => ({...p, profit: e.target.value}))} placeholder="Summarise your profit performance..." /></div>
                </div>
              </div>
            )}
          </div>
        )}

        {tab === "strategy" && (
          <div className="page">
            <div className="page-header"><div><h1>Pricing Strategy</h1><p className="page-desc">Document your team's strategic plan</p></div></div>
            <div className="strategy-form">
              {[{key:"generic",label:"Overall Generic Strategy",hint:"Penetration, skimming, competitive parity..."},
                {key:"year1",label:"Year 1 Objectives (Q1–Q4)",hint:"Market entry goals..."},
                {key:"year2",label:"Year 2 Objectives (Q5–Q8)",hint:"Growth and promotion goals..."},
                {key:"year3",label:"Year 3 Objectives (Q9–Q12)",hint:"Recession survival goals..."}
              ].map(({key,label,hint}) => (
                <div className="strat-field" key={key}><label>{label}</label><textarea rows={3} value={strategy[key]} onChange={e => setStrategy(p => ({...p,[key]:e.target.value}))} placeholder={hint} /></div>
              ))}
            </div>
          </div>
        )}

        {tab === "history" && (
          <div className="page">
            <div className="page-header"><div><h1>Performance History</h1><p className="page-desc">{quarters.length} quarters completed</p></div></div>

            {/* Pre-game Strategy */}
            {strategy.generic && (
              <div className="ss-card" style={{marginBottom:"1rem",borderLeft:"4px solid var(--green-700)"}}>
                <h3>Pre-Game Strategy</h3>
                <div className="strategy-history">
                  <div className="sh-item"><span className="sh-label">Generic Strategy</span><p>{strategy.generic}</p></div>
                  {strategy.year1 && <div className="sh-item"><span className="sh-label">Year 1 (Q1–Q4)</span><p>{strategy.year1}</p></div>}
                  {strategy.year2 && <div className="sh-item"><span className="sh-label">Year 2 (Q5–Q8)</span><p>{strategy.year2}</p></div>}
                  {strategy.year3 && <div className="sh-item"><span className="sh-label">Year 3 (Q9–Q12)</span><p>{strategy.year3}</p></div>}
                </div>
              </div>
            )}

            {quarters.length === 0 ? (
              <div className="empty-state"><p>No quarters played yet. Go to the Play tab to start.</p></div>
            ) : (
              <>
                <div className="history-kpis">
                  <div className="hk"><span className="hk-v">{fmt(totRev)}</span><span className="hk-l">Total Revenue</span></div>
                  <div className={`hk ${totProf >= 0 ? "pos" : "neg"}`}><span className="hk-v">{fmt(totProf)}</span><span className="hk-l">Total Profit</span></div>
                  <div className="hk"><span className="hk-v">{totSales.toLocaleString("en-IN")}</span><span className="hk-l">Meals Sold</span></div>
                  <div className="hk"><span className="hk-v">{quarters.length}</span><span className="hk-l">Quarters</span></div>
                </div>
                <div className="history-table-wrap">
                  <table className="history-table">
                    <thead><tr><th>Qtr</th><th>Phase</th><th>Price</th><th>AIP</th><th>Promo</th><th>New</th><th>Ret.</th><th>Sales</th><th>Revenue</th><th>Profit</th></tr></thead>
                    <tbody>
                      {quarters.map((q, i) => (
                        <tr key={i} className={i % 2 === 0 ? "even" : ""}>
                          <td><strong>Q{i+1}</strong></td>
                          <td><span className={`phase-chip p${q.phase}`}>P{q.phase}</span></td>
                          <td>{fmtD(q.price)}</td><td>{fmtD(q.avgComp)}</td><td>{q.promo > 0 ? fmt(q.promo) : "—"}</td>
                          <td>{q.newCust}</td><td>{q.retained}</td><td><strong>{q.totalSales}</strong></td>
                          <td>{fmt(q.revenue)}</td>
                          <td className={q.profit >= 0 ? "profit-pos" : "profit-neg"}><strong>{fmt(q.profit)}</strong></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Quarter-by-quarter observations and strategies */}
                {quarters.some(q => q.observations || q.nextStrategy) && (
                  <div className="ss-card" style={{marginTop:"1rem"}}>
                    <h3>Quarter Observations & Strategies</h3>
                    {quarters.map((q, i) => (
                      (q.observations || q.nextStrategy) ? (
                        <div className="qc-history-row" key={i}>
                          <span className="qc-h-q">Q{i+1}</span>
                          <div className="qc-h-content">
                            {q.observations && <div><span className="qc-h-label">Observations:</span> {q.observations}</div>}
                            {q.nextStrategy && <div><span className="qc-h-label">Next Strategy:</span> {q.nextStrategy}</div>}
                          </div>
                        </div>
                      ) : null
                    ))}
                  </div>
                )}
              </>
            )}
          </div>
        )}

        {tab === "formulas" && (
          <div className="page">
            <div className="page-header"><div><h1>Rules & Formulas</h1><p className="page-desc">Formulas for the current phase{unlockedPhase < 3 ? " — next phase formulas will be revealed when you reach them" : ""}</p></div></div>
            <div className="formulas-content">
              <div className="formula-card"><h3>Initial Conditions</h3><div className="fc-grid"><div>Initial Price: <strong>₹10/meal</strong></div><div>Initial Customers: <strong>100</strong></div><div>Price Increments: <strong>₹0.50</strong></div></div></div>
              <div className="formula-card"><h3>Customer Retention (All Phases)</h3>
                <table className="formula-table"><thead><tr><th>Your Price vs Last AIP</th><th>Retention</th></tr></thead><tbody>
                  <tr><td>More than 20% lower</td><td>30%</td></tr><tr><td>10–20% lower</td><td>15%</td></tr><tr><td>Less than 10% lower</td><td>10%</td></tr><tr><td>≥ AIP</td><td>0%</td></tr>
                </tbody></table></div>
              {[1,2,3].filter(p => p <= unlockedPhase).map(p => { const c = PHASE_CONFIG[p]; return (
                <div className="formula-card" key={p} style={{borderLeft: `4px solid ${c.color}`}}>
                  <div className="fc-phase-tag" style={{background: c.color}}>{c.tag}</div>
                  <h3>{c.name}</h3>
                  <div className="fc-formula">
                    {p === 1 && <code>New Customers = 400 − 40×(Price) + 21×(Avg Competitor Price)</code>}
                    {p === 2 && <><code>New Customers = 400 − 40×(Price) + 21×(Avg Comp Price) + 0.10×(Promo)</code><p className="fc-note">Max promo = 20% of last quarter revenue</p></>}
                    {p === 3 && <><code>New Customers = 400 − 40×(Price) + 21×(Avg Comp Price) + 0.20×(Promo) + 100×(Last Price − This Price)</code><p className="fc-note">VC = ₹5/meal • FC = ₹1000/quarter</p></>}
                  </div>
                  <div className="fc-formulas-sub"><code>Revenue = Total Sales × Own Price</code><code>Profit = Total Sales × (Price − VC) − FC − Promo</code></div>
                </div>
              ); })}
              {unlockedPhase < 3 && (
                <div className="formula-card" style={{opacity:0.5,borderLeft:"4px dashed var(--border)"}}>
                  <h3>🔒 Phase {unlockedPhase + 1} formulas will be revealed after Q{unlockedPhase * 4} is complete</h3>
                </div>
              )}
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

// ─── DEMO DATA ───────────────────────────────────────────────────────────────
function generateDemoTeams() {
  const names = ["Alpha Pricers","Beta Margins","Gamma Foods","Delta Dine","Epsilon Eats","Zeta Meals","Eta Canteen","Theta Bites",
    "Iota Kitchen","Kappa Cuisine","Lambda Lunch","Mu Mess","Nu Nourish","Xi Xpress","Omicron Fare","Pi Plates"];
  const sections = ["A","A","A","A","B","B","B","B","A","A","B","B","A","B","A","B"];
  return names.map((name, i) => ({
    id: `team-${i}`, name, section: `Section ${sections[i]}`,
    members: [{name:`Student ${i*4+1}`,seat:`${i+1}A`,enrol:`2026${String(i*4+1).padStart(3,"0")}`},{name:`Student ${i*4+2}`,seat:`${i+1}B`,enrol:`2026${String(i*4+2).padStart(3,"0")}`},{name:`Student ${i*4+3}`,seat:`${i+1}C`,enrol:`2026${String(i*4+3).padStart(3,"0")}`},{name:`Student ${i*4+4}`,seat:`${i+1}D`,enrol:`2026${String(i*4+4).padStart(3,"0")}`}],
    quarters: [], pendingPrice: null, pendingPromo: null, submitted: Math.random() > 0.4,
  }));
}

// ═══════════════════════════════════════════════════════════════════════════
// ROOT APP
// ═══════════════════════════════════════════════════════════════════════════
export default function App() {
  const [session, setSession] = useState(null);
  if (!session) return <LoginScreen onLogin={setSession} />;
  if (session.role === "faculty") return <FacultyDashboard onLogout={() => setSession(null)} />;
  return <StudentDashboard session={session} onLogout={() => setSession(null)} />;
}
