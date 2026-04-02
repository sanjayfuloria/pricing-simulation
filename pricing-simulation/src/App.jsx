import { useState, useEffect, useCallback, useMemo, useRef } from "react";

/* ═══════════════════════════════════════════════════════════════════════════
   PRICING SIMULATION PLATFORM — B2B Marketing, IBS Hyderabad
   Multi-user competitive pricing game with Faculty & Student portals
   ═══════════════════════════════════════════════════════════════════════════ */

// ─── CONSTANTS ───────────────────────────────────────────────────────────────
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
    nc = 1000 - 40*price + 21*avgComp + 0.20*promo + 100*drop;
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
  try {
    const res = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
    const data = await res.json();
    return { ok: true, message: data.message || "Saved!" };
  } catch (e) { return { ok: false, message: e.message }; }
}

// ═══════════════════════════════════════════════════════════════════════════
// LOGIN SCREEN
// ═══════════════════════════════════════════════════════════════════════════
function LoginScreen({ onLogin }) {
  const [role, setRole] = useState(null);
  const [facultyPass, setFacultyPass] = useState("");
  const [teamName, setTeamName] = useState("");
  const [section, setSection] = useState("");
  const [gameCode, setGameCode] = useState("");
  const [error, setError] = useState("");
  const [members, setMembers] = useState([{name:"",seat:"",enrol:""},{name:"",seat:"",enrol:""},{name:"",seat:"",enrol:""},{name:"",seat:"",enrol:""}]);

  const updateMember = (i, f, v) => { const c=[...members]; c[i]={...c[i],[f]:v}; setMembers(c); };

  const handleFacultyLogin = () => {
    if (facultyPass === "ibs2026") { onLogin({ role: "faculty" }); }
    else { setError("Invalid faculty password"); }
  };

  const handleStudentLogin = () => {
    if (!teamName.trim()) { setError("Team name is required"); return; }
    if (!section.trim()) { setError("Section is required"); return; }
    if (!members[0].name.trim()) { setError("At least one member name required"); return; }
    onLogin({ role: "student", teamName: teamName.trim(), section: section.trim(), members, gameCode: gameCode.trim() });
  };

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
            <button className="role-btn student-btn" onClick={() => { setRole("student"); setError(""); }}>
              <div className="role-icon">👩‍🎓</div>
              <span className="role-title">Student Portal</span>
              <span className="role-desc">Join a game, set prices, compete with teams</span>
            </button>
          </div>
          <div className="login-footer-text">
            <p>Supports up to 200 concurrent teams • Real-time leaderboard • Google Sheets export</p>
          </div>
        </div>
      </div>
    );
  }

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

  return (
    <div className="login-wrapper">
      <div className="login-bg-pattern" />
      <div className="login-container login-form-container student-form">
        <button className="back-btn" onClick={() => setRole(null)}>← Back</button>
        <div className="login-brand compact">
          <div className="brand-icon small">👩‍🎓</div>
          <h2>Team Registration</h2>
        </div>
        <div className="login-form">
          <div className="form-row-2">
            <div className="form-field">
              <label>Team Name *</label>
              <input value={teamName} onChange={e => setTeamName(e.target.value)} placeholder="e.g. Alpha Pricers" autoFocus />
            </div>
            <div className="form-field">
              <label>Section *</label>
              <input value={section} onChange={e => setSection(e.target.value)} placeholder="e.g. Section A" />
            </div>
          </div>
          <div className="form-field">
            <label>Game Code <span className="optional">(if provided by faculty)</span></label>
            <input value={gameCode} onChange={e => setGameCode(e.target.value)} placeholder="e.g. GAME-2026" />
          </div>
          <div className="members-section">
            <label className="section-label">Team Members</label>
            {members.map((m, i) => (
              <div className="member-row" key={i}>
                <span className="member-num">{i+1}</span>
                <input placeholder="Name" value={m.name} onChange={e => updateMember(i, "name", e.target.value)} />
                <input placeholder="Seat" value={m.seat} onChange={e => updateMember(i, "seat", e.target.value)} className="short" />
                <input placeholder="Enrolment #" value={m.enrol} onChange={e => updateMember(i, "enrol", e.target.value)} />
              </div>
            ))}
          </div>
          {error && <div className="login-error">{error}</div>}
          <button className="login-submit student-submit" onClick={handleStudentLogin}>Join Simulation</button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// FACULTY DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════
function FacultyDashboard({ onLogout }) {
  const [tab, setTab] = useState("control");
  const [gameState, setGameState] = useState({
    currentQuarter: 1, status: "waiting", aipHistory: [],
    teams: generateDemoTeams(), sheetsUrl: "", gameCode: "GAME-2026"
  });
  const [aipInput, setAipInput] = useState(INIT_PRICE);
  const [pushStatus, setPushStatus] = useState(null);

  const currentPhase = getPhase(gameState.currentQuarter);
  const cfg = PHASE_CONFIG[currentPhase];

  function advanceQuarter() {
    setGameState(prev => {
      const newAipHistory = [...prev.aipHistory, { quarter: prev.currentQuarter, aip: aipInput }];
      const updatedTeams = prev.teams.map(team => {
        const prevData = team.quarters.length > 0 ? team.quarters[team.quarters.length - 1] : null;
        const phase = getPhase(prev.currentQuarter);
        const result = simulateQuarter(team.pendingPrice || INIT_PRICE, aipInput, phase, team.pendingPromo || 0, prevData);
        return { ...team, quarters: [...team.quarters, result], pendingPrice: null, pendingPromo: null, submitted: false };
      });
      return { ...prev, currentQuarter: prev.currentQuarter + 1, aipHistory: newAipHistory,
               teams: updatedTeams, status: prev.currentQuarter >= 12 ? "finished" : "active" };
    });
  }

  const sortedTeams = useMemo(() => {
    return [...gameState.teams].sort((a, b) => {
      const profA = a.quarters.reduce((s, q) => s + q.profit, 0);
      const profB = b.quarters.reduce((s, q) => s + q.profit, 0);
      return profB - profA;
    });
  }, [gameState.teams]);

  const submittedCount = gameState.teams.filter(t => t.submitted).length;

  const tabs = [
    { id: "control", label: "Game Control", icon: Icons.settings },
    { id: "leaderboard", label: "Leaderboard", icon: Icons.trophy },
    { id: "teams", label: "All Teams", icon: Icons.users },
    { id: "analytics", label: "Analytics", icon: Icons.chart },
    { id: "sheets", label: "Google Sheets", icon: Icons.grid },
  ];

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
            <button key={t.id} className={`nav-item ${tab === t.id ? "active" : ""}`} onClick={() => setTab(t.id)}>
              <Icon d={t.icon} size={18} />
              <span>{t.label}</span>
            </button>
          ))}
        </div>
        <div className="sidebar-footer">
          <div className="game-status-badge">
            <span className={`status-dot ${gameState.status}`} />
            Q{Math.min(gameState.currentQuarter, 12)} of 12 • Phase {currentPhase}
          </div>
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
                <div className="hb" style={{borderColor: cfg.color}}>
                  <span className="hb-label">Phase</span>
                  <span className="hb-value" style={{color: cfg.color}}>{currentPhase}</span>
                </div>
                <div className="hb">
                  <span className="hb-label">Quarter</span>
                  <span className="hb-value">{Math.min(gameState.currentQuarter, 12)}</span>
                </div>
                <div className="hb">
                  <span className="hb-label">Teams</span>
                  <span className="hb-value">{gameState.teams.length}</span>
                </div>
              </div>
            </div>

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

            {gameState.currentQuarter <= 12 ? (
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
                      <div className="meter-fill" style={{width: `${(submittedCount / gameState.teams.length)*100}%`}} />
                    </div>
                    <span className="meter-text">{submittedCount} / {gameState.teams.length} teams submitted</span>
                  </div>
                </div>

                <button className="advance-btn" onClick={advanceQuarter}>
                  <Icon d={Icons.play} size={20} />
                  Process Quarter {gameState.currentQuarter} & Advance
                </button>
              </div>
            ) : (
              <div className="control-card finished-card">
                <div className="finished-icon">🏁</div>
                <h2>Simulation Complete</h2>
                <p>All 12 quarters played. Check Leaderboard and Analytics for results.</p>
              </div>
            )}

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
          </div>
        )}

        {tab === "leaderboard" && (
          <div className="page">
            <div className="page-header">
              <div><h1>Leaderboard</h1><p className="page-desc">Real-time rankings by cumulative profit</p></div>
            </div>
            <div className="leaderboard">
              <div className="lb-header-row">
                <span className="lb-rank">#</span>
                <span className="lb-team">Team</span>
                <span className="lb-section">Section</span>
                <span className="lb-stat">Revenue</span>
                <span className="lb-stat">Profit</span>
                <span className="lb-stat">Meals Sold</span>
                <span className="lb-stat">Avg Price</span>
              </div>
              {sortedTeams.map((team, i) => {
                const totRev = team.quarters.reduce((s,q) => s+q.revenue, 0);
                const totProf = team.quarters.reduce((s,q) => s+q.profit, 0);
                const totSales = team.quarters.reduce((s,q) => s+q.totalSales, 0);
                const avgP = team.quarters.length > 0 ? team.quarters.reduce((s,q) => s+q.price, 0)/team.quarters.length : 0;
                return (
                  <div className={`lb-row ${i < 3 ? "lb-top" : ""}`} key={team.id}>
                    <span className={`lb-rank rank-${i+1}`}>{i < 3 ? ["🥇","🥈","🥉"][i] : i+1}</span>
                    <span className="lb-team">{team.name}</span>
                    <span className="lb-section">{team.section}</span>
                    <span className="lb-stat">{fmt(totRev)}</span>
                    <span className={`lb-stat ${totProf >= 0 ? "profit-pos" : "profit-neg"}`}>{fmt(totProf)}</span>
                    <span className="lb-stat">{totSales.toLocaleString("en-IN")}</span>
                    <span className="lb-stat">{fmtD(avgP)}</span>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab === "teams" && (
          <div className="page">
            <div className="page-header">
              <div><h1>All Teams ({gameState.teams.length})</h1><p className="page-desc">Detailed view of every registered team</p></div>
            </div>
            <div className="teams-grid">
              {gameState.teams.map(team => {
                const totProf = team.quarters.reduce((s,q) => s+q.profit, 0);
                const totRev = team.quarters.reduce((s,q) => s+q.revenue, 0);
                return (
                  <div className="team-tile" key={team.id}>
                    <div className="tt-header">
                      <h3>{team.name}</h3>
                      <span className="tt-section">{team.section}</span>
                    </div>
                    <div className="tt-members">{team.members.filter(m => m.name).map(m => m.name).join(", ") || "—"}</div>
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
              <div><h1>Google Sheets Export</h1><p className="page-desc">Push all team data for grading</p></div>
            </div>
            <div className="sheets-setup">
              <div className="ss-card">
                <h3>Web App URL</h3>
                <p>Deploy the provided Google Apps Script and paste the URL below.</p>
                <div className="ss-input-row">
                  <input type="url" placeholder="https://script.google.com/macros/s/.../exec"
                    value={gameState.sheetsUrl} onChange={e => setGameState(p => ({...p, sheetsUrl: e.target.value}))} />
                  <button className="btn-push" onClick={async () => {
                    setPushStatus(null);
                    const payload = { teams: gameState.teams.map(t => ({
                      name: t.name, section: t.section, members: t.members,
                      quarters: t.quarters.map((q,i) => ({ quarter: i+1, ...q }))
                    })), aipHistory: gameState.aipHistory, timestamp: new Date().toISOString() };
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
  const [tab, setTab] = useState("play");
  const [quarters, setQuarters] = useState([]);
  const [currentInput, setCurrentInput] = useState({ price: INIT_PRICE, promo: 0 });
  const [strategy, setStrategy] = useState({ generic: "", year1: "", year2: "", year3: "" });
  const [aip, setAip] = useState(INIT_PRICE);
  const [submitted, setSubmitted] = useState(false);
  const [conclusions, setConclusions] = useState({ revenue: "", profit: "" });

  const currentQuarter = quarters.length + 1;
  const phase = getPhase(Math.min(currentQuarter, 12));
  const cfg = PHASE_CONFIG[phase];
  const prevData = quarters.length > 0 ? quarters[quarters.length - 1] : null;
  const lastRev = prevData ? prevData.revenue : INIT_CUSTOMERS * INIT_PRICE;
  const maxPromo = cfg.promo ? Math.floor(lastRev * 0.2) : 0;

  const preview = currentQuarter <= 12 ? simulateQuarter(currentInput.price, aip, phase, Math.min(currentInput.promo, maxPromo), prevData) : null;

  const submitQuarter = () => {
    const result = simulateQuarter(currentInput.price, aip, phase, Math.min(currentInput.promo, maxPromo), prevData);
    setQuarters(prev => [...prev, result]);
    setSubmitted(true);
    setTimeout(() => { setSubmitted(false); setCurrentInput({ price: currentInput.price, promo: 0 }); }, 1500);
  };

  const totRev = quarters.reduce((s,q) => s+q.revenue, 0);
  const totProf = quarters.reduce((s,q) => s+q.profit, 0);
  const totSales = quarters.reduce((s,q) => s+q.totalSales, 0);

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
            <div className="sb-role">{session.section}</div>
          </div>
        </div>
        <div className="sidebar-nav">
          {tabs.map(t => (
            <button key={t.id} className={`nav-item ${tab === t.id ? "active" : ""}`} onClick={() => setTab(t.id)}>
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
          <button className="nav-item logout-item" onClick={onLogout}>
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
                    <div className="aip-setter">
                      <label>Average Industry Price (announced by faculty)</label>
                      <div className="aip-inline"><span>₹</span><input type="number" value={aip} onChange={e => setAip(+e.target.value)} step={PRICE_STEP} min={0} /></div>
                    </div>
                    {cfg.promo && (
                      <div className="promo-setter">
                        <label>Promotion Expense <span className="promo-max">(max {fmt(maxPromo)})</span></label>
                        <div className="promo-inline"><span>₹</span><input type="number" value={currentInput.promo} min={0} max={maxPromo} onChange={e => setCurrentInput(p => ({...p, promo: Math.min(+e.target.value, maxPromo)}))} /></div>
                      </div>
                    )}
                    <button className={`submit-btn ${submitted ? "submitted" : ""}`} onClick={submitQuarter} disabled={submitted}>
                      {submitted ? <><Icon d={Icons.check} size={18} /> Submitted!</> : <><Icon d={Icons.send} size={18} /> Submit Quarter {currentQuarter}</>}
                    </button>
                  </div>

                  <div className="play-preview-card">
                    <h3>Live Preview</h3>
                    <p className="preview-note">Updates as you change your price</p>
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
              </>
            )}
          </div>
        )}

        {tab === "formulas" && (
          <div className="page">
            <div className="page-header"><div><h1>Rules & Formulas</h1><p className="page-desc">Complete reference for all three phases</p></div></div>
            <div className="formulas-content">
              <div className="formula-card"><h3>Initial Conditions</h3><div className="fc-grid"><div>Initial Price: <strong>₹10/meal</strong></div><div>Initial Customers: <strong>100</strong></div><div>Price Increments: <strong>₹0.50</strong></div></div></div>
              <div className="formula-card"><h3>Customer Retention (All Phases)</h3>
                <table className="formula-table"><thead><tr><th>Your Price vs Last AIP</th><th>Retention</th></tr></thead><tbody>
                  <tr><td>More than 20% lower</td><td>30%</td></tr><tr><td>10–20% lower</td><td>15%</td></tr><tr><td>Less than 10% lower</td><td>10%</td></tr><tr><td>≥ AIP</td><td>0%</td></tr>
                </tbody></table></div>
              {[1,2,3].map(p => { const c = PHASE_CONFIG[p]; return (
                <div className="formula-card" key={p} style={{borderLeft: `4px solid ${c.color}`}}>
                  <div className="fc-phase-tag" style={{background: c.color}}>{c.tag}</div>
                  <h3>{c.name}</h3>
                  <div className="fc-formula">
                    {p === 1 && <code>New Customers = 400 − 40×(Price) + 21×(Avg Competitor Price)</code>}
                    {p === 2 && <><code>New Customers = 400 − 40×(Price) + 21×(Avg Comp Price) + 0.10×(Promo)</code><p className="fc-note">Max promo = 20% of last quarter revenue</p></>}
                    {p === 3 && <><code>New Customers = 1000 − 40×(Price) + 21×(Avg Comp Price) + 0.20×(Promo) + 100×(Last Price − This Price)</code><p className="fc-note">VC = ₹5/meal • FC = ₹1000/quarter</p></>}
                  </div>
                  <div className="fc-formulas-sub"><code>Revenue = Total Sales × Own Price</code><code>Profit = Total Sales × (Price − VC) − FC − Promo</code></div>
                </div>
              ); })}
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
