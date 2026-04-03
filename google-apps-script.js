/**
 * Google Apps Script — Pricing Simulation Data Receiver (v4)
 * 
 * Creates 5 sheets:
 *   "Games"       — one row per game push with metadata & top scores
 *   "Teams"       — one row per team/individual with scores, strategy, members
 *   "Quarters"    — one row per quarter per team with full detail
 *   "AIP History" — one row per quarter with AIP values
 *   "Debug"       — error logging (auto-created on errors)
 * 
 * SETUP:
 * 1. Create a new Google Sheet (name it anything)
 * 2. Extensions → Apps Script → delete default code → paste this → Save
 * 3. Deploy → New deployment → Web app
 *    Execute as: Me | Who has access: Anyone
 * 4. Authorize when prompted
 * 5. Copy the Web App URL into the simulation app
 * 
 * IMPORTANT: After ANY code change:
 *   Deploy → Manage deployments → pencil icon → Version: "New version" → Deploy
 */

function doPost(e) {
  try {
    var raw = "";
    if (e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    } else if (e.postData && e.postData.contents) {
      raw = e.postData.contents;
    } else {
      throw new Error("No data received. Keys: " + Object.keys(e || {}).join(","));
    }
    
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.teams) {
      saveGameSummary(ss, data);
      saveAIPHistory(ss, data);
      saveMultiTeamData(ss, data);
    } else {
      saveSingleTeamData(ss, data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({status:"ok", message:"Data saved successfully!"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    logError(err, e);
    return ContentService
      .createTextOutput(JSON.stringify({status:"error", message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveGameSummary(ss, data) {
  var sheet = getOrCreateSheet(ss, "Games", [
    "Timestamp", "Game ID", "Total Teams", "Total Individuals",
    "Quarters Played", "SD Penalty Enabled",
    "Top Team", "Top Score", "Top Profit",
    "Avg Revenue", "Avg Profit"
  ]);

  var teams = data.teams || [];
  var teamCount = teams.filter(function(t) { return !t.isIndividual; }).length;
  var indCount = teams.filter(function(t) { return t.isIndividual; }).length;
  var maxQ = 0;
  teams.forEach(function(t) {
    var ql = (t.quarters || []).length;
    if (ql > maxQ) maxQ = ql;
  });

  var topTeam = "", topScore = 0, topProfit = 0;
  var totalRev = 0, totalProf = 0;
  teams.forEach(function(t) {
    if ((t.score || 0) > topScore) {
      topScore = t.score || 0;
      topTeam = t.name || "";
      topProfit = t.totProfit || 0;
    }
    totalRev += (t.totRevenue || 0);
    totalProf += (t.totProfit || 0);
  });
  var avgRev = teams.length > 0 ? Math.round(totalRev / teams.length) : 0;
  var avgProf = teams.length > 0 ? Math.round(totalProf / teams.length) : 0;

  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.gameId || "",
    teamCount, indCount, maxQ,
    data.sdPenaltyEnabled ? "Yes" : "No",
    topTeam, topScore, topProfit, avgRev, avgProf
  ]);
}

function saveAIPHistory(ss, data) {
  var aipHistory = data.aipHistory || [];
  if (aipHistory.length === 0) return;

  var sheet = getOrCreateSheet(ss, "AIP History", [
    "Timestamp", "Game ID", "Quarter", "Phase", "AIP"
  ]);

  aipHistory.forEach(function(h) {
    var phase = h.quarter <= 4 ? 1 : h.quarter <= 8 ? 2 : 3;
    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.gameId || "",
      h.quarter, phase, h.aip
    ]);
  });
}

function saveSingleTeamData(ss, data) {
  var teamsSheet = getOrCreateSheet(ss, "Teams", [
    "Timestamp", "Game ID",
    "Team/Individual Name", "Selected Team ID", "Type", "Section",
    "Member 1", "Member 2", "Member 3", "Member 4", "Member 5",
    "Rank", "Score",
    "Total Revenue", "Total Profit", "Total Meals Sold",
    "Penalties", "Growth %",
    "Generic Strategy", "Year 1 Strategy", "Year 2 Strategy", "Year 3 Strategy",
    "Revenue Comments", "Profit Comments"
  ]);

  var members = data.members || [];
  var strategy = data.strategy || {};
  var quarters = data.quarters || [];
  var conclusions = data.conclusions || {};
  var isInd = data.isIndividual ? "Individual" : "Team";

  var totalRevenue = data.totRevenue || 0;
  var totalProfit = data.totProfit || 0;
  var totalSales = data.totSales || 0;
  if (!totalRevenue && quarters.length) {
    for (var i = 0; i < quarters.length; i++) {
      totalRevenue += (quarters[i].revenue || 0);
      totalProfit += (quarters[i].profit || 0);
      totalSales += (quarters[i].totalSales || 0);
    }
  }

  teamsSheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.gameId || "",
    data.teamName || data.name || "",
    data.selectedTeamId || "",
    isInd,
    data.section || "",
    getMemberName(members, 0),
    getMemberName(members, 1),
    getMemberName(members, 2),
    getMemberName(members, 3),
    getMemberName(members, 4),
    data.rank || "",
    data.score || "",
    totalRevenue, totalProfit, totalSales,
    data.penalties || 0,
    data.growth || 0,
    strategy.generic || "",
    strategy.year1 || "",
    strategy.year2 || "",
    strategy.year3 || "",
    conclusions.revenueComments || conclusions.revenue || "",
    conclusions.profitComments || conclusions.profit || ""
  ]);

  var qSheet = getOrCreateSheet(ss, "Quarters", [
    "Timestamp", "Game ID",
    "Team Name", "Type", "Section",
    "Quarter", "Phase",
    "Own Price", "Avg Industry Price",
    "Promotion Expense",
    "Retention Rate", "Retained Customers",
    "New Customers", "Total Sales",
    "Revenue", "Profit",
    "SD Penalty", "Z-Score", "Adjusted Profit",
    "Observations", "Next Quarter Strategy"
  ]);

  for (var j = 0; j < quarters.length; j++) {
    var q = quarters[j];
    qSheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.gameId || "",
      data.teamName || data.name || "",
      isInd, data.section || "",
      q.quarter, q.phase,
      q.price || q.ownPrice || 0,
      q.avgComp || q.avgCompPrice || 0,
      q.promo || q.promoExpense || 0,
      q.retRate || 0,
      q.retained || q.retainedCustomers || 0,
      q.newCust || q.newCustomers || 0,
      q.totalSales || 0,
      q.revenue || 0,
      q.profit || 0,
      q.sdPenalty || 0,
      q.zScore || "",
      q.adjustedProfit || q.profit || 0,
      q.observations || "",
      q.nextStrategy || ""
    ]);
  }
}

function saveMultiTeamData(ss, data) {
  var teams = data.teams || [];
  for (var t = 0; t < teams.length; t++) {
    var team = teams[t];
    team.timestamp = data.timestamp || new Date().toISOString();
    team.gameId = data.gameId || "";
    team.teamName = team.name || ("Team " + (t + 1));
    saveSingleTeamData(ss, team);
  }
}

function getMemberName(members, index) {
  if (!members || index >= members.length) return "";
  var m = members[index];
  if (!m) return "";
  var name = m.name || "";
  var enrol = m.enrol || m.enrolment || "";
  return enrol ? name + " (" + enrol + ")" : name;
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    try { sheet.autoResizeColumns(1, headers.length); } catch(e) {}
  }
  return sheet;
}

function logError(err, e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = getOrCreateSheet(ss, "Debug", ["Timestamp", "Error", "Raw Data (first 500 chars)"]);
    var rawSnippet = "";
    try {
      if (e && e.postData) rawSnippet = e.postData.contents.substring(0, 500);
      else if (e && e.parameter) rawSnippet = JSON.stringify(e.parameter).substring(0, 500);
    } catch(x) { rawSnippet = "Could not extract raw data"; }
    debugSheet.appendRow([new Date().toISOString(), err.toString(), rawSnippet]);
  } catch(x) {}
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: "ok",
      message: "Pricing Simulation API v4 running.",
      version: "4.0",
      sheets: ["Games", "Teams", "Quarters", "AIP History", "Debug"]
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
