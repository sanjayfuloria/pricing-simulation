/**
 * Google Apps Script — Pricing Simulation Data Receiver (v3)
 * 
 * Handles: Single team, multi-team, scores, SD penalties, individual/team type
 * 
 * SETUP:
 * 1. Create a new Google Sheet
 * 2. Extensions → Apps Script → paste this code → Save
 * 3. Deploy → New deployment → Web app
 *    Execute as: Me | Who has access: Anyone
 * 4. Copy Web App URL into the simulation app
 * 
 * After code changes: Deploy → Manage deployments → Edit → Version: New version → Deploy
 */

function doPost(e) {
  try {
    var raw = "";
    if (e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    } else if (e.postData && e.postData.contents) {
      raw = e.postData.contents;
    } else {
      throw new Error("No data received");
    }
    
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.teams) {
      saveMultiTeamData(ss, data);
    } else {
      saveSingleTeamData(ss, data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({status:"ok",message:"Data saved successfully!"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    try {
      var debugSheet = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), "Debug",
        ["Timestamp", "Error", "Raw Data"]);
      var rawSnippet = "";
      try { rawSnippet = (e.postData ? e.postData.contents : JSON.stringify(e.parameter || {})).substring(0, 500); } catch(x) {}
      debugSheet.appendRow([new Date().toISOString(), err.toString(), rawSnippet]);
    } catch(x) {}
    
    return ContentService
      .createTextOutput(JSON.stringify({status:"error",message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveSingleTeamData(ss, data) {
  // ── Teams sheet (summary) ──────────────────────────────────
  var teamsSheet = getOrCreateSheet(ss, "Teams", [
    "Timestamp", "Team/Individual Name", "Type", "Section",
    "Member 1", "Member 2", "Member 3", "Member 4",
    "Rank", "Score", "Total Revenue", "Total Profit", "Total Meals Sold",
    "Penalties", "Growth %",
    "Strategy", "Revenue Comments", "Profit Comments"
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
    data.teamName || data.name || "",
    isInd,
    data.section || "",
    members[0] ? members[0].name : "",
    members[1] ? members[1].name : "",
    members[2] ? members[2].name : "",
    members[3] ? members[3].name : "",
    data.rank || "",
    data.score || "",
    totalRevenue,
    totalProfit,
    totalSales,
    data.penalties || 0,
    data.growth || 0,
    strategy.generic || "",
    conclusions.revenueComments || conclusions.revenue || "",
    conclusions.profitComments || conclusions.profit || ""
  ]);

  // ── Quarters sheet (detail) ────────────────────────────────
  var qSheet = getOrCreateSheet(ss, "Quarters", [
    "Timestamp", "Team Name", "Type", "Section",
    "Quarter", "Phase", "Own Price", "Avg Industry Price",
    "Promotion Expense", "New Customers", "Retained Customers",
    "Total Sales", "Revenue", "Profit",
    "SD Penalty", "Z-Score", "Adjusted Profit",
    "Observations", "Next Quarter Strategy"
  ]);

  for (var j = 0; j < quarters.length; j++) {
    var q = quarters[j];
    qSheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.teamName || data.name || "",
      isInd,
      data.section || "",
      q.quarter, q.phase,
      q.price || q.ownPrice || 0,
      q.avgComp || q.avgCompPrice || 0,
      q.promo || q.promoExpense || 0,
      q.newCust || q.newCustomers || 0,
      q.retained || q.retainedCustomers || 0,
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
    team.teamName = team.name || ("Team " + (t+1));
    saveSingleTeamData(ss, team);
  }
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sheet;
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status:"ok",message:"Pricing Simulation API v3 running."}))
    .setMimeType(ContentService.MimeType.JSON);
}
