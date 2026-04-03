/**
 * Google Apps Script — Pricing Simulation Data Receiver
 * 
 * SETUP:
 * 1. Create a new Google Sheet
 * 2. Extensions → Apps Script → paste this code → Save
 * 3. Deploy → New deployment → Web app
 *    Execute as: Me | Who has access: Anyone
 * 4. Copy Web App URL into the simulation app
 * 
 * IMPORTANT: After code changes, create a NEW deployment version:
 *   Deploy → Manage deployments → Edit → Version: New version → Deploy
 */

function doPost(e) {
  try {
    var raw = "";
    
    // Form submissions send data in e.parameter
    if (e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    }
    // Raw JSON POST sends data in e.postData
    else if (e.postData && e.postData.contents) {
      raw = e.postData.contents;
    }
    else {
      throw new Error("No data received. Keys: " + Object.keys(e).join(","));
    }
    
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Handle both single-team and multi-team payloads
    if (data.teams) {
      // Multi-team payload from instructor view
      saveMultiTeamData(ss, data);
    } else {
      // Single team payload
      saveSingleTeamData(ss, data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({status:"ok",message:"Data saved successfully!"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    // Log error to a Debug sheet for troubleshooting
    try {
      var debugSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Debug");
      if (!debugSheet) {
        debugSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Debug");
        debugSheet.appendRow(["Timestamp", "Error", "Raw Data"]);
      }
      var rawData = "";
      try { rawData = e.postData ? e.postData.contents.substring(0, 500) : "no postData"; } catch(x) {}
      try { rawData = rawData || (e.parameter ? JSON.stringify(e.parameter).substring(0, 500) : "no parameter"); } catch(x) {}
      debugSheet.appendRow([new Date().toISOString(), err.toString(), rawData]);
    } catch(x) {}
    
    return ContentService
      .createTextOutput(JSON.stringify({status:"error",message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveSingleTeamData(ss, data) {
  var teamsSheet = getOrCreateSheet(ss, "Teams", [
    "Timestamp","Team Name","Section",
    "Member 1","Member 2","Member 3","Member 4",
    "Generic Strategy","Year 1 Objectives","Year 2 Objectives","Year 3 Objectives",
    "Total Revenue","Total Profit","Total Meals Sold",
    "Revenue Comments","Profit Comments"
  ]);

  var members = data.members || [];
  var strategy = data.strategy || {};
  var quarters = data.quarters || [];
  var conclusions = data.conclusions || {};

  var totalRevenue = 0, totalProfit = 0, totalSales = 0;
  for (var i = 0; i < quarters.length; i++) {
    totalRevenue += (quarters[i].revenue || 0);
    totalProfit += (quarters[i].profit || 0);
    totalSales += (quarters[i].totalSales || 0);
  }

  teamsSheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.teamName || "", data.section || "",
    members[0] ? members[0].name : "", members[1] ? members[1].name : "",
    members[2] ? members[2].name : "", members[3] ? members[3].name : "",
    strategy.generic || "", strategy.year1 || "",
    strategy.year2 || "", strategy.year3 || "",
    totalRevenue, totalProfit, totalSales,
    conclusions.revenueComments || "", conclusions.profitComments || ""
  ]);

  var qSheet = getOrCreateSheet(ss, "Quarters", [
    "Timestamp","Team Name","Section","Quarter","Phase",
    "Own Price","Avg Industry Price","Promotion Expense",
    "New Customers","Retained Customers","Total Sales",
    "Revenue","Profit","Observations","Next Quarter Strategy"
  ]);

  for (var j = 0; j < quarters.length; j++) {
    var q = quarters[j];
    qSheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.teamName || "", data.section || "",
      q.quarter, q.phase, q.ownPrice, q.avgCompPrice,
      q.promoExpense || 0, q.newCustomers || 0, q.retainedCustomers || 0,
      q.totalSales || 0, q.revenue || 0, q.profit || 0,
      q.observations || "", q.nextStrategy || ""
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
    .createTextOutput(JSON.stringify({status:"ok",message:"Pricing Simulation API is running."}))
    .setMimeType(ContentService.MimeType.JSON);
}
